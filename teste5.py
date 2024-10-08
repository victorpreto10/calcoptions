import streamlit as st
import yfinance as yf
import numpy as np
import pandas as pd
from scipy.stats import norm
from scipy.optimize import bisect
import math as m
from decimal import Decimal, getcontext
import math
import tempfile
import shutil
import os
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
from datetime import datetime, timedelta
import subprocess
from io import StringIO, BytesIO
import requests
import re
import time
import scipy.stats as ss
from arch import arch_model
import streamlit as st
import matplotlib.pyplot as plt
import io


getcontext().prec = 28  # Definir precisão para operações Decimal

# Funções Gerais


def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=1)
    if 'Price' in df.columns and df['Price'].dtype == 'object':
        df['Price'] = df['Price'].str.replace(',', '').astype(float)
    df['Operation'] = df['Qtde'].apply(lambda x: 'Buy' if x > 0 else 'Sell')
    if 'Price' in df.columns:
        df['Total Value'] = df['Price'] * df['Qtde']
        grouped_df = df.groupby(['Operation', 'Ticker Bloomberg']).apply(
            lambda x: pd.Series({
                'Qtde': x['Qtde'].sum(),
                'Weighted Price': (x['Total Value']).sum() / x['Qtde'].sum()
            })).reset_index()
    else:
        grouped_df = df.groupby(['Operation', 'Ticker Bloomberg']).agg({'Qtde': 'sum'}).reset_index()
    return grouped_df


def download_data(asset, start, end, max_retries=5):
    for i in range(max_retries):
        try:
            data = yf.download(asset, start=start, end=end)
            return data
        except Exception as e:
            st.warning(f"Tentativa {i+1} falhou. Tentando novamente...")
            time.sleep(2)
    st.error("Falha ao baixar dados após múltiplas tentativas.")
    return None

def calculate_weighted_average(df):
    df['Weighted_Price'] = df['Price'] * df['Quantity']
    result_df = df.groupby(['Action', 'Ticker', 'Date', 'Option Type', 'Strike Price']).agg(
        Total_Quantity=('Quantity', 'sum'),
        Total_Weighted_Price=('Weighted_Price', 'sum')
    ).reset_index()
    result_df['Average_Price'] = result_df['Total_Weighted_Price'] / result_df['Total_Quantity']
    return result_df.drop(columns=['Total_Weighted_Price'])

def parse_number_input(input_str):
    input_str = input_str.lower().strip()
    if input_str.endswith('k'):
        return float(input_str[:-1]) * 1000
    elif input_str.endswith('m'):
        return float(input_str[:-1]) * 1000000
    elif input_str.replace('.', '', 1).isdigit():
        return float(input_str)
    else:
        raise ValueError("Invalid input: Please enter a valid number with 'k' for thousands or 'm' for millions if needed.")

def get_real_time_price(ticker, api_key):
    url = f'https://finnhub.io/api/v1/quote?symbol={ticker}&token={api_key}'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return data['c']  # Preço atual
    else:
        st.error("Falha ao buscar dados na Finnhub.")
        return None

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=1)
    if 'Price' in df.columns and df['Price'].dtype == 'object':
        df['Price'] = df['Price'].str.replace(',', '').astype(float)
    df['Operation'] = df['Qtde'].apply(lambda x: 'Buy' if x > 0 else 'Sell')
    if 'Price' in df.columns:
        df['Total Value'] = df['Price'] * df['Qtde']
        grouped_df = df.groupby(['Operation', 'Ticker Bloomberg']).apply(
            lambda x: pd.Series({
                'Qtde': x['Qtde'].sum(),
                'Weighted Price': (x['Total Value']).sum() / x['Qtde'].sum()
            })).reset_index()
    else:
        grouped_df = df.groupby(['Operation', 'Ticker Bloomberg']).agg({'Qtde': 'sum'}).reset_index()
    return grouped_df

# Funções de Processamento de Dados

def process_data(start_date, end_date):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Caixa de Entrada
    sentbox = outlook.GetDefaultFolder(5)  # Itens Enviados

    consolidated_df = pd.DataFrame()

    for folder in [inbox, sentbox]:
        consolidated_df = pd.concat([consolidated_df, process_folder(folder, start_date, end_date)], ignore_index=True)

    if not consolidated_df.empty:
        consolidated_df.iloc[:, 5] = pd.to_numeric(consolidated_df.iloc[:, 5], errors='coerce')
        consolidated_df.iloc[:, 6] = pd.to_numeric(consolidated_df.iloc[:, 6], errors='coerce')
        consolidated_df.dropna(subset=[consolidated_df.columns[5], consolidated_df.columns[6]], inplace=True)

    return consolidated_df

def process_folder(folder, start_date, end_date):
    df_list = []
    for single_date in (start_date + timedelta(days=n) for n in range((end_date - start_date).days + 1)):
        formatted_date1 = single_date.strftime("Today Commissions %d%B%y")
        formatted_date2 = single_date.strftime("Today Commissions %d-%b")
        subject_filter = f"[Subject] = '{formatted_date1}' OR [Subject] = '{formatted_date2}'"
        daily_messages = folder.Items.Restrict(subject_filter)
        for message in daily_messages:
            try:
                html_body = message.HTMLBody
                html_stream = StringIO(html_body)
                tables = pd.read_html(html_stream)
                if tables:
                    df = tables[0]
                    df = df[~df.apply(lambda row: row.astype(str).str.contains('Executed Quantity|Time').any(), axis=1)]
                    df_list.append(df)
            except Exception as e:
                print(f"Erro ao processar o e-mail de {single_date.strftime('%d-%b')}: {e}")
    return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

def parse_data(data):
    data = StringIO(data)
    df = pd.read_csv(data, sep='\t', engine='python')
    df.columns = df.columns.str.strip()
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def calculate_average_price(df):
    if 'Maturity' not in df.columns:
        st.error("Column 'Maturity' is missing from the input. Please check your data and try again.")
        return pd.DataFrame()
    df['Maturity'] = pd.to_datetime(df['Maturity'], format='%m/%d/%Y', errors='coerce')
    if df['Maturity'].isna().any():
        st.error("Some dates in 'Maturity' column could not be parsed. Please check the format.")
        return pd.DataFrame()
    grouped_df = df.groupby(['Symbol', 'Side', 'Strike', 'CALL / PUT', 'Maturity']).agg(
        Quantity_Total=('Quantity', 'sum'),
        Average_Execution_Price=('Execution Price', 'mean'),
        Total_Commission=('Commission', 'sum')
    ).reset_index()
    return grouped_df

def sum_quantities_by_operation(df):
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    quantities_sum = df.groupby('Type')['Quantity'].sum().to_dict()
    return quantities_sum

# Funções de Volatilidade e Preços de Opções

def calcular_volatilidade_historica(ticker, periodo):
    ativo = yf.Ticker(ticker)
    dados = ativo.history(period=periodo)
    retornos_log = np.log(dados['Close'] / dados['Close'].shift(1))
    volatilidade = retornos_log.std() * np.sqrt(252)  # Anualizando
    return volatilidade

def calcular_gregas_bs(preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos, volatilidade, tipo_opcao):
    d1 = (m.log(preco_subjacente / preco_exercicio) + (taxa_juros - dividendos + 0.5 * volatilidade ** 2) * tempo) / (
            volatilidade * m.sqrt(tempo))
    d2 = d1 - volatilidade * m.sqrt(tempo)
    delta = norm.cdf(d1) if tipo_opcao == 'Call' else norm.cdf(d1) - 1
    gamma = norm.pdf(d1) / (preco_subjacente * volatilidade * m.sqrt(tempo))
    vega = preco_subjacente * norm.pdf(d1) * m.sqrt(tempo) * 0.01
    theta = (-preco_subjacente * norm.pdf(d1) * volatilidade / (
                2 * m.sqrt(tempo)) - taxa_juros * preco_exercicio * m.exp(
        -taxa_juros * tempo) * norm.cdf(d2)) / 365
    rho = preco_exercicio * tempo * m.exp(-taxa_juros * tempo) * norm.cdf(
        d2) * 0.01 if tipo_opcao == 'Call' else -preco_exercicio * tempo * m.exp(-taxa_juros * tempo) * norm.cdf(
        -d2) * 0.01
    return {'Delta': delta, 'Gamma': gamma, 'Vega': vega, 'Theta': theta, 'Rho': rho}

def calcular_opcao(tipo_opcao, metodo_solucao, preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos,
                   volatilidade, num_simulacoes=10000, num_periodos=100):
    try:
        dt = tempo / num_periodos
        discount_factor = np.exp(-taxa_juros * dt)
        z = np.random.normal(size=(num_simulacoes, num_periodos))
        S = np.zeros_like(z)
        S[:, 0] = preco_subjacente
        for t in range(1, num_periodos):
            S[:, t] = S[:, t - 1] * np.exp(
                (taxa_juros - dividendos - 0.5 * volatilidade ** 2) * dt + volatilidade * np.sqrt(dt) * z[:, t])
        if tipo_opcao == 'Europeia':
            if metodo_solucao == 'Black-Scholes':
                d1 = (m.log(preco_subjacente / preco_exercicio) + (
                            taxa_juros - dividendos + 0.5 * volatilidade ** 2) * tempo) / (
                                 volatilidade * np.sqrt(tempo))
                d2 = d1 - volatilidade * np.sqrt(tempo)
                preco_opcao_compra = preco_subjacente * np.exp(-dividendos * tempo) * norm.cdf(
                    d1) - preco_exercicio * np.exp(-taxa_juros * tempo) * norm.cdf(d2)
                preco_opcao_venda = preco_exercicio * np.exp(-taxa_juros * tempo) * norm.cdf(
                    -d2) - preco_subjacente * np.exp(-dividendos * tempo) * norm.cdf(-d1)
                return preco_opcao_compra, preco_opcao_venda
            elif metodo_solucao == 'Monte Carlo':
                preco_opcao_compra = np.mean(np.maximum(S[:, -1] - preco_exercicio, 0) * discount_factor)
                preco_opcao_venda = np.mean(np.maximum(preco_exercicio - S[:, -1], 0) * discount_factor)
                return preco_opcao_compra, preco_opcao_venda
            elif metodo_solucao == 'Binomial':
                num_steps = st.number_input('Número de Passos', value=100)
                if st.button('Calcular Preço e Gregas das Opções'):
                    preco_opcao_compra = calcular_opcao_binomial(preco_subjacente, preco_exercicio, tempo, taxa_juros,
                                                                 dividendos,
                                                                 volatilidade, num_steps)
                    preco_opcao_venda = calcular_opcao_binomial(preco_subjacente, preco_exercicio, tempo, taxa_juros,
                                                                dividendos,
                                                                volatilidade, num_steps)
                    gregas_compra = calcular_gregas_bs(preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos,
                                                       volatilidade, 'Call')
                    gregas_venda = calcular_gregas_bs(preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos,
                                                      volatilidade, 'Put')
                    st.success(
                        f'Preço da Opção de Compra: {preco_opcao_compra:.2f}, Preço da Opção de Venda: {preco_opcao_venda:.2f}')
                    st.write('Gregas da Opção de Compra:')
                    st.json(gregas_compra)
                    st.write('Gregas da Opção de Venda:')
                    st.json(gregas_venda)

        elif tipo_opcao == 'Americana':
            if metodo_solucao == 'Monte Carlo':
                preco_opcao_compra = np.mean(np.maximum(S[:, -1] - preco_exercicio, 0) * discount_factor)
                preco_opcao_venda = np.mean(np.maximum(preco_exercicio - S[:, -1], 0) * discount_factor)
                return preco_opcao_compra, preco_opcao_venda
            elif metodo_solucao == 'Binomial':
                num_steps = st.number_input('Número de Passos', value=100)
                if st.button('Calcular Preço'):
                    preco_opcao_compra = calcular_opcao_binomial(preco_subjacente, preco_exercicio, tempo, taxa_juros,
                                                                 dividendos,
                                                                 volatilidade, num_steps)
                    preco_opcao_venda = calcular_opcao_binomial(preco_subjacente, preco_exercicio, tempo, taxa_juros,
                                                                dividendos,
                                                                volatilidade, num_steps)
    except Exception as e:
        st.error(f"An error occurred: {e}")

def ParisianPricer(S, K, r, T, sigma, barrier, barrier_duration, runs):
    dt = T / 365.0
    sqrt_dt = np.sqrt(dt)
    num_steps = int(T / dt)
    stock_prices = np.zeros((runs, num_steps + 1))
    stock_prices[:, 0] = S
    parisian_active = np.zeros((runs, num_steps + 1), dtype=bool)
    for i in range(1, num_steps + 1):
        z = np.random.normal(0, 1, runs)
        stock_prices[:, i] = stock_prices[:, i - 1] * np.exp((r - 0.5 * sigma ** 2) * dt + sigma * sqrt_dt * z)
        parisian_active[:, i] = parisian_active[:, i - 1] | (stock_prices[:, i] > barrier)
        active_duration = np.sum(parisian_active[:, :int(i - barrier_duration / dt)], axis=1) * dt
        parisian_active[np.logical_and(active_duration >= barrier_duration, parisian_active[:, i]), i] = False
    option_payoffs = np.maximum(stock_prices[:, -1] - K, 0)
    option_payoffs[np.any(parisian_active[:, 1:], axis=1)] = 0
    discount_factor = np.exp(-r * T)
    price = discount_factor * np.mean(option_payoffs)
    return price

def imp_vol(S0, K, T, r, market, Otype):
    e = Decimal('1e-15')
    x0 = Decimal('0.2')  # Um palpite inicial para sigma

    def newtons_method(S0, K, T, r, market, Otype, x0, e):
        delta = call_bsm(S0, K, r, T, Otype, x0) - market
        while abs(delta) > e:
            adjustment = (call_bsm(S0, K, r, T, Otype, x0) - market) / vega(S0, K, r, T, x0)
            x0 = x0 - adjustment
            delta = call_bsm(S0, K, r, T, Otype, x0) - market
        return x0

    sig = newtons_method(S0, K, T, r, market, Otype, x0, e)
    return sig * 100

# Funções de Processamento de Dados para Planilhas

def processar_dados_recap(dados_recap):
    linhas_recap = []
    for linha in dados_recap.strip().split("\n"):
        try:
            operacao, ticker, quantidade_str, valor = linha.split()
            quantidade = int(quantidade_str.replace(',', ''))
            valor = float(valor.replace(',', '.'))
            linhas_recap.append([operacao, ticker, quantidade, valor])
        except ValueError as e:
            st.error(f"Erro ao processar a linha: {linha}. Verifique o formato dos dados. Detalhes: {e}")
            continue
    return pd.DataFrame(linhas_recap, columns=["Operacao", "Ticker", "Quantidade", "Valor"])

def processar_dados_spread(dados):
    operacoes = {}
    regex = re.compile(r'([CV])\s(\S+)\s.*?(\d+)(k|K)\s.*?@\s*(\+?-?\d+)')
    for linha in dados:
        match = regex.search(linha)
        if match:
            operacao, ticker, quantidade_str, milhar, valor = match.groups()
            quantidade = int(quantidade_str) * 1000
            valor = int(valor)
            linha_id = f"{ticker}-{quantidade}-{valor}-{operacao}"
            if ticker not in operacoes:
                operacoes[ticker] = {"C": [], "V": []}
            operacoes[ticker][operacao].append((linha, quantidade, valor, linha_id))
    return operacoes

def mostrar_operacoes_spread(operacoes, ticker_escolhido, px_ref):
    if ticker_escolhido in operacoes:
        if 'selecionados' not in st.session_state:
            st.session_state['selecionados'] = set()
        compras = sorted(operacoes[ticker_escolhido]["C"], key=lambda x: x[2], reverse=True)
        vendas = sorted(operacoes[ticker_escolhido]["V"], key=lambda x: x[2])
        for lista_operacoes, tipo in [(compras, "Compras"), (vendas, "Vendas")]:
            st.subheader(f"{tipo} para {ticker_escolhido}:")
            for index, operacao in enumerate(lista_operacoes):
                diferencial = ((-operacao[2] / 10000) * px_ref) if tipo == "Compras" else (operacao[2] / 10000) * px_ref
                unique_key = f"{ticker_escolhido}-{index}-{tipo}"
                check = st.checkbox(f"{operacao[0]} | Diferencial: {diferencial:.6f} R$", key=unique_key,
                                    value=unique_key in st.session_state['selecionados'])
                if check:
                    st.session_state['selecionados'].add(unique_key)
                else:
                    st.session_state['selecionados'].discard(unique_key)

    pass
    
def calcular_niveis_kapitalo(client_orders, prices_df):
    # Supondo que 'client_orders' é um DataFrame com as colunas ['Cliente', 'Ativo', 'Quantidade', 'Preco_Referencia']
    # e 'prices_df' é um DataFrame com os preços atuais dos ativos

    # Unir os dados do cliente com os preços atuais
    merged_data = pd.merge(client_orders, prices_df, on='Ativo')

    # Calcular o nível (preço atual dividido pelo preço de referência)
    merged_data['Nivel'] = merged_data['Preco_Atual'] / merged_data['Preco_Referencia']

    return merged_data
    
def processar_dados_kapitalo(dados):
    operacoes = {}
    regex = re.compile(r'([CV])\s(\S+)\s.*?(\d+)(k|K)\s.*?@\s*(\+?-?\d+)')
    for linha in dados.strip().split('\n'):
        match = regex.search(linha)
        if match:
            operacao, ticker, quantidade_str, milhar, valor = match.groups()
            quantidade = int(quantidade_str) * 1000
            valor = int(valor)
            linha_id = f"{ticker}-{quantidade}-{valor}-{operacao}"
            if ticker not in operacoes:
                operacoes[ticker] = {"C": [], "V": []}
            operacoes[ticker][operacao].append((linha, quantidade, valor, linha_id))
    return operacoes

def mostrar_niveis_kapitalo(operacoes, ticker_escolhido, px_ref):
    if ticker_escolhido in operacoes:
        if 'selecionados_kapitalo' not in st.session_state:
            st.session_state['selecionados_kapitalo'] = set()
        compras = sorted(operacoes[ticker_escolhido]["C"], key=lambda x: x[2], reverse=True)
        vendas = sorted(operacoes[ticker_escolhido]["V"], key=lambda x: x[2])
        for lista_operacoes, tipo in [(compras, "Compras"), (vendas, "Vendas")]:
            st.subheader(f"{tipo} para {ticker_escolhido}:")
            for index, operacao in enumerate(lista_operacoes):
                diferencial = ((-operacao[2] / 10000) * px_ref) if tipo == "Compras" else (operacao[2] / 10000) * px_ref
                unique_key = f"{ticker_escolhido}-{index}-{tipo}"
                check = st.checkbox(f"{operacao[0]} | Diferencial: {diferencial:.6f} R$", key=unique_key,
                                    value=unique_key in st.session_state['selecionados_kapitalo'])
                if check:
                    st.session_state['selecionados_kapitalo'].add(unique_key)
                else:
                    st.session_state['selecionados_kapitalo'].discard(unique_key)

    pass    
def processar_dados_cash(dado, data_hoje):
    linhas = []
    for linha in dado.strip().split('\n'):
        try:
            operacao, produto, resto = linha.split(' ', 2)
            qtde, preco = resto.split(' @ ')
            qtde = qtde.replace('.', '').replace(',', '.')  # Remover pontos e ajustar vírgulas
            preco = preco.replace('.', '').replace(',', '.')  # Ajusta o preço para o formato correto
            qtde = float(qtde) * (-1 if operacao == 'V' else 1)
            linhas.append([data_hoje, produto, qtde, float(preco), "LIQUIDEZ"])
        except ValueError as e:
            st.error(f"Erro ao processar a linha: {linha}. Verifique o formato dos dados. Detalhes: {e}")
            continue
    return linhas

def processar_dados_futuros(dado, data_hoje, trader):
    linhas = []
    for linha in dado.strip().split('\n'):
        partes = linha.split('\t')
        if len(partes) == 4:
            operacao, produto, qtde, preco = partes
            qtde = qtde.replace(',', '')
            preco = preco.replace(',', '')
            qtde = float(qtde) * (-1 if operacao == 'S' else 1)
            book = "Hedge" if produto.startswith(("WDO", "DOL")) else "Direcional_Indice"
            linhas.append([data_hoje, produto, qtde, float(preco), book, "", trader, "LIQUIDEZ", "ITAU"])
    return linhas

def processar_dados_inoa_cash(dado, data_hoje):
    linhas = []
    for linha in dado.strip().split('\n'):
        partes = linha.split('\t')
        if len(partes) == 4:
            operacao, produto, qtde, preco = partes
            qtde = qtde.replace(',', '')
            preco = preco.replace(',', '')
            qtde = float(qtde) * (-1 if operacao == 'S' else 1)
            linhas.append([data_hoje, produto, qtde, float(preco), "LIQUIDEZ"])
    return linhas

def processar_dados_futuros_murilo(dado, data_hoje):
    linhas = []
    for linha in dado.strip().split('\n'):
        partes = linha.split('\t')
        if len(partes) == 4:
            operacao, produto, qtde, preco = partes
            qtde = qtde.replace(',', '')
            preco = preco.replace(',', '')
            qtde = float(qtde) * (-1 if operacao == 'S' else 1)
            linhas.append(["", data_hoje, produto, "Murilo Ortiz", "LIQUIDEZ", "ITAU", float(preco), qtde])
    return linhas

# Funções para SPX

def generate_xml(action, ticker, date, quantity, price, option_type, strike_price):
    formatted_date = datetime.strptime(str(date), '%Y-%m-%d').strftime('%m/%d/%y')
    formatted_date2 = datetime.strptime(str(date), '%Y-%m-%d').strftime('%d/%m/%Y')
    action_prefix = 'blis-xml;' + ('Buy' if action == 'Buy' else 'Sell')
    option_label = 'P' if option_type == 'Put' else 'C'
    ticker_formatted = f"{ticker} US {formatted_date} {option_label}{int(strike_price)}"
    xml_string = f"{action_prefix};{ticker_formatted};{option_type};{int(strike_price)};{formatted_date2};{quantity};{price:.6f}"
    return xml_string

def parse_trade_instructions_adjusted(text):
    lines = text.strip().split('\n')
    table1 = []
    table2 = []

    for line in lines:
        words = line.split()
        if len(words) >= 3:
            operation = 'S' if words[0] in ('S', 'SS') else 'B'
            quantity = words[1].replace(",", "")
            ticker = words[2].split('.')[0].upper()

            table1.append([operation, f'{ticker}.US', int(quantity)])
            inverted_operation = 'B' if operation == 'S' else 'S'
            table2.append([inverted_operation, f'{ticker}.US', int(quantity)])

    return table1, table2

# Configuração inicial do estado
if "abas_futuros" not in st.session_state:
    st.session_state.abas_futuros = {}
if "dados_futuros" not in st.session_state:
    st.session_state.dados_futuros = {}
if "selected_category" not in st.session_state:
    st.session_state.selected_category = None

# Estrutura de navegação
st.sidebar.title("Menu de Navegação")

# Primeiro, escolha a categoria principal
st.sidebar.subheader("📊 Selecione uma Categoria")
category = st.sidebar.selectbox("Categoria:", ["Nenhuma", "Arbitragem", "Opções", "Confirmações"])

# Armazenar a categoria selecionada no estado da sessão
st.session_state.selected_category = category

# Mostrar opções com base na categoria selecionada
if st.session_state.selected_category == "Arbitragem":
    st.sidebar.subheader("Escolha uma opção de Arbitragem")
    arb_opcoes = st.sidebar.radio(
        "Opções:",
        ('Spreads Arb', 'Estrutura a Termo de Vol', 'Niveis Kapitalo', 'Basket Fidessa')
    )

elif st.session_state.selected_category == "Opções":
    st.sidebar.subheader("Escolha uma opção de Opções")
    opcao_opcoes = st.sidebar.radio(
        "Opções:",
        ('XML Opção', 'Consolidado opções', 'Notional to shares', 'Planilha SPX', 'Pegar Volatilidade Histórica', 'Pegar Open Interest','Pegar Open Interest2', 'Pegar Open Interest3','Calcular Preço de Opções', 'Calcular Volatilidade Implícita')
    )

elif st.session_state.selected_category == "Confirmações":
    st.sidebar.subheader("Escolha uma opção de Confirmações")
    confirmacao_opcoes = st.sidebar.radio(
        "Opções:",
        ('Update com participação', 'Leitor Recap Kap', 'Gerar Excel', 'Comissions')
    )

# Implementação das funcionalidades com base na opção selecionada
if st.session_state.selected_category == "Arbitragem":
    if arb_opcoes == 'Spreads Arb':
        st.title("Spreads Arb")
        st.markdown("Conteúdo relacionado a Spreads Arb...")
        # Coloque aqui o código relacionado a Spreads Arb...

    elif arb_opcoes == 'Estrutura a Termo de Vol':
        st.title('Projeção de Volatilidade com GARCH')
        st.sidebar.header('Parâmetros')
        asset = st.sidebar.text_input('Ativo', value='^BVSP')
        start_date = st.sidebar.date_input('Data de Início', value=pd.to_datetime('2023-01-01'))
        end_date = st.sidebar.date_input('Data de Fim', value=pd.to_datetime('2024-07-01'))
        forecast_horizon = st.sidebar.number_input('Horizonte de Previsão (dias)', min_value=1, max_value=365, value=30)

        data = download_data(asset, start_date, end_date)
        if data is not None:
            returns = 100 * data['Adj Close'].pct_change().dropna()

            model = arch_model(returns, vol='Garch', p=1, q=1)
            model_fit = model.fit(disp='off')
            st.write(model_fit.summary())

            forecasts = model_fit.forecast(horizon=forecast_horizon)
            vol_forecast_daily = np.sqrt(forecasts.variance.values[-1, :])
            vol_forecast_annual = vol_forecast_daily * np.sqrt(252)

            dates = pd.date_range(start=returns.index[-1], periods=forecast_horizon, freq='B')
            vol_df = pd.DataFrame({'Date': dates, 'Volatility': vol_forecast_annual})
            vol_df.set_index('Date', inplace=True)

            plt.figure(figsize=(10, 6))
            plt.plot(vol_df.index, vol_df['Volatility'], marker='o')
            plt.title(f'Estrutura a Termo de Volatilidade Anualizada para {asset}')
            plt.xlabel('Data')
            plt.ylabel('Volatilidade Anualizada (%)')
            plt.grid(True)
            st.pyplot(plt)

            csv = vol_df.to_csv().encode('utf-8')
            st.download_button(
                label="Download CSV",
                data=csv,
                file_name='volatility_term_structure.csv',
                mime='text/csv',
            )

    elif arb_opcoes == 'Niveis Kapitalo':
        
        st.title('Níveis Kapitalo')
        
        if 'dados_operacoes_kapitalo' not in st.session_state:
            st.session_state['dados_operacoes_kapitalo'] = []
        
        with st.sidebar:
            st.header("Inserir e Gerenciar Dados - Níveis Kapitalo")
            dados_raw_kapitalo = st.text_area("Cole os dados das operações aqui:", height=150)
            if st.button("Salvar Dados Iniciais - Kapitalo"):
                if dados_raw_kapitalo:
                    linhas = dados_raw_kapitalo.strip().split("\n")
                    st.session_state['dados_operacoes_kapitalo'].extend(linhas)
                    st.success("Dados adicionados com sucesso!")
        
            dados_adicionais_kapitalo = st.text_area("Cole operações adicionais aqui:", height=150)
            if st.button("Adicionar Operações - Kapitalo"):
                if dados_adicionais_kapitalo:
                    linhas_adicionais = dados_adicionais_kapitalo.strip().split("\n")
                    st.session_state['dados_operacoes_kapitalo'].extend(linhas_adicionais)
                    st.success("Operações adicionais adicionadas com sucesso!")
        
            if st.button("Apagar Todos os Dados - Kapitalo"):
                st.session_state['dados_operacoes_kapitalo'] = []
                st.experimental_rerun()
        
        if 'dados_operacoes_kapitalo' in st.session_state and st.session_state['dados_operacoes_kapitalo']:
            operacoes_processadas_kapitalo = processar_dados_kapitalo(st.session_state['dados_operacoes_kapitalo'])
            tickers_kapitalo = list(operacoes_processadas_kapitalo.keys())
            ticker_escolhido_kapitalo = st.selectbox("Escolha um ticker - Kapitalo:", [""] + tickers_kapitalo)
            if ticker_escolhido_kapitalo:
                px_ref_kapitalo = st.number_input("Px Ref.:", min_value=0.01, step=0.01, format="%.2f", key=f"px_ref_{ticker_escolhido_kapitalo}")
                mostrar_niveis_kapitalo(operacoes_processadas_kapitalo, ticker_escolhido_kapitalo, px_ref_kapitalo)
                
    elif arb_opcoes == 'Basket Fidessa':
        st.title("Basket Fidessa")
        cliente = st.text_input("Nome do Cliente",)
        trade_text = st.text_area("Enter Trade Instructions:", height=300, value="S 506 ABBV\nS 500 AMZN\n...")

        if st.button("Generate Baskets"):
            table1, table2 = parse_trade_instructions_adjusted(trade_text)

            df_table1 = pd.DataFrame(table1, columns=['Type', 'Ticker', 'Quantity'])
            df_table2 = pd.DataFrame(table2, columns=['Type', 'Ticker', 'Quantity'])

            today = datetime.now().strftime('%m-%d-%Y')

            df_table1['Zero'] = 0
            df_table2['Zero'] = 0
            quantities_sum_table1 = sum_quantities_by_operation(df_table1)

            output1 = BytesIO()
            df_table1.to_csv(output1, index=False)
            output1.seek(0)

            output2 = BytesIO()
            df_table2.to_csv(output2, index=False)
            output2.seek(0)

            file_name1 = f"{cliente}_BASKET_{today}_table1.csv"
            file_name2 = f"{cliente}_BASKET_{today}_table2.csv"

            st.download_button("Download Table 1", data=output1, file_name=file_name1, mime='text/csv')

            st.write("Quantities Sum by side: ")
            st.write(quantities_sum_table1)

elif st.session_state.selected_category == "Opções":
    if opcao_opcoes == 'XML Opção':
        st.title("Options Data Input and XML Generator")
        with st.expander("Input Options Form"):
            with st.form("options_form"):
                cols = st.columns(4)
                with cols[0]:
                    action = st.selectbox("Action (Buy/Sell):", options=["Buy", "Sell"])
                with cols[1]:
                    ticker = st.text_input("Ticker (e.g., PBR):")
                with cols[2]:
                    date = st.date_input("Expiration Date:")
                with cols[3]:
                    quantity = st.number_input("Quantity:", min_value=0)

                cols2 = st.columns(3)
                with cols2[0]:
                    price = st.number_input("Option Price:", format="%.6f")
                with cols2[1]:
                    option_type = st.selectbox("Option Type (Call/Put):", ["Call", "Put"])
                with cols2[2]:
                    strike_price = st.number_input("Strike Price:", format="%.2f")

                submit_button = st.form_submit_button("Generate XML")

        if submit_button and all([action, ticker, date, quantity, price, option_type, strike_price]):
            commission = quantity * 0.25
            xml_result = generate_xml(action, ticker, date, quantity, price, option_type, strike_price)
            new_data = {
                "Action": action, "Ticker": ticker, "Date": date, "Quantity": quantity,
                "Price": price, "Option Type": option_type, "Strike Price": strike_price,
                "Commission": commission, "XML": xml_result
            }
            st.session_state['options_df'] = st.session_state['options_df'].append(new_data, ignore_index=True)

        with st.expander("Options Dashboard"):
            if not st.session_state['options_df'].empty:
                st.dataframe(st.session_state['options_df'], height=300)
                st.text_area("XML to Copy:", "\n".join(st.session_state['options_df']['XML']), height=100)

        with st.expander("Consolidated Dashboard"):
            if not st.session_state['options_df'].empty:
                consolidated_data = calculate_weighted_average(st.session_state['options_df'])
                formatted_data = consolidated_data.style.format({'Average_Price': '{:.6f}'})
                st.write("Consolidated Data with Average Prices:")
                st.dataframe(formatted_data)

        if st.button("Clear Data"):
            st.session_state['options_df'] = pd.DataFrame(columns=[
                "Action", "Ticker", "Date", "Quantity", "Price", "Option Type", "Strike Price", "Commission", "XML"
            ])
            st.rerun()

    elif opcao_opcoes == 'Consolidado opções':
        st.title("Options Data Analysis")
        with st.expander("Paste Data Here"):
            raw_data = st.text_area("Paste data in the following format: \nSide\tSymbol\tQuantity\tExecution Price\tStrike\tMaturity\tCALL / PUT\tCommission", height=300)
            process_button = st.button("Process Data")

        if process_button and raw_data:
            df = parse_data(raw_data)
            result_df = calculate_average_price(df)
            st.write("Aggregated and Averaged Data:")
            st.dataframe(result_df)

    elif opcao_opcoes == 'Notional to shares':
        st.title("Notional to Shares Calculator")
        api_key = "cnj4ughr01qkq94g9magcnj4ughr01qkq94g9mb0"
        ticker = st.text_input("Enter the stock ticker (e.g., AAPL):")
        notional_str = st.text_input("Enter the notional amount in dollars (e.g., 100k, 2m):")

        if st.button("Calculate Shares"):
            try:
                notional_dollars = parse_number_input(notional_str)
                if ticker and api_key:
                    price = get_real_time_price(ticker.upper(), api_key)
                    if price is not None:
                        shares = notional_dollars / price
                        formatted_shares = "{:,.0f}".format(shares)
                        st.write(f"Current Price: ${price:.2f}")
                        st.write(f"Number of Shares: {formatted_shares}")
            except ValueError as e:
                st.error(str(e))

    elif opcao_opcoes == 'Planilha SPX':
        st.title("Gerador de Planilha SPX")
        if st.button("Adicionar uma nova aba para Futuros"):
            nova_aba = f"Futuro_{len(st.session_state.abas_futuros) + 1}"
            st.session_state.abas_futuros[nova_aba] = ""  # Adiciona ao dicionário
            st.session_state.dados_futuros[nova_aba] = ""

        with st.form("input_form"):
            trader = st.text_input("Nome do Trader", value="LUCAS ROSSI")
            nome_arquivo = st.text_input("Nome do Excel", value="SPX_LUCAS_PRIMEIRA_TRANCHE")
            dados_cash = st.text_area("Cole os dados de CASH aqui: ex: V PETR4 159.362 @ 40,382615", height=150)
            dados_cash_inoa = st.text_area("Cole os dados de CASH INOA aqui: ex: S PETR3 639,342 41.779994", height=150)

            for aba in st.session_state.abas_futuros:
                st.session_state.dados_futuros[aba] = st.text_area(f"Cole os dados para {aba}:", height=150, key=aba)

            planilha_murilo = st.checkbox("Planilha do Murilo?")
            submitted = st.form_submit_button("Processar e Baixar Excel")

        if submitted:
            data_hoje = datetime.now().strftime('%d/%m/%Y')

            linhas_cash = processar_dados_cash(dados_cash, data_hoje)
            linhas_cash_inoa = processar_dados_inoa_cash(dados_cash_inoa, data_hoje)

            linhas_cash_total = linhas_cash + linhas_cash_inoa
            df_cash = pd.DataFrame(linhas_cash_total, columns=["Data", "Produto", "Qtde", "Preço", "Dealer"])

            futuros_dfs = {}
            for nome_aba, dados in st.session_state.dados_futuros.items():
                linhas_futuros = processar_dados_futuros(dados, data_hoje, trader)
                futuros_dfs[nome_aba] = pd.DataFrame(linhas_futuros, columns=["Data", "Produto", "Qtde", "Preço", "Book", "Fundo", "Trader", "Dealer", "Settle Dealer"])

            if planilha_murilo:
                dados_murilo = st.text_area("Cole os dados do Murilo aqui:", height=150)
                linhas_futuros_murilo = processar_dados_futuros_murilo(dados_murilo, data_hoje)
                df_futuros_murilo = pd.DataFrame(linhas_futuros_murilo, columns=["strategy", "date", "future", "trader", "dealer", "settle_dealer", "rate", "amount"])

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_cash.to_excel(writer, sheet_name='CASH', index=False)

                for nome_aba, df_futuros in futuros_dfs.items():
                    df_futuros.to_excel(writer, sheet_name=nome_aba, index=False)

                if planilha_murilo:
                    df_futuros_murilo.to_excel(writer, sheet_name='Murilo_Futuros', index=False)

            output.seek(0)
            today = datetime.now().strftime('%m_%d_%y')
            nome_do_arquivo_final = f"{nome_arquivo}_{today}.xlsx"

            st.download_button(label="Baixar Dados em Excel",
                               data=output,
                               file_name=nome_do_arquivo_final,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif opcao_opcoes == 'Pegar Volatilidade Histórica':
        ticker = st.text_input('Ticker do Ativo:', value='PETR4.SA')
        periodo = st.selectbox('Período', ['1mo', '3mo', '6mo', '1y'])
        if st.button('Buscar Volatilidade Histórica'):
            volatilidade = calcular_volatilidade_historica(ticker, periodo)
            st.success(f'Volatilidade Histórica para {ticker} no período de {periodo}: {volatilidade * 100:.2f}%')

    elif opcao_opcoes == 'Calcular Preço de Opções':
        tipo_opcao = st.selectbox('Tipo de Opção', ['Europeia', 'Americana', 'Parisian'])
        metodo_solucao = st.selectbox('Método de Solução', {
            'Europeia': ['Black-Scholes', 'Monte Carlo'],
            'Americana': ['Monte Carlo'],
            'Parisian': ['Parisian']
        }.get(tipo_opcao, ['Monte Carlo']))

        preco_subjacente = st.number_input('Preço do Ativo Subjacente', value=25.0)
        preco_exercicio = st.number_input('Preço de Exercício', value=30.0)
        data_vencimento = st.date_input('Data de Vencimento')
        taxa_juros = st.number_input('Taxa de Juros Livre de Risco (%)', value=0.0) / 100
        dividendos = st.number_input('Dividendos (%)', value=0.0) / 100
        volatilidade = st.number_input('Volatilidade (%)', value=20.0) / 100

        if metodo_solucao in ['Black-Scholes', 'Monte Carlo']:
            num_simulacoes = st.number_input("Número de simulações:", value=10000)
            num_periodos =  st.number_input("Número de períodos:", value=100)

        elif metodo_solucao == 'Parisian':
            barrier = st.number_input('Barreira', value=125.0)
            barrier_duration = st.number_input('Duração da Barreira (dias)', value=5.0)
            runs = st.number_input('Simulações Monte Carlo', value=100000, step=1000)

        hoje = pd.Timestamp('today').floor('D')
        vencimento = pd.Timestamp(data_vencimento)
        dias_corridos = (vencimento - hoje).days
        tempo = dias_corridos / 360

        if dias_corridos == 0:
            st.error('A data de vencimento não pode ser hoje. Por favor, selecione uma data futura.')
        else:
            tipo_opcao_escolhida = st.radio("Escolha o tipo da Opção", ('Call', 'Put'))

            if st.button('Calcular Preço das Opções e Gregas'):
                if metodo_solucao == 'Parisian':
                    price = ParisianPricer(preco_subjacente, preco_exercicio, taxa_juros, tempo, volatilidade, barrier, barrier_duration / 365, runs)
                    st.success(f'Preço da Opção Parisiense: {price:.4f}')
                else:
                    preco_opcao_compra, preco_opcao_venda = calcular_opcao(tipo_opcao, metodo_solucao, preco_subjacente,
                                                                           preco_exercicio, tempo, taxa_juros, dividendos,
                                                                           volatilidade, num_simulacoes, num_periodos)
                    gregas_compra = calcular_gregas_bs(preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos,
                                                       volatilidade, 'Call')
                    gregas_venda = calcular_gregas_bs(preco_subjacente, preco_exercicio, tempo, taxa_juros, dividendos,
                                                      volatilidade, 'Put')

                    if tipo_opcao_escolhida == 'Call':
                        st.success(f'Preço da Opção de Compra: {preco_opcao_compra:.4f}')
                        st.write('Gregas da Opção de Compra:')
                        st.json(gregas_compra)
                    elif tipo_opcao_escolhida == 'Put':
                        st.success(f'Preço da Opção de Venda: {preco_opcao_venda:.4f}')
                        st.write('Gregas da Opção de Venda:')
                        st.json(gregas_venda)

    elif opcao_opcoes == 'Calcular Volatilidade Implícita':
        Otype = st.radio("Tipo de Opção", ['Call', 'Put'])
        market_price = st.number_input('Preço de Mercado da Opção', value=5.0)
        S0 = st.number_input('Preço do Ativo Subjacente', value=100.0)
        K = st.number_input('Preço de Exercício', value=100.0)
        data_vencimento = st.date_input('Data de Vencimento')
        r = st.number_input('Taxa de Juros Livre de Risco (%)', value=0.0) / Decimal(100)
        hoje = pd.Timestamp.today().floor('D')

        vencimento = pd.Timestamp(data_vencimento)
        dias_corridos = (vencimento - hoje).days
        tempo = Decimal(dias_corridos) / Decimal(252)  # Conversão para anos em termos de dias de negociação

        if st.button('Calcular Volatilidade Implícita'):
            try:
                implied_vol = imp_vol(S0, K, tempo, r, market_price, Otype)
                if implied_vol is not None:
                    st.success(f'Volatilidade Implícita para {Otype} de: {implied_vol:.2f}% ')
                else:
                    st.error("Não foi possível calcular a volatilidade implícita. Verifique os inputs.")
            except Exception as e:
                st.error(f"Erro ao calcular a volatilidade implícita: {e}")

    elif opcao_opcoes == 'Pegar Open Interest':
        ticker_symbol = st.text_input('Insira o Ticker do Ativo (ex.: AAPL)')
        if ticker_symbol:
            ticker = yf.Ticker(ticker_symbol)
            expiries = ticker.options  # Pegar datas de vencimento disponíveis

            if expiries:
                selected_expiries = st.multiselect('Escolha as Datas de Vencimento:', expiries)

                if st.button('Gerar PDFs de Open Interest'):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        for expiry in selected_expiries:
                            opts = ticker.option_chain(expiry)
                            calls = opts.calls[['strike', 'openInterest']]
                            puts = opts.puts[['strike', 'openInterest']]

                            pdf_path = os.path.join(temp_dir, f'{ticker_symbol}_{expiry}.pdf')
                            with PdfPages(pdf_path) as pdf:
                                fig, axes = plt.subplots(1, 3, figsize=(30, 8))

                                calls_oi_grouped = calls.groupby('strike')['openInterest'].sum().reset_index()
                                axes[0].barh(calls_oi_grouped['strike'], calls_oi_grouped['openInterest'], color='skyblue')
                                axes[0].set_title(f'Calls Open Interest for {expiry}')
                                axes[0].set_ylabel('Strike Price')
                                axes[0].set_xlabel('Open Interest')

                                puts_oi_grouped = puts.groupby('strike')['openInterest'].sum().reset_index()
                                axes[1].barh(puts_oi_grouped['strike'], puts_oi_grouped['openInterest'], color='salmon')
                                axes[1].set_title(f'Puts Open Interest for {expiry}')
                                axes[1].set_ylabel('Strike Price')
                                axes[1].set_xlabel('Open Interest')

                                combined = pd.merge(calls_oi_grouped, puts_oi_grouped, on='strike', how='outer', suffixes=('_call', '_put')).fillna(0)
                                combined['difference'] = combined['openInterest_call'] - combined['openInterest_put']
                                axes[2].barh(combined['strike'], combined['difference'], color='purple')
                                axes[2].set_title(f'Difference (Calls - Puts) for {expiry}')
                                axes[2].set_ylabel('Strike Price')
                                axes[2].set_xlabel('Difference in Open Interest')

                                pdf.savefig(fig)
                                plt.close(fig)

                            with open(pdf_path, "rb") as f:
                                st.download_button(label=f"Download PDF for {expiry}",
                                                   data=f.read(),
                                                   file_name=os.path.basename(pdf_path),
                                                   mime='application/octet-stream')
            else:
                st.error("Não há datas de vencimento disponíveis para este ticker.")
        else:
            st.warning("Por favor, insira um ticker válido.")

    elif opcao_opcoes == 'Pegar Open Interest2':
        
        ticker_symbol = st.text_input('Insira o Ticker do Ativo (ex.: AAPL)')
        
        if ticker_symbol:
            ticker = yf.Ticker(ticker_symbol)
            expiries = ticker.options  # Pegar datas de vencimento disponíveis
    
            if expiries:
                if st.button('Gerar PDFs de Open Interest para Todos os Vencimentos'):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        for expiry in expiries:
                            opts = ticker.option_chain(expiry)
                            calls = opts.calls[['strike', 'openInterest']]
                            puts = opts.puts[['strike', 'openInterest']]
    
                            pdf_path = os.path.join(temp_dir, f'{ticker_symbol}_{expiry}.pdf')
                            with PdfPages(pdf_path) as pdf:
                                fig, axes = plt.subplots(1, 3, figsize=(30, 8))
    
                                calls_oi_grouped = calls.groupby('strike')['openInterest'].sum().reset_index()
                                axes[0].barh(calls_oi_grouped['strike'], calls_oi_grouped['openInterest'], color='skyblue')
                                axes[0].set_title(f'Calls Open Interest for {expiry}')
                                axes[0].set_ylabel('Strike Price')
                                axes[0].set_xlabel('Open Interest')
    
                                puts_oi_grouped = puts.groupby('strike')['openInterest'].sum().reset_index()
                                axes[1].barh(puts_oi_grouped['strike'], puts_oi_grouped['openInterest'], color='salmon')
                                axes[1].set_title(f'Puts Open Interest for {expiry}')
                                axes[1].set_ylabel('Strike Price')
                                axes[1].set_xlabel('Open Interest')
    
                                combined = pd.merge(calls_oi_grouped, puts_oi_grouped, on='strike', how='outer', suffixes=('_call', '_put')).fillna(0)
                                combined['difference'] = combined['openInterest_call'] - combined['openInterest_put']
                                axes[2].barh(combined['strike'], combined['difference'], color='purple')
                                axes[2].set_title(f'Difference (Calls - Puts) for {expiry}')
                                axes[2].set_ylabel('Strike Price')
                                axes[2].set_xlabel('Difference in Open Interest')
    
                                pdf.savefig(fig)
                                plt.close(fig)
    
                            # Adiciona botão de download para cada vencimento
                            with open(pdf_path, "rb") as f:
                                st.download_button(label=f"Download PDF for {expiry}",
                                                   data=f.read(),
                                                   file_name=os.path.basename(pdf_path),
                                                   mime='application/octet-stream')
            else:
                st.error("Não há datas de vencimento disponíveis para este ticker.")
        else:
            st.warning("Por favor, insira um ticker válido.")        



    elif opcao_opcoes == 'Pegar Open Interest3':
        # Input do Ticker
        ticker_symbol = st.text_input('Insira o Ticker do Ativo (ex.: AAPL)')
        
        if ticker_symbol:
            ticker = yf.Ticker(ticker_symbol)
            expiries = ticker.options  # Pegar datas de vencimento disponíveis
    
            if expiries:
                if st.button('Gerar Gráfico Consolidado de Open Interest e Baixar Excel'):
                    # DataFrames vazios para armazenar os valores agregados
                    all_calls = pd.DataFrame(columns=['strike', 'openInterest'])
                    all_puts = pd.DataFrame(columns=['strike', 'openInterest'])
    
                    # Criar lista para armazenar as options chains para salvar no Excel
                    calls_list = []
                    puts_list = []
    
                    # Iterar sobre todos os vencimentos e somar os valores de open interest por strike
                    for expiry in expiries:
                        opts = ticker.option_chain(expiry)
                        calls = opts.calls[['strike', 'openInterest']]
                        puts = opts.puts[['strike', 'openInterest']]
    
                        # Adicionar vencimento aos DataFrames
                        calls['expiry'] = expiry
                        puts['expiry'] = expiry

                        # Armazenar para salvar no Excel
                        calls_list.append(calls)
                        puts_list.append(puts)
    
                        # Somar o openInterest para cada strike nos calls e puts
                        all_calls = pd.concat([all_calls, calls]).groupby('strike', as_index=False).sum()
                        all_puts = pd.concat([all_puts, puts]).groupby('strike', as_index=False).sum()
    
                    # Plotar os gráficos consolidados
                    fig, ax = plt.subplots(1, 3, figsize=(30, 8))
    
                    # Gráfico de Calls Open Interest
                    ax[0].barh(all_calls['strike'], all_calls['openInterest'], color='skyblue')
                    ax[0].set_title(f'Calls Open Interest Consolidado')
                    ax[0].set_ylabel('Strike Price')
                    ax[0].set_xlabel('Open Interest')

                    # Adicionar rótulos com o valor correto do strike nas barras de Calls
                    for i, rect in enumerate(ax[0].patches):
                        strike = all_calls.iloc[i]['strike']
                        ax[0].annotate(f"{strike:.2f}",
                                       (rect.get_width() + 50, rect.get_y() + rect.get_height() / 2),
                                       va='center')
    
                    # Gráfico de Puts Open Interest
                    ax[1].barh(all_puts['strike'], all_puts['openInterest'], color='salmon')
                    ax[1].set_title(f'Puts Open Interest Consolidado')
                    ax[1].set_ylabel('Strike Price')
                    ax[1].set_xlabel('Open Interest')

                    # Adicionar rótulos com o valor correto do strike nas barras de Puts
                    for i, rect in enumerate(ax[1].patches):
                        strike = all_puts.iloc[i]['strike']
                        ax[1].annotate(f"{strike:.2f}",
                                       (rect.get_width() + 50, rect.get_y() + rect.get_height() / 2),
                                       va='center')
    
                    # Gráfico de Diferença entre Calls e Puts
                    combined = pd.merge(all_calls, all_puts, on='strike', how='outer', suffixes=('_call', '_put')).fillna(0)
                    combined['difference'] = combined['openInterest_call'] - combined['openInterest_put']
                    ax[2].barh(combined['strike'], combined['difference'], color='purple')
                    ax[2].set_title(f'Diferença (Calls - Puts) Consolidada')
                    ax[2].set_ylabel('Strike Price')
                    ax[2].set_xlabel('Diferença em Open Interest')

                    # Adicionar rótulos com o valor correto do strike nas barras de Diferença
                    for i, rect in enumerate(ax[2].patches):
                        strike = combined.iloc[i]['strike']
                        ax[2].annotate(f"{strike:.2f}",
                                       (rect.get_width() + 50, rect.get_y() + rect.get_height() / 2),
                                       va='center')
    
                    # Exibir os gráficos no Streamlit
                    st.pyplot(fig)

                    # Criar arquivo Excel em memória
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Concatenar todas as calls e puts
                        full_calls = pd.concat(calls_list)
                        full_puts = pd.concat(puts_list)
    
                        # Escrever os dados no Excel em abas separadas
                        full_calls.to_excel(writer, sheet_name='Calls', index=False)
                        full_puts.to_excel(writer, sheet_name='Puts', index=False)

                        writer.save()

                    # Oferecer o arquivo para download
                    st.download_button(
                        label="Baixar Excel das Options Chains",
                        data=output.getvalue(),
                        file_name=f'{ticker_symbol}_options_chain.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            else:
                st.error("Não há datas de vencimento disponíveis para este ticker.")
        else:
            st.warning("Por favor, insira um ticker válido.")



       

            


elif st.session_state.selected_category == "Confirmações":
    if confirmacao_opcoes == "Leitor Recap Kap":
        st.title('Leitor ADRxORD Kapitalo')
        uploaded_file = st.file_uploader("Escolha um arquivo")
        if uploaded_file is not None:
            processed_data = process_file(uploaded_file)
            st.write('Processed Data')
            st.dataframe(processed_data)

    elif confirmacao_opcoes == 'Gerar Excel':
        st.title("Gerar Excel a partir de Dados Colados")
        data = st.text_area("Cole os dados aqui, separados por espaço:", height=300)

        today = datetime.now().strftime("%Y%m%d")
        nome_arquivo = st.text_input("Nome do Arquivo Excel:", f"JP_BASKET{today}.xlsx")

        default_destinatario = "destinatario@example.com"
        default_assunto = f"JPM EXCEL {today}"
        default_corpo_email = ""

        destinatario = st.text_input("Email do Destinatário:", value=default_destinatario)
        assunto = st.text_input("Assunto do Email:", value=default_assunto)
        corpo_email = st.text_area("Corpo do Email:", value=default_corpo_email)

        if st.button('Gerar Excel'):
            if data:
                try:
                    data_io = StringIO(data)
                    df = pd.read_csv(data_io, sep="\s+", engine='python', skiprows=1)
                    df.to_excel(nome_arquivo, index=False, header=False)

                    with open(nome_arquivo, "rb") as f:
                        st.download_button("Baixar Excel", f.read(), file_name=nome_arquivo)
                    st.success("Excel gerado com sucesso!")
                    if st.button('Enviar Email via Outlook'):
                        try:
                            command = f'start "" "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Outlook.lnk" /c ipm.note /m "{destinatario}?subject={assunto}&body={corpo_email}" /a "{nome_arquivo}"'
                            result = subprocess.run(command, shell=True, capture_output=True, text=True)

                            if result.returncode != 0:
                                st.error("Falha ao abrir o Outlook")
                                st.error(f"Erro: {result.stderr}")
                            else:
                                st.success("Outlook aberto para envio de email!")

                        except Exception as e:
                            st.error(f"Ocorreu um erro ao tentar abrir o Outlook: {e}")
                except Exception as e:
                    st.error(f"Ocorreu um erro ao gerar o Excel: {e}")

    elif confirmacao_opcoes == 'Comissions':
        st.title("Comissions Off Shore")
        start_date = st.date_input('Data de Início', datetime(2023, 7, 1))
        end_date = st.date_input('Data de Término', datetime(2024, 1, 1))

        if st.button('Processar Dados'):
            consolidated_df = process_data(start_date, end_date)
            if not consolidated_df.empty:
                st.write("DataFrame consolidado criado com sucesso.")
                st.dataframe(consolidated_df)
                soma_produto = (consolidated_df.iloc[:, 5] * consolidated_df.iloc[:, 6]).sum()
                soma_shares = (consolidated_df.iloc[:, 5]).sum()
                st.write(f"A comissão consolidada para o período é de: {soma_produto:.2f} dólares")
                st.write(f"Total de shares é de: {soma_shares:.2f}")

                towrite = StringIO()
                consolidated_df.to_excel(towrite, index=False, engine='xlsxwriter')
                towrite.seek(0)
                st.download_button(label="Baixar Excel", data=towrite, file_name='comissoes.xlsx', mime='application/vnd.ms-excel')
            else:
                st.write("Nenhum dado encontrado para o período selecionado.")

    elif confirmacao_opcoes == "Update com participação":
        st.title("Market Participation Tracker")
        api_key = "cnj4ughr01qkq94g9magcnj4ughr01qkq94g9mb0"
        base_url = "https://finnhub.io/api/v1/quote"
        def get_stock_data(ticker):
            response = requests.get(f"{base_url}?symbol={ticker}&token={api_key}")
            if response.status_code == 200:
                data = response.json()
                price = data['c']  # Preço atual
                volume = data['v']  # Volume total do dia
                return price, volume
            else:
                return None, None

        if 'orders' not in st.session_state:
            st.session_state['orders'] = pd.DataFrame(columns=['Ticker', 'Shares', 'Initial Volume', 'Initial Participation'])

        with st.expander("Add New Order"):
            ticker = st.text_input("Enter the stock ticker (e.g., AAPL):", key='new_ticker')
            shares = st.number_input("Enter number of shares:", key='new_shares', min_value=0)
            initial_volume = st.number_input("Enter the initial volume when your order started:", key='new_initial_volume', min_value=0)

            if st.button("Add Order"):
                price, _ = get_stock_data(ticker.upper())
                if price:
                    st.session_state['orders'] = st.session_state['orders'].append({
                        'Ticker': ticker,
                        'Shares': shares,
                        'Initial Volume': initial_volume,
                        'Initial Participation': 0  # A ser calculado com o primeiro update
                    }, ignore_index=True)

        st.dataframe(st.session_state['orders'])

        with st.expander("Update Order"):
            selected_index = st.selectbox("Select Order to Update", st.session_state['orders'].index)
            new_volume = st.number_input("Enter new current volume:", key='new_volume_update', min_value=0)

            if st.button("Update Participation"):
                order = st.session_state['orders'].iloc[selected_index]
                _, current_volume = get_stock_data(order['Ticker'])
                if current_volume and new_volume > order['Initial Volume']:
                    participation = order['Shares'] / (new_volume - order['Initial Volume'])
                    st.session_state['orders'].loc[selected_index, 'Initial Participation'] = "{:.2%}".format(participation)
                    st.success(f"Updated participation for {order['Ticker']}.")

        st.dataframe(st.session_state['orders'])

# Inicializando o DataFrame de opções se não estiver presente
if 'options_df' not in st.session_state:
    st.session_state['options_df'] = pd.DataFrame(columns=["Action", "Ticker", "Date", "Quantity", "Price", "Option Type", "Strike Price", "XML"])

