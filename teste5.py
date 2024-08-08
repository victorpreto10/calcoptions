import streamlit as st
import yfinance as yf
import numpy as np
import pandas as pd
from scipy.stats import norm
from scipy.optimize import bisect
import math as m
from decimal import Decimal
import scipy.stats as ss
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
from datetime import datetime
import os
import subprocess
from io import StringIO
import re
from io import BytesIO
import requests
import streamlit as st
import pandas as pd
import yfinance as yf
from arch import arch_model
import numpy as np
import matplotlib.pyplot as plt
import time  

if "abas_futuros" not in st.session_state:
    st.session_state.abas_futuros = {}  # Certifique-se de que abas_futuros é um dicionário
if "dados_futuros" not in st.session_state:
    st.session_state.dados_futuros = {}


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





def calculate_weighted_average(df):
    df['Weighted_Price'] = df['Price'] * df['Quantity']
    result_df = df.groupby(['Action', 'Ticker', 'Date', 'Option Type', 'Strike Price']).agg(
        Total_Quantity=('Quantity', 'sum'),
        Total_Weighted_Price=('Weighted_Price', 'sum')
    ).reset_index()
    result_df['Average_Price'] = result_df['Total_Weighted_Price'] / result_df['Total_Quantity']
    return result_df.drop(columns=['Total_Weighted_Price'])

def parse_data(data):
    # Usando StringIO para converter a string em um dataframe
    data = StringIO(data)
    df = pd.read_csv(data, sep='\t', engine='python')
    # Remover espaços extras nos nomes das colunas e nos dados
    df.columns = df.columns.str.strip()
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def calculate_average_price(df):
    if 'Maturity' not in df.columns:
        st.error("Column 'Maturity' is missing from the input. Please check your data and try again.")
        return pd.DataFrame()
    # Converter a coluna 'Maturity' para datetime
    df['Maturity'] = pd.to_datetime(df['Maturity'], format='%m/%d/%Y', errors='coerce')
    if df['Maturity'].isna().any():
        st.error("Some dates in 'Maturity' column could not be parsed. Please check the format.")
        return pd.DataFrame()
    # Agrupar dados e calcular preço médio
    grouped_df = df.groupby(['Symbol', 'Side', 'Strike', 'CALL / PUT', 'Maturity']).agg(
        Quantity_Total=('Quantity', 'sum'),
        Average_Execution_Price=('Execution Price', 'mean'),
        Total_Commission=('Commission', 'sum')
    ).reset_index()
    return grouped_df



def format_date(date):
    return datetime.strptime(str(date), '%Y-%m-%d').strftime('%m/%d/%y')


def format_date2(date):
    return datetime.strptime(str(date), '%Y-%m-%d').strftime('%d/%m/%Y')

# Função para gerar a string XML
def generate_xml(action, ticker, date, quantity, price, option_type, strike_price):
    formatted_date = format_date(date)
    formatted_date2 = format_date2(date)
    action_prefix = 'blis-xml;' + ('Buy' if action == 'Buy' else 'Sell')
    option_label = 'P' if option_type == 'Put' else 'C'
    ticker_formatted = f"{ticker} US {formatted_date} {option_label}{int(strike_price)}"
    xml_string = f"{action_prefix};{ticker_formatted};{option_type};{int(strike_price)};{formatted_date2};{quantity};{price:.6f}"
    return xml_string


data_hoje = datetime.now().strftime('%m/%d/%Y')
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
    url = f'https://finnhub.io/api/v1/quote?symbol={ticker}&token=cnj4ughr01qkq94g9magcnj4ughr01qkq94g9mb0'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        current_price = data['c']  # 'c' is the current price field in the response
        return current_price
    else:
        st.error("Failed to fetch data from Finnhub.")
        return None



def sum_quantities_by_operation(df):
    # Convert the 'Quantity' column to numeric type to sum up correctly
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    # Group the DataFrame by the 'Type' and sum the 'Quantity' column
    quantities_sum = df.groupby('Type')['Quantity'].sum().to_dict()
    return quantities_sum

getcontext().prec = 28  # Definir precisão para operações Decimal

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

data_hoje = datetime.now().strftime('%Y-%m-%d')

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


def processar_dados_futuros(dado, data_hoje):
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


# Função para comparar DataFrame com dados colados
def compare_dataframes(df1, pasted_data):
    TESTDATA = StringIO(pasted_data)
    df2 = pd.read_csv(TESTDATA, sep="\t", header=None)
    df2.columns = ['Operation', 'Ticker Bloomberg', 'Qtde', 'Price']
    return df1.equals(df2)


def call_bsm(S0, K, r, T, Otype, sig):
    # Certifique-se de que todos os valores estão como Decimal
    S0, K, r, T, sig = map(Decimal, [S0, K, r, T, sig])

    d1 = (Decimal(math.log(S0 / K)) + (r + Decimal(0.5) * sig ** 2) * T) / (sig * Decimal(math.sqrt(T)))
    d2 = d1 - sig * Decimal(math.sqrt(T))

    if Otype == "Call":
        price = S0 * Decimal(norm.cdf(float(d1))) - K * Decimal(math.exp(-float(r * T))) * Decimal(norm.cdf(float(d2)))
    else:
        price = K * Decimal(math.exp(-float(r * T))) * Decimal(norm.cdf(float(-d2))) - S0 * Decimal(norm.cdf(float(-d1)))

    return price

def vega(S0, K, r, T, sig):
    # Ajuste similar para a função vega
    S0, K, r, T, sig = map(Decimal, [S0, K, r, T, sig])
    d1 = (Decimal(math.log(S0 / K)) + (r + Decimal(0.5) * sig ** 2) * T) / (sig * Decimal(math.sqrt(T)))
    return S0 * Decimal(norm.pdf(float(d1))) * Decimal(math.sqrt(T))

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


def calcular_volatilidade_historica(ticker, periodo):
    ativo = yf.Ticker(ticker)
    dados = ativo.history(period=periodo)
    retornos_log = np.log(dados['Close'] / dados['Close'].shift(1))
    volatilidade = retornos_log.std() * np.sqrt(252)  # Anualizando
    return volatilidade


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


# Estrutura de navegação
st.sidebar.title("Menu de Navegação")

# Separando as opções em categorias
st.sidebar.subheader("📊 Arbitragem")
arb_opcoes = st.sidebar.radio(
    "Escolha uma opção de Arbitragem:",
    ('Spreads Arb', 'Estrutura a Termo de Vol', 'Niveis Kapitalo', 'Basket Fidessa')
)

st.sidebar.subheader("⚙️ Opções")
opcao_opcoes = st.sidebar.radio(
    "Escolha uma opção de Opções:",
    ('XML Opção', 'Consolidado opções', 'Notional to shares', 'Planilha SPX', 'Pegar Volatilidade Histórica', 'Pegar Open Interest', 'Calcular Preço de Opções', 'Calcular Volatilidade Implícita')
)

st.sidebar.subheader("🔍 Confirmações")
confirmacao_opcoes = st.sidebar.radio(
    "Escolha uma opção de Confirmações:",
    ('Update com participação', 'Leitor Recap Kap', 'Gerar Excel', 'Comissions')
)

elif opcao_opcoes == 'Pegar Volatilidade Histórica':
    ticker = st.text_input('Ticker do Ativo:', value='PETR4.SA')
    st.text(
        "O ticker deve seguir o mesmo padrão presente no Yahoo Finance. ")
    periodo = st.selectbox('Período', ['1mo', '3mo', '6mo', '1y'])
    if st.button('Buscar Volatilidade Histórica'):
        volatilidade = calcular_volatilidade_historica(ticker, periodo)
        st.success(f'Volatilidade Histórica para {ticker} no período de {periodo}: {volatilidade * 100:.2f}%')

elif opcao_opcoes == 'Calcular Preço de Opções':
    # Seleção do tipo de opção
    tipo_opcao = st.selectbox('Tipo de Opção', ['Europeia', 'Americana', 'Parisian'])
    metodo_solucao = st.selectbox('Método de Solução', {
        'Europeia': ['Black-Scholes', 'Monte Carlo'],
        'Americana': ['Monte Carlo'],
        'Parisian': ['Parisian']
    }.get(tipo_opcao, ['Monte Carlo']))  # Mapeia tipos de opções com seus métodos correspondentes

    preco_subjacente = st.number_input('Preço do Ativo Subjacente', value=25.0)
    preco_exercicio = st.number_input('Preço de Exercício', value=30.0)
    data_vencimento = st.date_input('Data de Vencimento')
    taxa_juros = st.number_input('Taxa de Juros Livre de Risco (%)', value=0.0) / 100
    dividendos = st.number_input('Dividendos (%)', value=0.0) / 100
    volatilidade = st.number_input('Volatilidade (%)', value=20.0) / 100

    # Configuração baseada no método de solução
    if metodo_solucao in ['Black-Scholes', 'Monte Carlo']:
        num_simulacoes = st.number_input("Número de simulações:", value=10000)
        num_periodos =  st.number_input("Número de períodos:", value=100)

    elif metodo_solucao == 'Parisian':
        barrier = st.number_input('Barreira', value=125.0)
        barrier_duration = st.number_input('Duração da Barreira (dias)', value=5.0)
        runs = st.number_input('Simulações Monte Carlo', value=100000, step=1000)

    # Calculo do tempo até vencimento
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

                        # Criação do PDF para cada data de vencimento selecionada
                        pdf_path = os.path.join(temp_dir, f'{ticker_symbol}_{expiry}.pdf')
                        with PdfPages(pdf_path) as pdf:
                            fig, axes = plt.subplots(1, 3, figsize=(30, 8))

                            # Horizontal bar plot for Calls
                            calls_oi_grouped = calls.groupby('strike')['openInterest'].sum().reset_index()
                            axes[0].barh(calls_oi_grouped['strike'], calls_oi_grouped['openInterest'], color='skyblue')
                            axes[0].set_title(f'Calls Open Interest for {expiry}')
                            axes[0].set_ylabel('Strike Price')
                            axes[0].set_xlabel('Open Interest')

                            # Horizontal bar plot for Puts
                            puts_oi_grouped = puts.groupby('strike')['openInterest'].sum().reset_index()
                            axes[1].barh(puts_oi_grouped['strike'], puts_oi_grouped['openInterest'], color='salmon')
                            axes[1].set_title(f'Puts Open Interest for {expiry}')
                            axes[1].set_ylabel('Strike Price')
                            axes[1].set_xlabel('Open Interest')

                            # Horizontal bar plot for Differences
                            combined = pd.merge(calls_oi_grouped, puts_oi_grouped, on='strike', how='outer', suffixes=('_call', '_put')).fillna(0)
                            combined['difference'] = combined['openInterest_call'] - combined['openInterest_put']
                            axes[2].barh(combined['strike'], combined['difference'], color='purple')
                            axes[2].set_title(f'Difference (Calls - Puts) for {expiry}')
                            axes[2].set_ylabel('Strike Price')
                            axes[2].set_xlabel('Difference in Open Interest')

                            pdf.savefig(fig)
                            plt.close(fig)

                        # Providing a download button for each PDF
                        with open(pdf_path, "rb") as f:
                            st.download_button(label=f"Download PDF for {expiry}",
                                               data=f.read(),
                                               file_name=os.path.basename(pdf_path),
                                               mime='application/octet-stream')
        else:
            st.error("Não há datas de vencimento disponíveis para este ticker.")
    else:
        st.warning("Por favor, insira um ticker válido.")

    def custom_css():
        st.markdown(
            """
            <style>
                .big-font {
                    font-size:30px !important;
                    font-weight: bold;
                }
                .dataframe {
                    border: 2px solid #f0f0f0;
                    border-radius: 5px;
                }
            </style>
            """, unsafe_allow_html=True
        )
    
    custom_css()
    
    
    
elif arb_opcoes == 'Spreads Arb':

    # Título da página
    st.title('Dashboard de Arbitragem por Cliente')

    # Inicialização do DataFrame se não existir no state
    if 'data' not in st.session_state:
        st.session_state['data'] = pd.DataFrame(columns=['Cliente', 'Tipo', 'Ativo', 'BPS', 'Size'])
    
    with st.expander("Adicionar Nova Operação"):
        with st.form("operation_form"):
            cols = st.columns(3)
            cliente = cols[0].text_input('Nome do Cliente')
            tipo = cols[1].selectbox('Tipo', ['Buy', 'Sell'])
            ativo = cols[2].text_input('Ativo')
            
            cols2 = st.columns(2)
            bps = cols2[0].number_input('Nível de BPS')
            size = cols2[1].number_input('Size')
            submit_button = st.form_submit_button('Adicionar')
            
            if submit_button:
                new_data = {'Cliente': cliente, 'Tipo': tipo, 'Ativo': ativo, 'BPS': bps, 'Size': size}
                st.session_state['data'] = st.session_state['data'].append(new_data, ignore_index=True)
    
    cliente_selecionado = st.selectbox('Filtrar por Cliente:', ['Todos'] + list(st.session_state['data']['Cliente'].unique()))
    
    filtered_data = st.session_state['data'] if cliente_selecionado == 'Todos' else st.session_state['data'][st.session_state['data']['Cliente'] == cliente_selecionado]
    
    if not filtered_data.empty:
        aggregated_data = filtered_data.groupby(['Cliente', 'Tipo', 'Ativo', 'BPS', 'Size']).size().reset_index(name='Count')
        st.write("Dados de Arbitragem por Cliente:")
        st.dataframe(aggregated_data)
    else:
        st.write("Nenhum dado para mostrar.")
    
    if st.button('Limpar Dados'):
        st.session_state['data'] = pd.DataFrame(columns=['Cliente', 'Tipo', 'Ativo', 'BPS', 'Size'])
        st.experimental_rerun()


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
                        # Montando o comando para abrir o Outlook
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


elif confirmacao_opcoes == 'Leitor Recap Kap':
    st.title('Leitor ADRxORD Kapitalo')
    uploaded_file = st.file_uploader("Choose a file")
    if uploaded_file is not None:
        processed_data = process_file(uploaded_file)
        st.write('Processed Data')
        st.dataframe(processed_data)

    
    # Função para processar os dados de operações inseridos pelo usuário
def processar_dados(dados):
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


# Atualize esta função conforme necessário para o seu uso
def mostrar_operacoes(operacoes, ticker_escolhido, px_ref):
    if ticker_escolhido in operacoes:
        if 'selecionados' not in st.session_state:
            st.session_state['selecionados'] = set()
        compras = sorted(operacoes[ticker_escolhido]["C"], key=lambda x: x[2], reverse=True)
        vendas = sorted(operacoes[ticker_escolhido]["V"], key=lambda x: x[2])
        for lista_operacoes, tipo in [(compras, "Compras"), (vendas, "Vendas")]:
            st.subheader(f"{tipo} para {ticker_escolhido}:")
            for index, operacao in enumerate(lista_operacoes):
                # Creating a unique key for each checkbox using ticker, index, and operation type
                diferencial = ((-operacao[2] / 10000) * px_ref) if tipo == "Compras" else (operacao[2] / 10000) * px_ref
                unique_key = f"{ticker_escolhido}-{index}-{tipo}"  # Ensure this key is unique
                check = st.checkbox(f"{operacao[0]} | Diferencial: {diferencial:.6f} R$", key=unique_key,
                                    value=unique_key in st.session_state['selecionados'])
                if check:
                    st.session_state['selecionados'].add(unique_key)
                else:
                    st.session_state['selecionados'].discard(unique_key)

    pass


if 'dados_operacoes' not in st.session_state:
    st.session_state['dados_operacoes'] = []

if 'px_ref_por_ativo' not in st.session_state:
    st.session_state['px_ref_por_ativo'] = {}

if 'limpar_adicionais' not in st.session_state:
    st.session_state['limpar_adicionais'] = False

elif opcao == 'Niveis Kapitalo':
    if 'dados_operacoes' not in st.session_state:
        st.session_state['dados_operacoes'] = []
    
    with st.sidebar:
        st.header("Inserir e Gerenciar Dados")
        dados_raw = st.text_area("Cole os dados das operações aqui:", height=150)
        if st.button("Salvar Dados Iniciais"):
            if dados_raw:
                linhas = dados_raw.strip().split("\n")
                st.session_state['dados_operacoes'].extend(linhas)
                st.success("Dados adicionados com sucesso!")
    
        dados_adicionais = st.text_area("Cole operações adicionais aqui:", height=150)
        if st.button("Adicionar Operações"):
            if dados_adicionais:
                linhas_adicionais = dados_adicionais.strip().split("\n")
                st.session_state['dados_operacoes'].extend(linhas_adicionais)
                st.success("Operações adicionais adicionadas com sucesso!")
    
        if st.button("Apagar Todos os Dados"):
            st.session_state['dados_operacoes'] = []
            st.experimental_rerun()
    
    # Main display area
    if 'dados_operacoes' in st.session_state and st.session_state['dados_operacoes']:
        operacoes_processadas = processar_dados(st.session_state['dados_operacoes'])
        tickers = list(operacoes_processadas.keys())
        ticker_escolhido = st.selectbox("Escolha um ticker:", [""] + tickers)
        if ticker_escolhido:
            px_ref = st.number_input("Px Ref.:", min_value=0.01, step=0.01, format="%.2f", key=f"px_ref_{ticker_escolhido}")
            mostrar_operacoes(operacoes_processadas, ticker_escolhido, px_ref)
            
if "abas_futuros" not in st.session_state:
    st.session_state.abas_futuros = {}  # Inicializa como um dicionário
if "dados_futuros" not in st.session_state:
    st.session_state.dados_futuros = {}

elif opcao_opcoes == 'Planilha SPX':
    st.title("Gerador de Planilha SPX")
    
    if st.button("Adicionar uma nova aba para Futuros"):
        nova_aba = f"Futuro_{len(st.session_state.abas_futuros) + 1}"
        st.session_state.abas_futuros[nova_aba] = ""  # Adiciona ao dicionário
        st.session_state.dados_futuros[nova_aba] = ""
    
    # Formulário principal
    with st.form("input_form"):
        trader = st.text_input("Nome do Trader", value="LUCAS ROSSI")
        nome_arquivo = st.text_input("Nome do Excel", value="SPX_LUCAS_PRIMEIRA_TRANCHE")
        dados_cash = st.text_area("Cole os dados de CASH aqui: ex: V PETR4 159.362 @ 40,382615", height=150)
        dados_cash_inoa = st.text_area("Cole os dados de CASH INOA aqui: ex: S PETR3 639,342 41.779994", height=150)
        
        # Exibir caixas de texto para cada aba de futuros
        for aba in st.session_state.abas_futuros:
            st.session_state.dados_futuros[aba] = st.text_area(f"Cole os dados para {aba}:", height=150, key=aba)
        
        planilha_murilo = st.checkbox("Planilha do Murilo?")
        submitted = st.form_submit_button("Processar e Baixar Excel")
    
    if submitted:
        data_hoje = datetime.now().strftime('%d/%m/%Y')
        
        # Processar dados CASH
        linhas_cash = processar_dados_cash(dados_cash, data_hoje)
        linhas_cash_inoa = processar_dados_inoa_cash(dados_cash_inoa, data_hoje)
        
        # Consolidando todos os dados de CASH
        linhas_cash_total = linhas_cash + linhas_cash_inoa
        df_cash = pd.DataFrame(linhas_cash_total, columns=["Data", "Produto", "Qtde", "Preço", "Dealer"])
        
        # Processar dados FUTUROS
        futuros_dfs = {}
        for nome_aba, dados in st.session_state.dados_futuros.items():
            linhas_futuros = processar_dados_futuros(dados, data_hoje)
            futuros_dfs[nome_aba] = pd.DataFrame(linhas_futuros, columns=["Data", "Produto", "Qtde", "Preço", "Book", "Fundo", "Trader", "Dealer", "Settle Dealer"])

        if planilha_murilo:
            dados_murilo = st.text_area("Cole os dados do Murilo aqui:", height=150)
            linhas_futuros_murilo = processar_dados_futuros_murilo(dados_murilo, data_hoje)
            df_futuros_murilo = pd.DataFrame(linhas_futuros_murilo, columns=["strategy", "date", "future", "trader", "dealer", "settle_dealer", "rate", "amount"])
        
        # Gerar Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_cash.to_excel(writer, sheet_name='CASH', index=False)
            
            # Adicionando abas para futuros
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


elif arb_opcoes == 'Basket Fidessa':
    st.title("Basket Fidessa")
    
    cliente = st.text_input("Nome do Cliente",)
    trade_text = st.text_area("Enter Trade Instructions:", height=300, value="S 506 ABBV\nS 500 AMZN\n...")
    
    if st.button("Generate Baskets"):
        table1, table2 = parse_trade_instructions_adjusted(trade_text)
        
        df_table1 = pd.DataFrame(table1, columns=['Type', 'Ticker', 'Quantity'])
        df_table2 = pd.DataFrame(table2, columns=['Type', 'Ticker', 'Quantity'])
    
        today = datetime.now().strftime('%m-%d-%Y')
    
        # Add zero column
        df_table1['Zero'] = 0
        df_table2['Zero'] = 0
        quantities_sum_table1 = sum_quantities_by_operation(df_table1)
    
        # Saving dataframes to CSV in memory
        output1 = BytesIO()
        df_table1.to_csv(output1, index=False)
        output1.seek(0)
    
        output2 = BytesIO()
        df_table2.to_csv(output2, index=False)
        output2.seek(0)
    
        file_name1 = f"{cliente}_BASKET_{today}_table1.csv"
        file_name2 = f"{cliente}_BASKET_{today}_table2.csv"
    
        st.download_button("Download Table 1", data=output1, file_name=file_name1, mime='text/csv')
        
    
    # Optional: Display the tables in Streamlit
          # Display quantities sum
        st.write("Quantities Sum by side: ")
        st.write(quantities_sum_table1)

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

if 'options_df' not in st.session_state:
    st.session_state['options_df'] = pd.DataFrame(columns=["Action", "Ticker", "Date", "Quantity", "Price", "Option Type", "Strike Price", "XML"])

elif opcao_opcoes == "XML Opção":
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
            formatted_data = consolidated_data.style.format({'Average_Price': '{:.6f}'})  # Formatar para 6 casas decimais
            st.write("Consolidated Data with Average Prices:")
            st.dataframe(formatted_data)

    if st.button("Clear Data"):
        st.session_state['options_df'] = pd.DataFrame(columns=[
            "Action", "Ticker", "Date", "Quantity", "Price", "Option Type", "Strike Price", "Commission", "XML"
        ])
        st.rerun()


elif opcao_opcoes == 'Consolidado opções':
    st.title("Options Data Analysis")

# Aba para entrada de dados
    with st.expander("Paste Data Here"):
        raw_data = st.text_area("Paste data in the following format: \nSide\tSymbol\tQuantity\tExecution Price\tStrike\tMaturity\tCALL / PUT\tCommission", height=300)
        process_button = st.button("Process Data")
    
    if process_button and raw_data:
        df = parse_data(raw_data)
        result_df = calculate_average_price(df)
        st.write("Aggregated and Averaged Data:")
        st.dataframe(result_df)



elif confirmacao_opcoes == 'Comissions':
    st.title("Comissions Off Shore")
    start_date = st.date_input('Data de Início', datetime(2023, 7, 1))
    end_date = st.date_input('Data de Término', datetime(2024, 1, 1))

    if st.button('Processar Dados'):
        consolidated_df = process_data(start_date, end_date)
        if not consolidated_df.empty:
            st.write("DataFrame consolidado criado com sucesso.")
            st.dataframe(consolidated_df)
            # Calcular o produto das colunas 5 e 6 e somar os resultados
            soma_produto = (consolidated_df.iloc[:, 5] * consolidated_df.iloc[:, 6]).sum()
            soma_shares = (consolidated_df.iloc[:, 5]).sum()
            st.write(f"A comissão consolidada para o período é de: {soma_produto:.2f} dólares")
            st.write(f"Total de shares é de: {soma_shares:.2f}")

            # Opção de download do DataFrame como Excel
            towrite = StringIO()
            consolidated_df.to_excel(towrite, index=False, engine='xlsxwriter')
            towrite.seek(0)
            st.download_button(label="Baixar Excel", data=towrite, file_name='comissoes.xlsx', mime='application/vnd.ms-excel')
        else:
            st.write("Nenhum dado encontrado para o período selecionado.")
            
elif arb_opcoes == 'Estrutura a Termo de Vol':
    st.title('Projeção de Volatilidade com GARCH')

    # Inputs do usuário
    st.sidebar.header('Parâmetros')
    asset = st.sidebar.text_input('Ativo', value='^BVSP')
    start_date = st.sidebar.date_input('Data de Início', value=pd.to_datetime('2023-01-01'))
    end_date = st.sidebar.date_input('Data de Fim', value=pd.to_datetime('2024-07-01'))
    forecast_horizon = st.sidebar.number_input('Horizonte de Previsão (dias)', min_value=1, max_value=365, value=30)
    
    # Baixar os dados históricos com tentativa de repetição em caso de falha
    data = download_data(asset, start_date, end_date)
    if data is not None:
        returns = 100 * data['Adj Close'].pct_change().dropna()

        # Ajustar um modelo GARCH(1,1)
        model = arch_model(returns, vol='Garch', p=1, q=1)
        model_fit = model.fit(disp='off')
        st.write(model_fit.summary())

        # Prever a volatilidade futura
        forecasts = model_fit.forecast(horizon=forecast_horizon)
        vol_forecast_daily = np.sqrt(forecasts.variance.values[-1, :])
        vol_forecast_annual = vol_forecast_daily * np.sqrt(252)

        # Estrutura a termo de volatilidade
        dates = pd.date_range(start=returns.index[-1], periods=forecast_horizon, freq='B')
        vol_df = pd.DataFrame({'Date': dates, 'Volatility': vol_forecast_annual})
        vol_df.set_index('Date', inplace=True)

        # Plotar a estrutura a termo de volatilidade anualizada
        plt.figure(figsize=(10, 6))
        plt.plot(vol_df.index, vol_df['Volatility'], marker='o')
        plt.title(f'Estrutura a Termo de Volatilidade Anualizada para {asset}')
        plt.xlabel('Data')
        plt.ylabel('Volatilidade Anualizada (%)')
        plt.grid(True)
        st.pyplot(plt)

        # Salvar a estrutura a termo de volatilidade em um arquivo CSV
        csv = vol_df.to_csv().encode('utf-8')
        st.download_button(
            label="Download CSV",
            data=csv,
            file_name='volatility_term_structure.csv',
            mime='text/csv',
        )
    
    
else:
    st.image('trading.jpg', use_column_width=True)  # Coloque o caminho da sua imagem
    st.title("Bem-vindo ao Dashboard de Opções")
    st.markdown("Escolha uma das opções no menu lateral para começar.")


                                                       
                
                
           
      
