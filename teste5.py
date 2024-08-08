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
    st.session_state.abas_futuros = {}  # Certifique-se de que abas_futuros é um dicionário
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
        ('XML Opção', 'Consolidado opções', 'Notional to shares', 'Planilha SPX', 'Pegar Volatilidade Histórica', 'Pegar Open Interest', 'Calcular Preço de Opções', 'Calcular Volatilidade Implícita')
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
        # Código relacionado a Spreads Arb...

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
        st.title("Niveis Kapitalo")
        # Código relacionado a Niveis Kapitalo...

    elif arb_opcoes == 'Basket Fidessa':
        st.title("Basket Fidessa")
        cliente = st.text_input("Nome do Cliente")
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
            st.download_button("Download Table 2", data=output2, file_name=file_name2, mime='text/csv')

            st.write("Quantities Sum by side:")
            st.write(quantities_sum_table1)

elif st.session_state.selected_category == "Opções":
    if opcao_opcoes == 'XML Opção':
        st.title("XML Opção")
        # Código relacionado a XML Opção...

    elif opcao_opcoes == 'Consolidado opções':
        st.title("Consolidado opções")
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
        api_key = "your_api_key_here"
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
            st.session_state.abas_futuros[nova_aba] = ""  
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
        st.title("Pegar Volatilidade Histórica")
        ticker = st.text_input('Ticker do Ativo:', value='PETR4.SA')
        periodo = st.selectbox('Período', ['1mo', '3mo', '6mo', '1y'])
        if st.button('Buscar Volatilidade Histórica'):
            volatilidade = calcular_volatilidade_historica(ticker, periodo)
            st.success(f'Volatilidade Histórica para {ticker} no período de {periodo}: {volatilidade * 100:.2f}%')

    elif opcao_opcoes == 'Pegar Open Interest':
        st.title("Pegar Open Interest")
        ticker_symbol = st.text_input('Insira o Ticker do Ativo (ex.: AAPL)')
        if ticker_symbol:
            ticker = yf.Ticker(ticker_symbol)
            expiries = ticker.options  
            
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

    elif opcao_opcoes == 'Calcular Preço de Opções':
        st.title("Calcular Preço de Opções")
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
                    barrier = st.number_input('Barreira', value=125.0)
                    barrier_duration = st.number_input('Duração da Barreira (dias)', value=5.0)
                    runs = st.number_input('Simulações Monte Carlo', value=100000, step=1000)
                    price = ParisianPricer(preco_subjacente, preco_exercicio, taxa_juros, tempo, volatilidade, barrier, barrier_duration / 365, runs)
                    st.success(f'Preço da Opção Parisiense: {price:.4f}')
                else:
                    preco_opcao_compra, preco_opcao_venda = calcular_opcao(tipo_opcao, metodo_solucao, preco_subjacente,
                                                                           preco_exercicio, tempo, taxa_juros, dividendos,
                                                                           volatilidade)
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
        st.title("Calcular Volatilidade Implícita")
        Otype = st.radio("Tipo de Opção", ['Call', 'Put'])
        market_price = st.number_input('Preço de Mercado da Opção', value=5.0)
        S0 = st.number_input('Preço do Ativo Subjacente', value=100.0)
        K = st.number_input('Preço de Exercício', value=
