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




getcontext().prec = 28  # Definir precisão para operações Decimal


def call_bsm(S0, K, r, T, Otype, sig):


    S0, K, r, T, sig = map(Decimal, [S0, K, r, T, sig])  # Converter todos para Decimal

    d1 = Decimal(math.log(S0 / K)) + (r + (sig * sig) / 2) * T / (sig * Decimal(math.sqrt(T)))
    d2 = d1 - sig * Decimal(math.sqrt(T))

    if Otype == "Call":
        price = S0 * Decimal(norm.cdf(float(d1))) - K * Decimal(math.exp(-r * T)) * Decimal(norm.cdf(float(d2)))
    elif Otype == "Put":
        price = -S0 * Decimal(norm.cdf(float(-d1))) + K * Decimal(math.exp(-r * T)) * Decimal(norm.cdf(float(-d2)))
    return price
def vega(S0, K, r, T, sig):
    d1 = Decimal(m.log(S0 / K)) / (sig * Decimal(m.sqrt(T))) + Decimal(
        (r + (sig * sig) / 2) * T / (sig * Decimal(m.sqrt(T))))
    vega = S0 * Decimal(ss.norm.pdf(float(d1))) * Decimal(m.sqrt(T))
    return vega


def imp_vol(S0, K, T, r, market, flag):
    e = 10e-15
    x0 = Decimal(1)

    def newtons_method(S0, K, T, r, market, flag, x0, e):
        delta = call_bsm(S0, K, r, T, flag, x0) - market
        while delta > e:
            x0 = Decimal(x0 - (call_bsm(S0, K, r, T, flag, x0) - market) / vega(S0, K, r, T, x0))
            delta = abs(call_bsm(S0, K, r, T, flag, x0) - market)
        return Decimal(x0)

    sig = newtons_method(S0, K, T, r, market, flag, x0, e)
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
opcao = st.sidebar.radio(
    "Escolha uma opção:",
    ('Home', 'Calcular Volatilidade Implícita', 'Calcular Preço de Opções', 'Pegar Volatilidade Histórica','Pegar Open Interest', 'Gerar Excel'
))
if opcao == 'Home':
    st.image('trading.jpg', use_column_width=True)  # Coloque o caminho da sua imagem
    st.title("Bem-vindo ao Dashboard de Opções")
    st.markdown("Escolha uma das opções no menu lateral para começar.")


elif opcao == 'Pegar Volatilidade Histórica':
    ticker = st.text_input('Ticker do Ativo:', value='PETR4.SA')
    st.text(
        "O ticker deve seguir o mesmo padrão presente no Yahoo Finance. ")
    periodo = st.selectbox('Período', ['1mo', '3mo', '6mo', '1y'])
    if st.button('Buscar Volatilidade Histórica'):
        volatilidade = calcular_volatilidade_historica(ticker, periodo)
        st.success(f'Volatilidade Histórica para {ticker} no período de {periodo}: {volatilidade * 100:.2f}%')

elif opcao == 'Calcular Preço de Opções':
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


elif opcao == 'Calcular Volatilidade Implícita':
    Otype = st.radio("Tipo de Opção", ['Call', 'Put'])  # Opção para Put não implementada
    market_price = Decimal(st.number_input('Preço de Mercado da Opção', value=5.0))
    S0 = Decimal(st.number_input('Preço do Ativo Subjacente', value=100.0))
    K = Decimal(st.number_input('Preço de Exercício', value=100.0))
    data_vencimento = st.date_input('Data de Vencimento')
    r = Decimal(st.number_input('Taxa de Juros Livre de Risco (%)', value=0.0)) / Decimal(100)
    hoje = pd.Timestamp('today').floor('D')
    vencimento = pd.Timestamp(data_vencimento)
    dias_corridos = (vencimento - hoje).days
    tempo = Decimal(dias_corridos) / Decimal(252)  # Conversão para anos
    if st.button('Calcular Volatilidade Implícita'):
        implied_vol = imp_vol(S0, K, tempo, r, market_price, Otype)
        st.success(f'Volatilidade Implícita para {Otype} de: {implied_vol  :.2f}% ')
        st.experimental_rerun()

elif opcao == 'Pegar Open Interest':
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




elif opcao == 'Gerar Excel':
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
                
            except Exception as e:
                st.error(f"Ocorreu um erro ao gerar o Excel: {e}")

    if st.button('Enviar Email via Outlook'):
        if destinatario and assunto and corpo_email:
            try:
                command = f'"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE" /c ipm.note /m "{destinatario}?subject={assunto}&body={corpo_email}" /a "{nome_arquivo}"'

                result = subprocess.run(command, shell=True, capture_output=True, text=True)
                if result.returncode != 0:
                    st.error("Falha ao abrir o Outlook")
                    st.error(f"Erro: {result.stderr}")
                else:
                    st.success("Outlook aberto para envio de email!")
            except Exception as e:
                st.error(f"Ocorreu um erro ao tentar abrir o Outlook: {e}")
        else:
            st.error("Por favor, preencha todos os campos necessários para enviar o email.")
           
                                                       
                
                
           
      
