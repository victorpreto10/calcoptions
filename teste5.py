import streamlit as st
import pandas as pd
import yfinance as yf
import numpy as np
import matplotlib.pyplot as plt
from arch import arch_model
from io import StringIO, BytesIO
from datetime import datetime
import os
import tempfile
from matplotlib.backends.backend_pdf import PdfPages
from decimal import Decimal, getcontext
from scipy.stats import norm
import math as m

# Configurando a precisão de decimais
getcontext().prec = 28

# -------------------------- Funções Utilitárias -------------------------- #

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

def generate_xml(action, ticker, date, quantity, price, option_type, strike_price):
    formatted_date = datetime.strptime(str(date), '%Y-%m-%d').strftime('%m/%d/%y')
    formatted_date2 = datetime.strptime(str(date), '%Y-%m-%d').strftime('%d/%m/%Y')
    action_prefix = 'blis-xml;' + ('Buy' if action == 'Buy' else 'Sell')
    option_label = 'P' if option_type == 'Put' else 'C'
    ticker_formatted = f"{ticker} US {formatted_date} {option_label}{int(strike_price)}"
    xml_string = f"{action_prefix};{ticker_formatted};{option_type};{int(strike_price)};{formatted_date2};{quantity};{price:.6f}"
    return xml_string

def format_number_input(input_str):
    input_str = input_str.lower().strip()
    if input_str.endswith('k'):
        return float(input_str[:-1]) * 1000
    elif input_str.endswith('m'):
        return float(input_str[:-1]) * 1000000
    elif input_str.replace('.', '', 1).isdigit():
        return float(input_str)
    else:
        raise ValueError("Entrada inválida: insira um número válido com 'k' para milhares ou 'm' para milhões, se necessário.")

# -------------------------- Funções Específicas -------------------------- #

def pegar_volatilidade_historica():
    ticker = st.text_input('Ticker do Ativo:', value='PETR4.SA')
    periodo = st.selectbox('Período', ['1mo', '3mo', '6mo', '1y'])
    if st.button('Buscar Volatilidade Histórica'):
        volatilidade = calcular_volatilidade_historica(ticker, periodo)
        st.success(f'Volatilidade Histórica para {ticker} no período de {periodo}: {volatilidade * 100:.2f}%')

def calcular_preco_opcoes():
    tipo_opcao = st.selectbox('Tipo de Opção', ['Europeia', 'Americana'])
    metodo_solucao = st.selectbox('Método de Solução', {
        'Europeia': ['Black-Scholes', 'Monte Carlo'],
        'Americana': ['Monte Carlo']
    }.get(tipo_opcao, ['Monte Carlo']))  # Mapeia tipos de opções com seus métodos correspondentes

    preco_subjacente = st.number_input('Preço do Ativo Subjacente', value=25.0)
    preco_exercicio = st.number_input('Preço de Exercício', value=30.0)
    data_vencimento = st.date_input('Data de Vencimento')
    taxa_juros = st.number_input('Taxa de Juros Livre de Risco (%)', value=0.0) / 100
    dividendos = st.number_input('Dividendos (%)', value=0.0) / 100
    volatilidade = st.number_input('Volatilidade (%)', value=20.0) / 100

    # Calculo do tempo até vencimento
    hoje = pd.Timestamp('today').floor('D')
    vencimento = pd.Timestamp(data_vencimento)
    dias_corridos = (vencimento - hoje).days
    tempo = dias_corridos / 360

    if dias_corridos == 0:
        st.error('A data de vencimento não pode ser hoje. Por favor, selecione uma data futura.')
        return

    tipo_opcao_escolhida = st.radio("Escolha o tipo da Opção", ('Call', 'Put'))

    if st.button('Calcular Preço das Opções e Gregas'):
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

def gerar_excel_dados_colados():
    st.title("Gerar Excel a partir de Dados Colados")
    data = st.text_area("Cole os dados aqui, separados por espaço:", height=300)
    nome_arquivo = st.text_input("Nome do Arquivo Excel:", f"JP_BASKET{datetime.now().strftime('%Y%m%d')}.xlsx")

    if st.button('Gerar Excel'):
        if data:
            try:
                data_io = StringIO(data)
                df = pd.read_csv(data_io, sep="\s+", engine='python')
                with BytesIO() as output:
                    df.to_excel(output, index=False)
                    output.seek(0)
                    st.download_button("Baixar Excel", data=output, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.success("Excel gerado com sucesso!")
            except Exception as e:
                st.error(f"Ocorreu um erro ao gerar o Excel: {e}")

def pegar_open_interest():
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

# -------------------------- Função Principal -------------------------- #

def main():
    # Configurando estados
    if "abas_futuros" not in st.session_state:
        st.session_state.abas_futuros = {}
    if "dados_futuros" not in st.session_state:
        st.session_state.dados_futuros = {}
    if 'options_df' not in st.session_state:
        st.session_state.options_df = pd.DataFrame(columns=["Action", "Ticker", "Date", "Quantity", "Price", "Option Type", "Strike Price", "XML"])

    # Estrutura de navegação
    st.sidebar.title("Menu de Navegação")

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

    # Lógica para exibir o conteúdo baseado na seleção
    if opcao_opcoes == 'Pegar Volatilidade Histórica':
        pegar_volatilidade_historica()

    elif opcao_opcoes == 'Calcular Preço de Opções':
        calcular_preco_opcoes()

    elif confirmacao_opcoes == 'Gerar Excel':
        gerar_excel_dados_colados()

    elif opcao_opcoes == 'Pegar Open Interest':
        pegar_open_interest()

    else:
        st.image('trading.jpg', use_column_width=True)
        st.title("Bem-vindo ao Dashboard de Opções")
        st.markdown("Escolha uma das opções no menu lateral para começar.")

if __name__ == "__main__":
    main()
