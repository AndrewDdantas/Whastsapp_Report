from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import os
import locale as lo
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from datetime import datetime
import sqlite3

def fmt_num(valor, tipo, casas=0): # Função para formatar números.
    lo.setlocale(lo.LC_ALL, 'pt_BR.UTF-8')
    if tipo == 'REAL':
        return lo.format_string(f"R$ %0.{casas}f",valor,grouping=True)
    if tipo == 'CUBAGEM':
        return lo.format_string(f"%0.{casas}f m³",valor,grouping=True)
    if tipo == 'NORMAL':
        return lo.format_string(f"%0.{casas}f",valor,grouping=True)
    if tipo == "PORCENTAGEM":
        return f"{{:.{casas}%}}".format(valor).replace('.',',')

def acessar_navegador(): # Função que gera navegador padrão.
    chrome_options = webdriver.ChromeOptions()
    path = __file__
    path = path.replace("report.py", 'navegador\chromedriver.exe')
    chrome_options.add_argument("user-data-dir=" + os.path.dirname(path))
    browser = webdriver.Chrome(options=chrome_options)
    return browser

def table(ax, dados, t=0):
    ax.axis('off')
    tabela = ax.table(cellText=dados.values, colLabels=dados.columns, loc='upper left')
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(14)
    if t == 3:
        tabela.scale(3, 2.5)
    else:
        tabela.scale(1, 2.5)
    ax.set_ylim(-10000, 100000)



    for (i,j), cell in tabela._cells.items():
        cell.get_text().set_ha('center')
        if i == 0:
            cell.set_fontsize(14)
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor('#00B0F0')
            cell.get_text().set_ha('center')
        else:
            cell.set_text_props(weight='bold', color='black')


def enviar():

    conn = sqlite3.connect('') #Conexão para banco de dados

    carteira = pd.read_sql_query('SELECT * FROM PEDIDOS', conn)
    carteira['VALTOTAL'] = carteira['PRECOUNIT'] * carteira['QTCOMP']

    valor = carteira.loc[carteira['TIPO_PEDIDO'] == 'VENDAS']
    valor = valor['VALTOTAL'].sum()

    meta = 1446000

    top_pedidos = carteira.loc[(carteira['STATUS_OPERACAO_GERENCIAL'] == 'EM CARTEIRA') & (carteira['TIPO_PEDIDO'] == 'VENDAS')]
    top_pedidos = top_pedidos.groupby('NUMPEDVEN').agg({'VALTOTAL': 'sum'}).reset_index().sort_values('VALTOTAL', ascending=False).head(10)
    top_pedidos['VALTOTAL'] = top_pedidos['VALTOTAL'].apply(fmt_num, tipo='REAL')

    top_lotes = carteira.loc[(carteira['STATUS'] == '4-Lote em Separacao') & (carteira['TIPO_PEDIDO'] == 'VENDAS')]
    top_lotes = top_lotes.groupby('NUMLOTE').agg({'VALTOTAL': 'sum'}).reset_index().sort_values('VALTOTAL', ascending=False).head(10)
    top_lotes['VALTOTAL'] = top_lotes['VALTOTAL'].apply(fmt_num, tipo='REAL')
    top_lotes['NUMLOTE'] = top_lotes['NUMLOTE'].apply(fmt_num, tipo='NORMAL')

    status = carteira.loc[carteira['TIPO_PEDIDO'] == 'VENDAS']
    status = status.groupby('STATUS').agg({'VALTOTAL': 'sum'}).reset_index().sort_values('VALTOTAL', ascending=False)
    status['VALTOTAL'] = status['VALTOTAL'].apply(fmt_num, tipo='REAL')

    fig = plt.figure(figsize=(18, 15))

    gs1 = GridSpec(2, 2, wspace=1, hspace=1)

    ax1 = fig.add_subplot(gs1[0:, 0:1])
    ax2 = fig.add_subplot(gs1[0:, 1:2])
    ax3 = fig.add_subplot(gs1[1:, :1])

    table(ax1, top_lotes)
    table(ax2, top_pedidos)
    table(ax3, status,3)

    fig.savefig('fechamento.png')

    saldo = valor - meta

    valor = fmt_num(valor, 'REAL')

    if saldo <= 0:
        saldo = fmt_num(saldo, 'REAL')
        mensagem = f'Olá! Meta batida, carteira atual {valor}, estamos com o saldo positivo de {saldo}.'
    else:
        saldo = fmt_num(saldo, 'REAL')
        mensagem = f'Olá! Meta não alcançada, carteira atual {valor}, estamos com o saldo negativo de {saldo}.'


    navegador = acessar_navegador()

    sleep(1)

    navegador.get('https://web.whatsapp.com/')

    while True:
        try:
            navegador.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div[3]/div/div[2]/button/div/div[2]/div/div')
            break
        except:
            sleep(3)

    sleep(2)

    navegador.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys('GRUPO_TESTE')

    sleep(2)

    navegador.find_element(by=By.XPATH, value=f'//*[@title="GRUPO_TESTE"]').click()

    sleep(2)

    navegador.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(mensagem, Keys.ENTER)

    sleep(2)

    navegador.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span').click()

    sleep(1)
    caminho = __file__
    caminho = caminho.replace('report.py', 'fechamento.png')

    navegador.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input').send_keys(caminho)

    sleep(3)

    navegador.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()

    sleep(3)

    navegador.close()


while True:
    data = datetime.now()
    hora = data.strftime('%H')
    hora1 = int(hora)
    hora = hora1
    enviar()
    while hora1 == hora:
        sleep(3)
        data = datetime.now()
        hora = data.strftime('%H')
        hora = int(hora)