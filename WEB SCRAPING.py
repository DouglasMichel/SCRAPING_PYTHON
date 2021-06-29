from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from datetime import date
import schedule
import time
import pandas as pd

#Funcao para executar o codigo de acordo com as configuracoes do schedule no final do codigo
def EXECUTA_CODIGO_NOVAMENTE():

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    nav = webdriver.Chrome(chrome_options=options)

    #Login SITE
    print('Estou realizando o login no site.')
    nav.get("https://site")
    nav.find_element_by_xpath('/html/body/div[16]/div[1]/div/div[2]/form/div/div[1]/div[1]/div/input').send_keys('usuário de login')
    nav.find_element_by_xpath('/html/body/div[16]/div[1]/div/div[2]/form/div/div[1]/div[2]/div/input').send_keys('Senha')
    nav.find_element_by_xpath('/html/body/div[16]/div[1]/div/div[2]/form/div/div[1]/div[2]/div/input').send_keys(Keys.ENTER)

    #Abre página de estração de relatório
    time.sleep(1)
    print('Estou realizando o acesso a página de extração do relatório.')
    time.sleep(2)
    nav.get("https://página de extração de relatório")

    #Ajusta tela para localizar opções
    nav.set_window_size(1024, 600)
    nav.maximize_window()





#EXTRAIR PRIMEIRO RELATORIO


    #Clicar sobre caixa de seleção de opções
    time.sleep(1)
    print('Estou realizando a seleção da caixa de filtro do relatório.')
    time.sleep(5)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/button/span[1]').click()

    #Clicar sobre segunda caixa de seleção de opções
    time.sleep(2)
    print('Estou realizando a seleção das filiais no filtro aberto.')
    time.sleep(1)
    
    #Clica em uma caixa de opções
    time.sleep(5) 
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div/div/button[2]/span').click()
    
    #seleciona opção de filtro 1
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[2]/a/span[1]').click()

    #seleciona opção de filtro 2
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[7]/a/span[1]').click()

    #seleciona opção de filtro 3
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[8]/a/span[1]').click()

    #seleciona opção de filtro 4
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[9]/a/span[1]').click()

    #seleciona opção de filtro 5
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[11]/a/span[1]').click()

    #seleciona opção de filtro 6
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[15]/a/span[1]').click()

    #seleciona opção de filtro 7
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[16]/a/span[1]').click()

    #seleciona opção de filtro 8
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[19]/a/span[1]').click()

    #seleciona opção de filtro 9
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[20]/a/span[1]').click()
    
    #seleciona opção de filtro 10
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[22]/a/span[1]').click()

    #seleciona opção de filtro 11
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[30]/a/span[1]').click()

    #seleciona opção de filtro 12
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[31]/a/span[1]').click()

    #seleciona opção de filtro 13
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[32]/a/span[1]').click()

    #seleciona opção de filtro 14
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/ul/li[38]/a/span[1]').click()


    #Abre caixa SELECIONAR TIPO DE RELATÓRIO
    time.sleep(3)
    print('Estou realizando a seleção do tipo do relatório.')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/button/span[1]').click()
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/div/input').send_keys('DIGITA O NOME DO RELATORIO PARA SELECIONAR O MESMO')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/ul/li[11]/a/span[1]').click()


    #Define variaveis com datas
    time.sleep(1)
    print('Estou realizando o processamento da data ideal para preencher o campo de datas.')
    time.sleep(1)
    data_final = date.today()
    data_inicial = date.fromordinal(data_final.toordinal() - 120)

    #preenche a data inicial
    time.sleep(2)
    print('Estou realizando o preenchimento do campo de data inicial.')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[3]/div[1]/div/input').send_keys(str(data_inicial))

    #preenche a data final
    time.sleep(1)
    print('Estou realizando o preenchimento do campo de data final.')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[3]/div[2]/div/input').send_keys(str(data_final))

    #clica em filtrar
    time.sleep(1)
    print('Estou realizando a seleção do botão "Filtrar".')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[4]/button').click()

    #clica em baixar relatorio
    time.sleep(10)
    print('Estou realizando a seleção do botão "Baixar Relatório".')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[3]/div/div/a/h3').click()
    time.sleep(20)





#EXTRAIR SEGUNDO RELATÓRIO

    #Abre caixa para selecionar o tipo do relatório
    print('Estou realizando a seleção do tipo do relatório.')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/button/span[1]').click()
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/div/input').send_keys('DIGITA O NOME DO RELATORIO PARA SELECIONAR O MESMO')
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[1]/div/div/div/ul/li[8]/a/span[1]').click()

    #clica em filtrar
    time.sleep(1)
    print('Estou realizando a seleção do botão "Filtrar".')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[2]/div[4]/button').click()

    #clica em baixar relatorio
    time.sleep(10)
    print('Estou realizando a seleção do botão "Baixar Relatório".')
    time.sleep(1)
    nav.find_element_by_xpath('/html/body/div[16]/div[3]/div/div/div/div/div[2]/div[3]/div/div/a/h3').click()

    #fechar navegador
    time.sleep(20)
    nav.quit()

#codigo agendador de execução da repetição da tarefa
schedule.every().day.at('06:00').do(EXECUTA_CODIGO_NOVAMENTE)

while 1:
    schedule.run_pending()
    time.sleep(21600)