from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pa
import pandas as pd

# Definindo a variavel navegador para pegar as funções do Chrome webdriver
navegador = webdriver.Chrome()

# Abrindo o site no chrome
navegador.get("https://rpachallengeocr.azurewebsites.net")

df_lista = []

linha = 1
i = 1
# While referente ao número de paginas (3)
while i < 4:

    elementoTabela = navegador.find_element(By.XPATH, '//*[@id="tableSandbox"]')

    linhas = elementoTabela.find_elements(By.TAG_NAME, "tr")
    colunas = elementoTabela.find_elements(By.TAG_NAME, "td")

    for linhaAtual in linhas:
        print(linhaAtual.text)
        df_lista.append(linhaAtual.text) #Adicionando as linhas no dataframe

        linha = linha + 1

    i = i + 1
    pa.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="tableSandbox_next"]').click() #XPATH do botão NEXT e clica
    pa.sleep(2)

else:
    print("EXTRAIDO COM SUCESSO!")



# Cria a planilha e Prepara o arquivo do Excel usando xlsxwriter com mecanismo
arquivoExcel = pd.ExcelWriter('dados_completos.xlsx', engine='xlsxwriter')

# Puxando os dados pra planilha
planilha_dados = pd.DataFrame(df_lista, columns=['#;ID;Due date'])

# Joga nossos dados dentro do Arquivo excel criado e salva
planilha_dados.to_excel(arquivoExcel,sheet_name='Sheet1', index=False)

arquivoExcel.close() #Salva os dados no arquivos