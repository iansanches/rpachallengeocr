from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
from datetime import datetime
#from selenium.webdriver.support import expected_conditions as EC
import pytesseract
import requests
import numpy as np
import cv2
import re
import csv 
#tableSandbox > tbody > tr:nth-child(1) > td:nth-child(3)
#tableSandbox > tbody > tr:nth-child(2) > td:nth-child(3)
#tableSandbox > thead > tr > th:nth-child(3)
# tr.odd:nth-child(1) > td:nth-child(3)

#URL do site acessado
url = "https://rpachallengeocr.azurewebsites.net/"
driver = Firefox()
driver.implicitly_wait(20)
driver.get(url)
# Configura o caminho para o executável do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#começar desafio
botao_iniciar = driver.find_element(By.ID, "start").click()

#pegar somente a data
datas = driver.find_elements(By.CSS_SELECTOR, '[role="row"] td:nth-child(3)')

#cabecalho do csv
cabecalho = ["ID","DueDate","InvoiceNo","InvoiceDate","CompanyName","TotalDue"]

nome_arquivo = "dados.csv"
lista_resultados = []  # Lista para armazenar todos os resultados
with open(nome_arquivo, "w", newline="") as arquivo_csv:
    writer = csv.writer(arquivo_csv)
    writer.writerow(cabecalho)
    while True:
        for data in datas:
            data_atual = datetime.now()
            data_informada = datetime.strptime(data.text, "%d-%m-%Y")
            if data_informada <= data_atual:
                id_element = data.find_element(By.XPATH, './preceding-sibling::td[1]')  # Localiza o elemento ID
                id_value = id_element.text  # Obtém o valor do ID
                botao = data.find_element(By.XPATH, './following-sibling::td/a') #pegar o botao
                href = botao.get_attribute("href") #pegar o valor dentro do botao
                botao.click()
                # Faz o download do conteúdo da imagem a partir da URL
                response = requests.get(href, stream=True).raw
                # Converte o conteúdo da imagem em uma matriz de pixels
                img_array = np.asarray(bytearray(response.read()), dtype=np.uint8)
                img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
                # Realiza o OCR na imagem
                resultado = pytesseract.image_to_string(img)
                valor_fatura = re.search(r"Total (\d+\.\d+)", resultado) #pegando valor da fatura
                data_fatura_regex = re.search(r"[0-9]{4}[-][0-9]{2}[-][0-9]{2}", resultado) #pegando valor data_fatura
                data_fatura = datetime.strptime(data_fatura_regex.group(0), "%Y-%m-%d").strftime("%d-%m-%Y")
                numero_fatura = re.search(r"Invoice #(\d+)", resultado)
                primeira_linha = resultado.split('\n')[0] #pegando a primeira linha
                palavras = primeira_linha.split(' ') #separando a primeira linha em palavras demoninando pelo espaço
                nome_empresa = ' '.join(palavras[:-1]) #fazendo join das palavras excluindo a ultima
                lista_valores = [id_value, data.text, numero_fatura.group(1), data_fatura, nome_empresa, valor_fatura.group(1)]
                lista_resultados.append(lista_valores)  # Adiciona a lista_valores à lista_resultados
        # Verifique se a classe "paginate_button next" existe
        next_button = driver.find_elements(By.CSS_SELECTOR, '.paginate_button.next')

        if len(next_button) > 0 and "disabled" not in next_button[0].get_attribute("class"):
            next_button[0].click()  # Clique no botão "Próximo"
            datas = driver.find_elements(By.CSS_SELECTOR, '[role="row"] td:nth-child(3)')  # Atualize os elementos de datas da proxima pagian
        else:
            break  # Se não houver mais botões "Próximo", saia do loop
    writer.writerows(lista_resultados)

#sleep(2)
input_arquivo = driver.find_element(By.NAME, 'csv')
input_arquivo.send_keys(r"C:\Users\infra\Documents\Python\Sharepoint_testes\dados.csv")
    

