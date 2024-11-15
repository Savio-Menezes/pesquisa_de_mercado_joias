from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import openpyxl

# Caminho do driver do Edge
driver_path = r"C:\Users\savio\OneDrive\Área de Trabalho\projetos pro github\automação_pesquisa_de_mercado\msedgedriver.exe"

service = Service(driver_path)

driver = webdriver.Edge(service=service)

driver.get("https://www.lazzulijoias.com.br/?utm_source=google-search&utm_medium=cpc&utm_campaign=ga-institucional&utm_content=ad01&parceiro=7854&gad_source=1&gclid=Cj0KCQiA_9u5BhCUARIsABbMSPvGVVAqZX0yO6RksbqxKA0FxB7wwlSCXkkUpnt82WYUSeadvnjUliwaAq4XEALw_wcB")

# Encontrar titulos
titulos = driver.find_elements(By.XPATH, "//*[contains(@class, 'product-name')]")

# Encontrar Preços dos produtos
precos = driver.find_elements(By.XPATH, "//*[contains(@class, 'current-price')]")

# Criando a planilha
workbook = openpyxl.Workbook()
workbook.create_sheet('produtos') 
sheets_produtos = workbook['produtos']

sheets_produtos["A1"].value = "Produto"
sheets_produtos["B1"].value = "Preco"

# Adicionar títulos e preços na planilha
for titulo, preco in zip(titulos, precos):
    sheets_produtos.append([titulo.text, preco.text])

# Salvar a planilha
workbook.save('pesquisa_de_mercado.xlsx')
