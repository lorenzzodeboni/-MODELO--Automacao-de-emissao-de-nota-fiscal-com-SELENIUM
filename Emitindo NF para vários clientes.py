#!/usr/bin/env python
# coding: utf-8

# In[40]:



from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\loren\downloads",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=options)


# In[41]:


import os
caminho = os.getcwd()
arquivo = caminho + r'/login.html'
navegador.get(arquivo)


# In[36]:


# importando a base de clientes
import pandas as pd
tabela = pd.read_excel(r"NotasEmitir.xlsx")
display(df)


# In[42]:


# Fazendo Login
navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('lorenzzodeboni')
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('123456')
navegador.find_element(By.XPATH, '/html/body/div/form/button').click()


# In[38]:


for linha in tabela.index:
    # preencher os dados da NF
    navegador.find_element(By.XPATH, '//*[@id="nome"]').send_keys(tabela.loc[linha, 'Cliente'])
    navegador.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, 'Endereço'])
    navegador.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, 'Bairro'])
    navegador.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, 'Municipio'])
    navegador.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, 'CEP']))
    navegador.find_element(By.NAME, 'uf').send_keys(tabela.loc[linha, 'UF'])
    navegador.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, 'CPF/CNPJ']))
    navegador.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha,'Inscricao Estadual']))
    navegador.find_element(By.NAME, 'descricao').send_keys(tabela.loc[linha, 'Descrição'])
    navegador.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha,'Quantidade']))
    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, 'Valor Unitario']))
    navegador.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, 'Valor Total']))
    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    # recarregar a pagina
    navegador.refresh()
    
    

