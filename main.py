from selenium import webdriver
import time
import openpyxl
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

base = pd.read_excel("base.xlsx",dtype=str)

print("Manipulando dados...")
for i in base["Unid"]:
    if len(i) < 4:
        while len(i) < 4:
            base = base.replace(i, (str(0)+i))
            i = str(0)+i

for i in base["CNPJ"]:
    if len(i) < 14:
        while len(i) < 14:
            base = base.replace(i, (str(0)+i))
            i = str(0)+i

for i in base["CPF"]:
    if isinstance(i, str):
        if len(i) < 11:
            while len(i) < 11:
                base = base.replace(i, (str(0)+i))
                i = str(0)+i
    
for i in base["codigo"]:
    if isinstance(i, str):
        if len(i) < 12:
            while len(i) < 12:
                base = base.replace(i, (str(0)+i))
                i = str(0)+i
print("Dados manipulados!")
                
base.sort_values(by=["Unid"], inplace=True)

print("\n\n\nBase de dados formatada!")
print("Iniciando automação do formulário em 3 segundos!\n\n\n")


jaforam = 0

for row in base.iterrows():
    driver = webdriver.Chrome()
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1200x600')
    url = "https://simplesnacional.efiscal.tec.br/?unidade="
    unidade = row[1]['Unid']
    cnpj = row[1]['CNPJ']
    cpf = row[1]['CPF']
    codigo = row[1]['codigo']
    segmento = row[1]['Segmento']
    estado = row[1]['Estado']
    ano = row[1]['Ano']

    driver.get(url+unidade)
    campocnpj = driver.find_element_by_xpath('//*[@id="form_diag_cod_acesso"]/div[1]/input')
    campocpf = driver.find_element_by_xpath('//*[@id="form_diag_cod_acesso"]/div[2]/input')
    campocodigo = driver.find_element_by_xpath('//*[@id="form_diag_cod_acesso"]/div[3]/input')
    camposegmento = driver.find_element_by_xpath('//*[@id="form_diag_cod_acesso"]/div[4]/select')
    campoestado = driver.find_element_by_xpath('//*[@id="uf"]')
    campoano = driver.find_element_by_xpath('//*[@id="form_diag_cod_acesso"]/div[6]/input')

    campocnpj.click()
    campocnpj.send_keys(cnpj)

    campocpf.click()
    if isinstance(cpf, float):
        campocpf.send_keys("00000000000")
    else:
        campocpf.send_keys(cpf)

    campocodigo.click()
    if isinstance(codigo, float):
        campocodigo.send_keys("000000000000")
    else:
        campocodigo.send_keys(codigo)

    campoano.click()
    campoano.send_keys(ano)

    select = Select(camposegmento)
    camposegmento.click()
    certosegmento = select.select_by_visible_text(segmento)
    select = Select(campoestado)
    campoestado.click()
    certoestado = select.select_by_visible_text(estado)

    campoano.click()
    campoano.send_keys(ano)

    botaoprocessar = driver.find_element_by_xpath('//*[@id="bt_envia_solicitacao_diagnostico"]')
    botaoprocessar.click()

    time.sleep(4)
    codigoposprocessamento = driver.find_element_by_xpath('//*[@id="md_retorno_solicitacao_diagnostico"]/div/div/div[2]/p/strong')
    codigoposprocessamento = codigoposprocessamento.text
    
    row[1]['codigofinal'] = codigoposprocessamento

    with open("codigosfinais.txt", 'a') as arquivo:
        arquivo.write(codigoposprocessamento+"\n")

    jaforam += 1
    print("Já foram:", jaforam)
    driver.quit()

print(base)
base.to_excel('output.xlsx', index=False)

print("Acabou!")