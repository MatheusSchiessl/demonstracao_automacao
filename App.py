from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QMessageBox
import pandas as pd
import os
from selenium import webdriver #Navegador
from selenium.webdriver.common.by import By #encontrar os elementos
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
import time
win_user = os.getlogin()
app = QtWidgets.QApplication([])
tela=uic.loadUi("resources/interface.ui")
def procura_arq():
    arquivo = QtWidgets.QFileDialog.getOpenFileName(tela,"Abrir Arquivo","C:/Users/matheusrs/Desktop","Excel(*.xlsx *.xls)")[0]
    tela.textEdit.setText(arquivo)


def executa():
    try:
        arquivo = tela.textEdit.toPlainText()
        df = pd.read_excel(arquivo)
        service = Service(executable_path='resources/chromedriver.exe')
        chrome = webdriver.Chrome(service=service)
        chrome.get("https://docs.google.com/forms/d/e/1FAIpQLSdGpqZgr_-Pm-CtG6S2n7Qk2kpj-B0N-uW6HQJL2qqB0Z_Llg/viewform?usp=sf_link")
        wait = WebDriverWait(chrome,10)
        for index, row in df.iterrows():
            
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')))
            elemento_texto_codigo = chrome.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_texto_nome=   chrome.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_texto_valor= chrome.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_texto_codigo.send_keys(row['CODIGO'])
            elemento_texto_nome.send_keys(row['NOME'])
            elemento_texto_valor.send_keys(row['VALOR'])
            time.sleep(1)
            chrome.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()
            chrome.find_element(By.XPATH,'/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() 
        chrome.quit()
        tela.textEdit.setText('')
        QMessageBox.about(tela,"Status da execução", "Execução concluída!")
    except Exception as e:
        print('O erro foi: ',e)

def fechar():
    tela.close()

tela.pushButton.clicked.connect(procura_arq)
tela.pushButton_2.clicked.connect(executa)
tela.pushButton_3.clicked.connect(fechar)

tela.show()
app.exec()