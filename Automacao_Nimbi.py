import pyautogui
import pandas as pd
import time

tabela = pd.read_excel("Arquivo_Auxiliar.xlsx", sheet_name="PERIODO")

pyautogui.PAUSE = 0.5

pyautogui.hotkey("alt","tab") #Sair do VSCode para a tela do NIMBI
for linha in tabela.index:
    if tabela.loc[linha, "CONT"] == "SIM":
        data_inicio = pd.to_datetime(tabela.loc[linha, "DT INICIO"]).strftime("%d/%m/%Y")
        data_fim = pd.to_datetime(tabela.loc[linha, "DT FIM"]).strftime("%d/%m/%Y")
        pyautogui.click(x=421, y=580) #clicar na data inicio
        pyautogui.write(data_inicio)
        pyautogui.press("enter")
        time.sleep(1.5)
        pyautogui.click(x=378, y=651) #clicar na data inicio
        pyautogui.write(data_fim)
        pyautogui.press("enter")
        time.sleep(240) #Espera de 4 minutos para gerar o relatório
        pyautogui.click(x=1771, y=720) #clicar em Opções
        pyautogui.click(x=1769, y=760) #clicar em downloads
        time.sleep(60) #Tempo de esperar para gerar o download do arquivo