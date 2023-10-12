import os
import pyautogui
import win32com.client as win32
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


try:
    # Configurando o Selenium de maneira otimizada
    url_rpa_challenge = ("https://rpachallenge.com/")
    options = Options()
    options.page_load_strategy = 'eager'
    options.add_argument("--incognito")
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))
    driver.get(url_rpa_challenge)
    driver.maximize_window()


    # Salva arquivo e inicia o desafio
    v_challenge_file_path = r'C:\Users\Administrador\Downloads\challenge.xlsx'
    if os.path.exists(v_challenge_file_path):
        os.remove(v_challenge_file_path)
    driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a').click()
    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()


    # Instanciando a planilha de trabalho
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    wb = xl.Workbooks.Open(v_challenge_file_path)
    ws = wb.Sheets('Sheet1')

    v_ultima_linha_preenchida = ws.Cells(ws.Rows.Count,1).End(-4162).Row

    # Preenchimento do formulário com as variáveis extraídas da planilha
    for linha in range(2, v_ultima_linha_preenchida+1):
        
        v_first_name = ws.Range("A" + str(linha)).Value
        v_last_name = ws.Range("B" + str(linha)).Value
        v_company_name = ws.Range("C" + str(linha)).Value
        v_role_in_company = ws.Range("D" + str(linha)).Value
        v_address = ws.Range("E" + str(linha)).Value
        v_email = ws.Range("F" + str(linha)).Value
        v_phone_number = int(ws.Range("G" + str(linha)).Value)

        
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]').send_keys(v_first_name)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]').send_keys(v_last_name)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]').send_keys(v_company_name)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]').send_keys(v_role_in_company)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]').send_keys(v_address)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]').send_keys(v_email)
        driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]').send_keys(v_phone_number)


        print(str(linha), v_first_name, v_last_name, v_company_name, v_role_in_company, v_address, v_email, v_phone_number)
        sleep(3)

        driver.find_element(By.CLASS_NAME, 'btn.uiColorButton').click()

    wb.Close()

except Exception as err:
    print(err)
