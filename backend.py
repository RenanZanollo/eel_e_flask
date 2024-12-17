from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium import webdriver as wd
import win32com.client as w32
from datetime import date
import os
import pythoncom
# Importações ^^

#______________________________________________________________________________________________________________________#
# Variáveis globais
path = os.getcwd() + '\\'
user = os.getcwd().split('\\')[2]
data = str(date.today()).split('-')
data = f'{int(data[2])}/{data[1]}/{data[0]}'
with open(f'{path}email.txt') as file:
    To, CC = file.read().strip().split(']', 1)
html = '<table border="1" cellpadding="5" cellspacing="0">'
exception = ''


#______________________________________________________________________________________________________________________#
# Funções

# Função para escrever os dados do login
def txtLoginWrite(id, password):
    with open(path + "login.txt", 'w') as file:
        file.write(id + ';' + password)


# Função para ler os dados do login
def txtLogin():
    with open(path + "login.txt", 'r') as file:
        id, password = file.read().strip().split(';' , 1)
    return id, password


# Função de inicialização do WebDriver
def __webdriver():
    chrome_options = Options()
    preferences = {"extensions.disabled": True,
                   "download.extensions_to_open": False,
                   "download.default_directory": path,
                   "directory_upgrade": True,
                   "safebrowsing.enabled": True,
                   "useAutomationExtension": False,
                   'safebrowsing.disable_download_protection': True}
    chrome_options.add_experimental_option("prefs", preferences)
    chrome_options.add_argument("--headless=old")  # roda em background
    chrome_options.add_argument('--log-level=2')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--remote-debugging-port=9222')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--incognito")  # private mode
    chrome_options.add_argument("--disable-dev-shm-usage")  # private mode
    chrome_options.add_argument("--force--device-scale-factor=1")
    chrome_options.add_argument("--disable-build-check")

    chrome_options.add_experimental_option('excludeSwitches', ['load-extension', 'enable-automation'])

    service = Service(f"C:\\Users\\{user}\\Documents\\chromedriver.exe")

    driver = wd.Chrome(options=chrome_options, service=service)
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': path}}
    command_result = driver.execute("send_command", params)
    return driver


# Função que obtém e altera os dados do Daily Report
def daily_report():
    pythoncom.CoInitialize()
    global html
    global exception
    Dict = {0: 'Entrada', 2: 'Entrada do almoço', 1: 'Saída do almoço', 3: 'Saída'}
    try:
        id, password = txtLogin()
        driver = __webdriver()
        driver.get(
            "https://portal.facens.br/Corpore.Net/Login.aspx?autoload=false&ReturnUrl=%2fCorpore.Net%2fMain.aspx%3fSelectedMenuIDKey%3d%26ShowMode%3d2")
        driver.find_element(By.NAME, "txtUser").send_keys(id)
        driver.find_element(By.NAME, "txtPass").send_keys(password)
        driver.find_element(By.NAME, "btnLogin").click()
        while driver.find_elements(By.ID, "ctl18_REC_PtoEspCartaoActionWeb_LinkControl") == 0:
            driver.implicitly_wait(1)
        driver.find_element(By.ID, "ctl18_REC_PtoEspCartaoActionWeb_LinkControl").click()
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.NAME, "ctl26$dpInicioPerMes$txtDate")))
        driver.find_element(By.NAME, "ctl26$dpInicioPerMes$txtDate").click()
        driver.find_element(By.NAME, "ctl26$dpInicioPerMes$txtDate").send_keys(data)
        driver.find_element(By.NAME, "ctl26$dpFimPerMes$txtDate").click()
        driver.find_element(By.NAME, "ctl26$dpFimPerMes$txtDate").send_keys(data)
        driver.find_element(By.ID, "ctl26_btnAtualizar_tblabel").click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.NAME, "ctl26$dpInicioPerMes$txtDate")))
        horary = str(driver.find_element(By.CLASS_NAME, "RowGrid").text).split(" ")
        driver.quit()
        horary = [str(horary[2]), str(horary[3]), str(horary[4]), str(horary[5])]
        horaryint = [0, 1, 2, 3]
        for i in range(len(horary)):
            if horary[i] == '':
                horaryint[i] = 0
            else:
                horaryint[i] = (int(horary[i].split(':')[0]) * 60) + int(horary[i].split(':')[1])
        entry1 = horaryint[0]
        exit1 = horaryint[1]
        if exit1 - entry1 >= 360:
            lunchTime = 15
        else:
            entry2 = horaryint[2]
            exit2 = horaryint[3]
            lunchTime = entry2 - exit1
            for i in range(len(horaryint)):
                if horaryint[i] == 0:
                    exception = f'Horário {Dict.get(i)} não batido'
                    return False

        excel = w32.DispatchEx("Excel.Application")
        wb = excel.Workbooks.Open(path + "Daily Report.xlsx")
        ws = wb.Sheets(1)
        excel.Visible = False
        ws.Cells(2, 2).Value = data
        ws.Cells(2, 5).Value = str(horary[0] + 'Hrs')

        if lunchTime == 15:
            ws.Cells(2, 6).Value = str("0:15Hrs")
            ws.Cells(2, 7).Value = str(horary[1] + 'Hrs')
        else:
            ws.Cells(2, 7).Value = str(horary[3] + 'Hrs')
            if lunchTime % 60 < 10:
                ws.Cells(2, 6).Value = f'{int(lunchTime / 60)}:0{lunchTime % 60}Hrs'
            else:
                ws.Cells(2, 6).Value = f'{int(lunchTime / 60)}:{lunchTime % 60}Hrs'

        rangeTable = ws.Range('B1:H2')
        for row in rangeTable.Rows:
            html += '<tr>'
            for cell in row.Columns:
                value = cell.Value
                html += f'<td>{value if value is not None else ""}</td>'
            html += "</tr>"
        html += '</table>'

        wb.SaveAs(path + 'Daily Report2.xlsx')
        wb.Close(SaveChanges=True)
        excel.Quit()
        if os.path.exists(path + 'Daily Report2.xlsx'):
            os.remove(path + 'Daily Report.xlsx')
            os.rename(path + 'Daily Report2.xlsx', path + 'Daily Report.xlsx')
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        exception = e
        pythoncom.CoUninitialize()
        return False


# Função que envia o email contendo o Daily e a Assinatura
def send_mail():
    pythoncom.CoInitialize()
    global html
    global exception
    try:
        outlook = w32.Dispatch("Outlook.Application")

        mail = outlook.Createitem(0)
        mail.To = To
        if CC != '':
            mail.CC = CC
        mail.Subject = 'Daily Report'
        mail.Display()
        signature = mail.HTMLBody
        mail.HTMLBody = f"""
        <p>Boa tarde!</p>
        {html}
        """ + signature
        mail.Send()
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        exception = e
        pythoncom.CoUninitialize()
        return False
