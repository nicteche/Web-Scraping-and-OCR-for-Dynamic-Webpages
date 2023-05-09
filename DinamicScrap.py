import os
import time
import chromedriver_autoinstaller
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium import webdriver
import PyPDF2
import pandas as pd
from datetime import datetime
import shutil
from pathlib import Path
#Variables
inicio = time.time()
fechaF5 = []
fechaCheck = []
fechaAppG = []

uniqueF5 = []
uniqueCheck = []
uniqueAppG = []

thorughputF5 = []
thorughputCheck = []
thorughputAppG = []

fechaPRTG = []
entrada= []
salida= []

    

#Path(Root and Downloads for any device)
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
downloads_path = str(Path.home() / "Downloads")
#Path to save the cookies or configuration of Chrome
pathChrom= os.path.expanduser('~')+'\\AppData\\Local\\Google\\Chrome\\User Data\\telcombas'
#AutoInstaler of chrome driver
chromedriver_autoinstaller.install()
#Options for Chrome, to make the scrapping more undetectable and efficient
options_cs = webdriver.ChromeOptions()
#options_cs.add_argument('--headless')
options_cs.add_argument(f'--user-data-dir={pathChrom}')
options_cs.add_argument("--disable-gpu")
options_cs.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ""Chrome/95.0.4638.54 Safari/537.36")
options_cs.add_experimental_option("excludeSwitches", ["enable-automation"])
options_cs.add_argument("--disable-blink-features=AutomationControlled")
options_cs.add_experimental_option("useAutomationExtension", False)
options_cs.add_experimental_option("excludeSwitches", ["enable-automation"])
options_cs.add_argument('--disable-extensions')
options_cs.add_argument('--ignore-certificate-errors')
options_cs.add_argument('--print-to-pdf')  # Habilitar la impresión a PDF
options_cs.add_argument('--no-sandbox')
options_cs.add_argument('--disable-dev-shm-usage')
options_cs.add_argument('--enable-print-browser')
options_cs.add_argument('--kiosk-printing')
browserVpn = webdriver.Chrome(chrome_options=options_cs)
#browserVpn.implicitly_wait(5)
url = 'https://10.225.200.25/pages/saved_reports'
browserVpn.get(url)
main_window = browserVpn.current_window_handle
time.sleep(15)

#Function to enter credentials to the page
def login(browser):
    #find the box to input the username
    nombre =browser.find_element(By.XPATH, '//*[@id="usernameField"]')
    nombre.send_keys("admin2")
    
    #find the box to input the password
    passw = browser.find_element(By.ID, 'passwordField')
    passw.send_keys("Pacifico2023")

    login = browser.find_element(By.CSS_SELECTOR, '#loginButtonContainer > input')
    login.click()


#Main
try:
    login(browserVpn)
except:
    print("No se logeo")

#the Sleep function to wait for the page to load completely.    
time.sleep(20)
#
pdf_list = []
i = 4
#Initiates a while loop to click and download reports
while i > 2:
    #
    now = time.localtime()
    #The page is refreshed every 30 seconds, then the program will be executed every 30 seconds.
    if now.tm_sec % 10 == 3:
        #Beutiful soup to save the html information of the page and be able to search the buttons.
        soup = BeautifulSoup(browserVpn.page_source, 'html.parser')
        elementos = soup.find_all(id=True)  
        lista1 = []
        #We go through the list of elements that has all the variables that have id in the page and add to a list
        for elem in elementos:
            ya = str(elem.get('id'))
            lista1.append(ya)
        #  We scroll through list1 to find the IDs starting with "yui_3_16_0_5_" which are the buttons on the page.
        for x in lista1:
            if x.startswith("yui_3_16_0_5_") and x not in pdf_list:
                #Try so that we do not have problems in not being able to click on the ID 
                try:
                    pdf = browserVpn.find_element(By.XPATH,f'//*[@id="{x}"]/td[3]/div/i')#//*[@id="yui_3_16_0_5_1677680032231_12742"]/td[3]/div/i
                    pdf_list.append(pdf.get_attribute("href"))
                    pdf.click()                    
                except:
                    print("Ese no tiene PDF")
        i= 0
    time.sleep(1)
#We performed Scrooll down to obtain the remaining page information.
browserVpn.execute_script("window.scrollTo(0, document.body.scrollHeight);")    
#We do the same as in the previous loop but in the opposite direction to obtain all the information we want.       
while i == 0:
    now = time.localtime()
    #We do the same as in the previous loop but in the opposite direction to obtain all the information we want.
    if now.tm_sec % 10 == 8:
        soup = BeautifulSoup(browserVpn.page_source, 'html.parser')
        elementos2 = soup.find_all(id=True)
        lista2 = []
        for elem in elementos2:
            ya = str(elem.get('id'))
            lista2.append(ya)
        #We change the order of List2 to start from the last one and get the reports from the end.
        lista2 = reversed(lista2)
        for x in lista2:
            if x.startswith("yui_3_16_0_5_") and ya not in pdf_list:
                try:
                    pdf = browserVpn.find_element(By.XPATH,f'//*[@id="{x}"]/td[3]/div/i')#//*[@id="yui_3_16_0_5_1677680032231_12742"]/td[3]/div/i
                    pdf_list.append(pdf.get_attribute("href"))
                    pdf.click()
                except:
                    print("Ese no tiene PDF")
        i= 1
    time.sleep(1)


#we are waiting if some reports have not yet been downloaded
time.sleep(15)


#We find the reports on downloads and move to the root 
listaArchivosdown = os.listdir(downloads_path)
try:
    for x in listaArchivosdown:
        pdfr= str(x)
        if pdfr.startswith("get"):
            shutil.move(downloads_path+"/"+pdfr,ROOT_DIR+"/"+pdfr)
except:
    #if reports have not yet been downloaded
    print("No se descargo bien")

#In this part, we started with OCR to read pdf and extract the information we needed.
listaArchivos = os.listdir(ROOT_DIR)
for x in listaArchivos:
    archivo = str(x)
    #we identify the file to read
    if archivo.startswith("get") or archivo.startswith("Flujo"):
         # open the PDF to read
        with open(archivo, 'rb') as pdf_file:
            # Create a PDF obj to read 
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            # We section the pages that we need the information for 
            page_obj = pdf_reader.pages[0]
            page_obj = page_obj.extract_text()
            page_obj = page_obj.splitlines()
            lineDate = page_obj[0]
            lineDate = lineDate.split(" ")
            tipoVpn = lineDate[lineDate.index("VPN")+1][:-1]
            lineFecha = page_obj[0].split("(")[1].split(")")[0].split("-")[0].strip()
            #if the report is an f5 data, we extract the information as follows
            if tipoVpn == "F5":
                fechaC = str(lineFecha).replace("01", "00")
                fecha_dt = datetime.strptime(fechaC, "%b %d, %Y %I:%M %p")
                fechaF5.append(fecha_dt)
                for x in page_obj:
                    if x.startswith("Unique Client"):
                        unique = x.split(" ")[-1][1:]
                        uniqueF5.append(unique)
                    if x.startswith("Total Throughput"):
                        total = x.split(" ")
                        if total[3] == "kbps":
                            totalF = round(float(total[2])/1000,2)
                            thorughputF5.append(totalF)
                        else:
                            thorughputF5.append(total[2])
            #if the report is an Checkpoint data, we extract the information as follows
            elif tipoVpn == "CHeckPoint":
                fechaC = str(lineFecha).replace("01", "00")
                fecha_dt = datetime.strptime(fechaC, "%b %d, %Y %I:%M %p")
                fechaCheck.append(fecha_dt)
                for x in page_obj:
                    if x.startswith("Unique Client"):
                        unique = x.split(" ")[-1][1:]
                        uniqueCheck.append(unique)
                    if x.startswith("Total Throughput"):
                        total = x.split(" ")
                        if total[3] == "kbps":
                            totalF = round(float(total[2])/1000,2)
                            thorughputCheck.append(totalF)
                        else:
                            thorughputCheck.append(total[2])
            #if the report is an AppGate data, we extract the information as follows
            elif tipoVpn == "AppGate":
                fechaC = str(lineFecha).replace("01", "00")
                fecha_dt = datetime.strptime(fechaC, "%b %d, %Y %I:%M %p")
                fechaAppG.append(fecha_dt)
                for x in page_obj:
                    if x.startswith("Unique Client"):
                        unique = x.split(" ")[-1][1:]
                        uniqueAppG.append(unique)
                    if x.startswith("Total Throughput"):
                        total = x.split(" ")
                        if total[3] == "kbps":
                            totalF = round(float(total[2])/1000,2)
                            thorughputAppG.append(totalF)
                        else:
                            thorughputAppG.append(total[2])
    #So them we chain the name to dont repeat the files, if the file already exists , it will be delete
    try:                       
        os.rename(archivo,"Flujo"+ str(lineFecha.replace(",","").replace(":00","").replace(" ",""))+ str(tipoVpn)+".pdf")
    except:
        if archivo.startswith("get"):
            os.remove(archivo)
        print("No se pudo renombrar el archivo")
           
#For this part, it was not possible to automate the prtg files, because I cannot use Web Scrapping. 
listaArchivos = os.listdir(ROOT_DIR)
for x in listaArchivos:
    archivo = str(x)
    if archivo.startswith("PRTG"):
         #This OCR works in the same way as the previous one.
        with open(archivo, 'rb') as pdf_file:
                    prtg_pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for i in range(len(prtg_pdf_reader.pages)):
                        prtg_page = prtg_pdf_reader.pages[i]
                        prtg_page = prtg_page.extract_text()
                        prtg_page = prtg_page.splitlines()
                       #From a single file, we can extract the data flow in the network.
                        for x in prtg_page:
                            if "10:00:00 AM −" in x:                               
                                fecha_datetime = datetime.strptime(x.split('−')[0].strip(), '%m/%d/%Y %I:%M:%S %p')
                                fechaPRTGF= datetime.strptime(x.split('−')[0].split(" ")[0].strip(), '%m/%d/%Y')
                                print(fechaPRTGF)
                                fechaPRTG.append(fecha_datetime)
                                valoresES = x.split("kbit/s")
                                entradaKbit = valoresES[0].split(" ")[-2].replace(",", ".")
                                salidaKbit = valoresES[1].split(" ")[-2].replace(",", ".")
                                entradaKbit = round(float(entradaKbit),2)
                                salidaKbit = round(float(salidaKbit),2)
                                entrada.append(entradaKbit)
                                salida.append(salidaKbit)
                            if " 3:00:00 PM −" in x:
                                fecha_datetime = datetime.strptime(x.split('−')[0].strip(), '%m/%d/%Y %I:%M:%S %p')
                                fechaPRTG.append(fecha_datetime)
                                valoresES = x.split("kbit/s")
                                entradaKbit = valoresES[0].split(" ")[-2].replace(",", ".")
                                salidaKbit = valoresES[1].split(" ")[-2].replace(",", ".")
                                entradaKbit = round(float(entradaKbit),2)   
                                salidaKbit = round(float(salidaKbit),2)
                                entrada.append(entradaKbit)
                                salida.append(salidaKbit)
                                # 6:00:00 PM −
                            if " 6:00:00 PM −" in x:
                                print(x)
                                fecha_datetime = datetime.strptime(x.split('−')[0].strip(), '%m/%d/%Y %I:%M:%S %p')
                                fechaPRTG.append(fecha_datetime)
                                #print("Encontrado")
                                valoresES = x.split("kbit/s")
                                entradaKbit = valoresES[0].split(" ")[-2].replace(",", ".")
                                entradaKbit = round(float(entradaKbit),2)
                                salidaKbit = valoresES[1].split(" ")[-2].replace(",", ".")
                                salidaKbit = round(float(salidaKbit),2)
                                entrada.append(entradaKbit)
                                salida.append(salidaKbit)
    try:
        fechaName = fecha_datetime.strftime('%Y-%m-%d %H:%M:%S')
        fechadoc= fechaName.split(" ")[0]
        os.rename(archivo,f"PRTG{fechadoc}.pdf")
    except:
        if archivo.startswith("PRTG"):
            os.remove(archivo)
        print("No se pudo renombrar el archivo")
#In this part, we create a Dataframe with all the data extraction from the previous steps.
dicF5 = {"Fecha": fechaF5, "Unique F5": uniqueF5, "Throughput F5": thorughputF5}
dicCheck = {"Fecha": fechaCheck, "Unique CheackPoint": uniqueCheck, "Throughput CheackPoint": thorughputCheck}
dicAppG = {"Fecha": fechaAppG, "Unique AppGate": uniqueAppG, "Throughput AppGate": thorughputAppG}
dictPRTG = {'Fecha': fechaPRTG, 'Entrada': entrada, 'Salida': salida}

dfPRTG = pd.DataFrame(dictPRTG)
dfPRTG["Entrada"] = dfPRTG["Entrada"].astype(float)
dfPRTG["Salida"] = dfPRTG["Salida"].astype(float)

try:
    dfPRTG = dfPRTG.drop_duplicates()
except:
    print("No se borro los repetidos")
    
dfF5 = pd.DataFrame(dicF5)
dfF5.drop_duplicates(keep='last', inplace= True)
dfCheck = pd.DataFrame(dicCheck)
dfCheck.drop_duplicates(keep='last', inplace= True)
dfAppG = pd.DataFrame(dicAppG)
dfAppG.drop_duplicates(keep='last', inplace= True)


dfAppG["Unique AppGate"] = dfAppG["Unique AppGate"].astype(int)
dfF5["Unique F5"] = dfF5["Unique F5"].astype(int)
dfCheck["Unique CheackPoint"] = dfCheck["Unique CheackPoint"].astype(int)

dfCheck["Throughput CheackPoint"] = dfCheck["Throughput CheackPoint"].astype(float)
dfAppG["Throughput AppGate"] = dfAppG["Throughput AppGate"].astype(float)
dfF5["Throughput F5"] = dfF5["Throughput F5"].astype(float)

#We join all the tables created, by the date column and order
tabla_completa = pd.merge(dfF5, dfCheck, on='Fecha')
tabla_completa = pd.merge(tabla_completa, dfAppG, on='Fecha')
tabla_completa = pd.merge(tabla_completa, dfPRTG, on='Fecha',how='outer')
final = tabla_completa.sort_values(by="Fecha", ascending=True)
#We create 1 column with the sum of 3 columns of the table.
final['Total de Usuarios'] = final[['Unique F5', 'Unique CheackPoint', 'Unique AppGate']].sum(axis=1)
fecha = datetime.today()
mes = fecha.month
day = fecha.day

final.to_excel(f"Reporte{mes}-{day}.xlsx", index=False)

#We create 1 column with the sum of 3 columns of the table.
try:
    df_existente = pd.read_excel("VPNsfinal2.xlsx")
    df_combinado = pd.concat([df_existente, final], ignore_index=True)
    df_combinado.drop_duplicates(subset="Fecha" ,keep= "first" ,inplace= True)
    df_combinado=  df_combinado.sort_values(by='Fecha', ascending=True)

    df_combinado.to_excel("VPNsfinal2.xlsx", index=False)
except:
    print("No hay historico")
#The time it takes for the program to execute
fin = time.time()
duracion = fin - inicio
print("La duracion del codigo se realizo en  ", duracion, " Segundos" )