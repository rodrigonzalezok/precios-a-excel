from os import system
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions 
import time
from time import sleep
import pandas as pd

# Ingresamosruta del archivo Excel
ruta = #Colocar aquí ruta donde se encuentra el archivo excel

# Generamos el browser
browser = webdriver.Chrome(#colocar ruta donde se encuentra el driver)

# Logueo
email = 'rodrigo.gonzalez.doldan@gmail.com' 
pas = 'bullmarket123'

# Maximiza tamaño de pantalla
browser.maximize_window()

# Carga de la pagina
browser.get("https://www.bullmarketbrokers.com/")

# esperamos a que no se vea el aviso de cargando
WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))

# clic en el botón ingresar
browser.find_element(By.LINK_TEXT, "Ingresar").click()

# Esperamos a que se vea el boton de login del formulario 
WebDriverWait(browser, 30000).until(expected_conditions.visibility_of_element_located((By.ID, "btn_login_ok")))
browser.find_element(By.ID, "txt_modal_login_idNumber").send_keys(email)
browser.find_element(By.ID, "txt_modal_login_password").send_keys(pas)

# Clic en el botón de login
browser.find_element(By.ID, "btn_login_ok").click()

# Clic en el botón de Cotizaciones
WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
WebDriverWait(browser, 30000).until(expected_conditions.presence_of_all_elements_located((By.CLASS_NAME, "home-logged")))
time.sleep(2)

while(True):
    browser.find_element_by_link_text("COTIZACIONES").click()

    # Extraemos la tabla del Panel Lider
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until(expected_conditions.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]/tr[21]")))
    tablapanellider = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textpl = tablapanellider.text
    pl = textpl.split()

    panellider = []
    for i in range(0, len(pl), 12):
        panellider+=[[pl[i], pl[i+1], pl[i+2], pl[i+9]]]

    dataframepl = pd.DataFrame(panellider, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón ACCIONES
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[1]/span").click()

    # Click en el botón Panel General
    WebDriverWait(browser, 30000).until(expected_conditions.visibility_of_element_located((By.ID, "a_stock_panel")))
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[1]/ul/li[2]/a").click()

    # Extraemos la tabla del Panel General
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablapanelgeneral = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textpg = tablapanelgeneral.text
    pg = textpg.split()

    panelgeneral = []
    for i in range(0, len(pg), 12):
        panelgeneral+=[[pg[i], pg[i+1], pg[i+2], pg[i+9]]]

    dataframepg = pd.DataFrame(panelgeneral, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón BONOS
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[3]/span").click()

    # Click en el botón TODOS LOS BONOS
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[3]/ul/li[1]/a").click()

    # Extraemos la tabla de Bonos
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablabonos = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textbonos = tablabonos.text
    bonos = textbonos.split()

    panelbonos = []
    for i in range(0, len(bonos), 12):
        panelbonos+=[[bonos[i], bonos[i+1], bonos[i+2], bonos[i+9]]]

    dataframebonos = pd.DataFrame(panelbonos, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón CEDEARS
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[5]/a").click()

    # Extraemos la tabla de CEDEARS
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablacedear = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textcedear = tablacedear.text
    cedear = textcedear.split()

    panelcedear = []
    for i in range(0, len(cedear), 13):
        panelcedear+=[[cedear[i], cedear[i+1], cedear[i+2], cedear[i+9]]]

    dataframecedear = pd.DataFrame(panelcedear, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón OPCIONES
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[2]/a").click()

    # Extraemos la tabla de Opciones
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablaopciones = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textopciones = tablaopciones.text
    opciones = textopciones.split()

    panelopciones = []
    for i in range(0, len(opciones), 13):
        panelopciones+=[[opciones[i], opciones[i+1], opciones[i+2], opciones[i+10]]]

    dataframeopciones = pd.DataFrame(panelopciones, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón ROFEX
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[6]/span").click()

    # Click en el botón FUTUROS ROFEX
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[6]/ul/li[1]/a").click()

    # Extraemos la tabla de Futuros Rofex
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablafutrofex = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textfutrofex = tablafutrofex.text
    futurosrofex = textfutrofex.split()

    panelfuturosrofex = []
    for i in range(0, len(futurosrofex), 16):
        panelfuturosrofex+=[[futurosrofex[i], futurosrofex[i+1], futurosrofex[i+2], futurosrofex[i+11]]]

    dataframefuturosrofex = pd.DataFrame(panelfuturosrofex, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón ROFEX
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[6]/span").click()

    # Click en el botón OPCIONES ROFEX
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[6]/ul/li[2]/a").click()

    # Extraemos la tabla de Opciones Rofex
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablaoprofex = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textoprofex = tablaoprofex.text
    opcionesrofex = textoprofex.split()

    panelopcionesrofex = []
    for i in range(0, len(opcionesrofex), 14):
        panelopcionesrofex+=[[opcionesrofex[i], opcionesrofex[i+1]+opcionesrofex[i+2], opcionesrofex[i+3], opcionesrofex[i+4], opcionesrofex[i+11]]]

    dataframeopcionesrofex = pd.DataFrame(panelopcionesrofex, columns=['Ticker', 'Contratos', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón LETRAS
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[4]/span").click()

    # Click en el botón LETRAS EN PESOS
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[4]/ul/li[1]/a").click()

    # Extraemos la tabla de Letras en Pesos
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablaletraspesos = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textletraspesos = tablaletraspesos.text
    letrasenpesos = textletraspesos.split()

    panelletrasenpesos = []
    for i in range(0, len(letrasenpesos)-12, 12):
        panelletrasenpesos+=[[letrasenpesos[i], letrasenpesos[i+1], letrasenpesos[i+2], letrasenpesos[i+9]]]

    dataframeletrasenpesos = pd.DataFrame(panelletrasenpesos, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón LETRAS
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[4]/span").click()

    # Click en el botón LETRAS EN DÓLARES
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[4]/ul/li[2]/a").click()

    # Extraemos la tabla de Letras en Dólares
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablaletrasdol = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textletrasdol = tablaletrasdol.text
    letrasendolares = textletrasdol.split()

    panelletrasendolares = []
    for i in range(0, len(letrasendolares)-24, 12):
        panelletrasendolares+=[[letrasendolares[i], letrasendolares[i+1], letrasendolares[i+2], letrasendolares[i+9]]]

    dataframeletrasendolares = pd.DataFrame(panelletrasendolares, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])

    # Click en el botón CAUCIONES
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[2]/div/div[3]/div/nav[1]/div[2]/ul/li[10]/a").click()

    # Extraemos la tabla de Cauciones
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "loading")))
    WebDriverWait(browser, 30000).until_not(expected_conditions.visibility_of_element_located((By.ID, "custom_loader")))
    time.sleep(1)
    tablacauciones = browser.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/div/div[1]/div[1]/section/table/tbody[1]')
    textcauciones = tablacauciones.text
    cauciones = textcauciones.split()
    
    panelcauciones = []
    for i in range(0, len(cauciones), 17):
        panelcauciones+=[[cauciones[i]+cauciones[i+1]+cauciones[i+2]+cauciones[i+3]+cauciones[i+4]+cauciones[i+5], cauciones[i+6], cauciones[i+7], cauciones[i+14]]]

    dataframecauciones = pd.DataFrame(panelcauciones, columns=['Ticker', 'Precio', 'Variacion', 'Volumen'])


    with pd.ExcelWriter(ruta) as writer:
        dataframepl.to_excel(writer, sheet_name="Panel Lider")
        dataframepg.to_excel(writer, sheet_name="Panel General")
        dataframebonos.to_excel(writer, sheet_name="Bonos")
        dataframecedear.to_excel(writer, sheet_name="Cedears")
        dataframeopciones.to_excel(writer, sheet_name="Opciones")
        dataframefuturosrofex.to_excel(writer, sheet_name="Futuros Rofex")
        dataframeopcionesrofex.to_excel(writer, sheet_name="Opciones Rofex")
        dataframeletrasenpesos.to_excel(writer, sheet_name="Letras Pesos")
        dataframeletrasendolares.to_excel(writer, sheet_name="Letras Dolares")
        dataframecauciones.to_excel(writer, sheet_name="Cauciones")
    system('cls')
    print("Enviamos los datos")