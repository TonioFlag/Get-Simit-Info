import pandas as pd
import time
import random

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

service = Service(executable_path = 'msedgedriver.exe')
driver = webdriver.Edge(service=service)

link = "https://www.fcm.org.co/simit/#/estado-cuenta?numDocPlacaProp="

base = pd.read_excel("base.xlsx")
driver.get("https://www.fcm.org.co/simit/#/home-public")
time.sleep(10)
driver.find_element(By.CLASS_NAME, "close.modal-info-close").click()

resultados = []
detallesMultas = []

def getDataDetail(row):
    detalleMulta = []
    estado = row.find_element(By.XPATH, './/td[@data-label="Estado"]').text
    valor = row.find_element(By.XPATH, './/td[@data-label="Valor"]').text
    suma = row.find_element(By.XPATH, './/td[@data-label="Valor a pagar"]').text.replace("Detalle Pago","")
    
    if "Interés " in valor:
        valor = valor.split("Interés ")
    else:
        valor = [valor, 0]

    detalleMulta.append(i)
    detalleMulta.append(estado)
    detalleMulta.append(valor[0])
    detalleMulta.append(valor[1])
    detalleMulta.append(suma)

    boton = row.find_element(By.XPATH, './/td[@data-label="Tipo"]').find_element(By.TAG_NAME, "a")
    driver.execute_script("arguments[0].click();", boton)
    time.sleep(2)

    agregar = driver.find_element(By.CLASS_NAME, "card-body.p-3").find_elements(By.XPATH, './/p[@class="mb-0"]')
    detalleMulta.append(agregar[0].text)
    detalleMulta.append(agregar[1].text+" "+agregar[2].text)
    detalleMulta.append(agregar[3].text)
    detalleMulta.append(agregar[4].text)
    detalleMulta.append(agregar[5].text)
    detalleMulta.append(agregar[6].text)

    detalleMulta.append(agregar[7].text)
    detalleMulta.append(agregar[8].text)

    detalleMulta.append(agregar[10].text)
    detalleMulta.append(agregar[11].text)
    detalleMulta.append(agregar[12].text)
    detalleMulta.append(agregar[13].text)
    detalleMulta.append(agregar[14].text)

    detalleMulta.append(agregar[16].text)
    detalleMulta.append(agregar[17].text)
    detalleMulta.append(agregar[18].text)
            
    detalleMulta.append(agregar[19].text)
    detalleMulta.append(agregar[20].text)
    detalleMulta.append(agregar[21].text)
    detalleMulta.append(agregar[22].text)

    detallesMultas.append(detalleMulta)
    buttonBack = driver.find_element(By.XPATH, '//button[text()="Volver"]')
    driver.execute_script("arguments[0].scrollIntoView();", buttonBack)
    time.sleep(0.3)
    buttonBack.click()
    time.sleep(2)

for i in base["PLACA"].unique():
    busqueda = driver.find_element(By.ID, "txtBusqueda")
    busqueda.clear()
    time.sleep(0.7)
    busqueda.send_keys(i)

    time.sleep(0.6)
    try:
        driver.find_element(By.ID, "consultar").click()
    except Exception as e:
        driver.find_element(By.ID, "btnNumDocPlaca").click()

    try:
        existencia = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'No tienes comparendos ni multas registradas en Simit')]"))
        )
        resultados.append([i,"No",0,0,0,0])
        time.sleep(0.5)
    except Exception as e:
        resumen = driver.find_element(By.XPATH, "//*[@id='resumenEstadoCuenta']/div/div")
        datos = resumen.find_elements(By.TAG_NAME,"strong")
        textos = [dato.text for dato in datos]
        resultados.append([i,"Si",textos[0] or 0,textos[1] or 0,textos[2] or 0,textos[3] or 0])

        for indice in range(len(driver.find_elements(By.XPATH, "//table[@id='multaTable']//tbody/tr"))):
            time.sleep(0.2)
            rows = driver.find_elements(By.XPATH, "//table[@id='multaTable']//tbody/tr")

            if indice < len(rows):
                fila = rows[indice]
                getDataDetail(fila)
                    
driver.quit()

multas = pd.DataFrame(resultados, columns=["PLACA","INFRACCIONES","COMPARENDOS","MULTAS","ACUERDOS DE PAGO", "TOTAL"])
detallesMultas = pd.DataFrame(detallesMultas, columns= ["Placa","Estado","Valor","Interes","Valor Total","No. coparendo","Fecha","Direccion","Fuente infracción","Secretaría","Agente","Código", "Descripcion","Tipo documento","Numero Documento", "Nombres","Apellidos","Tipo Infractor","No. Licencia Vehiculo", "Tipo","Servicio", "No.Licencia", "Fecha Vencimiento", "Categoría", "Secretaria"])

with pd.ExcelWriter("MULTAS.xlsx", engine="openpyxl") as writer:
    multas.to_excel(writer, sheet_name="Multas", index=False)
    detallesMultas.to_excel(writer, sheet_name="Detalles", index=False)