import pandas as pd
import time
import random
import os

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class getInfoSimit:

    def __init__(self):
        service = Service(executable_path = 'msedgedriver.exe')

        options = webdriver.EdgeOptions()
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-webgl')
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-logging')

        self.driver = webdriver.Edge(service=service, options=options)
        ruta_archivo = 'base.xlsx'

        self.driver.get("https://www.fcm.org.co/simit/#/home-public")

        if not os.path.isfile(ruta_archivo):
            print(f"Error: El archivo '{ruta_archivo}' no se encuentra.")
            raise Exception("Por favor asegúrese de que el archivo esté en la misma carpeta que el script de Python.")

        self.base = pd.read_excel(ruta_archivo)

        columnas_archivo = self.base.columns.tolist()
        columnas_esperadas = ["PLACA"]

        if columnas_esperadas in columnas_archivo:
            print(f"Columnas encontradas: {columnas_archivo}")
            print(f"Columnas necesitada: {columnas_esperadas}")
            raise Exception("No se presenta la columna necesaria")
        
        self.resultados = []
        self.detallesMultas = []
    
    def getDetailsRow(self, row, placa):
        try:
            detalleMulta = []
            estado = row.find_element(By.XPATH, './/td[@data-label="Estado"]').text
            valor = row.find_element(By.XPATH, './/td[@data-label="Valor"]').text
            suma = row.find_element(By.XPATH, './/td[@data-label="Valor a pagar"]').text.replace("Detalle Pago","")
            
            if "Interés " in valor:
                valor = valor.split("Interés ")
            else:
                valor = [valor, 0]

            detalleMulta.append(placa)
            detalleMulta.append(estado)
            detalleMulta.append(valor[0])
            detalleMulta.append(valor[1])
            detalleMulta.append(suma)

            boton = row.find_element(By.XPATH, './/td[@data-label="Tipo"]').find_element(By.TAG_NAME, "a")
            self.driver.execute_script("arguments[0].click();", boton)
            time.sleep(2)
            return detalleMulta
        except Exception as e:
            time.sleep(random.uniform(5,7))
            return self.getDetailsRow(row, placa)
    
    def getDetailsExact(self, detalleMulta):
        try:
            detalleMulta = detalleMulta[:5]
            agregar = self.driver.find_element(By.CLASS_NAME, "card-body.p-3").find_elements(By.XPATH, './/p[@class="mb-0"]')
            
            for i in range(23):
                if i in {2, 9, 15}:
                    continue
                elif i == 1:
                    detalleMulta.append(agregar[i].text+" "+agregar[i+1].text)
                else:
                    detalleMulta.append(agregar[i].text)

            self.detallesMultas.append(detalleMulta)
            buttonBack = self.driver.find_element(By.XPATH, '//button[text()="Volver"]')
            self.driver.execute_script("arguments[0].scrollIntoView();", buttonBack)
            time.sleep(random.uniform(0.1,0.3))
            buttonBack.click()
            time.sleep(2)
        except Exception as e:
            time.sleep(random.uniform(5,7))
            self.getDetailsExact(detalleMulta)

    def buscarPlaca(self, placa):
        try:
            busqueda = self.driver.find_element(By.ID, "txtBusqueda")
            busqueda.clear()
            time.sleep(random.uniform(0.4,0.7))
            busqueda.send_keys(placa)
            time.sleep(random.uniform(0.4,0.7))
            try:
                self.driver.find_element(By.ID, "consultar").click()
            except Exception as e:
                self.driver.find_element(By.ID, "btnNumDocPlaca").click()
        except Exception as e:
            time.sleep(random.uniform(5,6.5))
            self.buscarPlaca(placa)

    def getStatusTrueFalse(self, placa):
        try:
            try:
                WebDriverWait(self.driver, 4).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'No tienes comparendos ni multas registradas en Simit')]"))
                )
                self.resultados.append([placa,"No",0,0,0,0])
                time.sleep(random.uniform(0.5,0.7))
                return False
            except Exception as e:
                resumen = self.driver.find_element(By.XPATH, "//*[@id='resumenEstadoCuenta']/div/div")
                datos = resumen.find_elements(By.TAG_NAME,"strong")
                textos = [dato.text for dato in datos]
                self.resultados.append([placa,"Si",textos[0] or 0,textos[1] or 0,textos[2] or 0,textos[3] or 0])
                return True
        except Exception as e:
            time.sleep(random.uniform(5,6.5))
            return self.getStatusTrueFalse(placa)

    def getStatusMulta(self, placa):
        if(self.getStatusTrueFalse(placa)):
            for indice in range(len(self.driver.find_elements(By.XPATH, "//table[@id='multaTable']//tbody/tr"))):
                time.sleep(random.uniform(0.2,0.5))
                rows = self.driver.find_elements(By.XPATH, "//table[@id='multaTable']//tbody/tr")

                if indice < len(rows):
                    fila = rows[indice]
                    detalleMulta = self.getDetailsRow(fila, placa)
                    self.getDetailsExact(detalleMulta)

    def saveInfo(self):
        self.driver.quit()
        multas = pd.DataFrame(self.resultados, columns=["PLACA","INFRACCIONES","COMPARENDOS","MULTAS","ACUERDOS DE PAGO", "TOTAL"])
        detalles = pd.DataFrame(self.detallesMultas, columns= ["Placa","Estado","Valor","Interes","Valor Total","No. comparendo","Fecha","Direccion","Fuente infracción","Secretaría","Agente","Código", "Descripcion","Tipo documento","Numero Documento", "Nombres","Apellidos","Tipo Infractor","No. Licencia Vehiculo", "Tipo","Servicio", "No.Licencia", "Fecha Vencimiento", "Categoría", "Secretaria"])
        with pd.ExcelWriter("MULTAS.xlsx", engine="openpyxl") as writer:
            multas.to_excel(writer, sheet_name="Multas", index=False)
            detalles.to_excel(writer, sheet_name="Detalles", index=False)
    
    def app(self):
        try:
            self.driver.find_element(By.CLASS_NAME, "close.modal-info-close").click()
        except Exception as e:
            time.sleep(random.uniform(5,7))
            self.app()

        for placa in self.base["PLACA"].unique():
            self.buscarPlaca(placa)
            self.getStatusMulta(placa)

        self.saveInfo()

app = getInfoSimit()
app.app()