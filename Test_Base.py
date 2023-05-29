import time
import unittest

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from Funciones import Funciones_Globales
from Funciones_Excel import Funciones_Excel

t=0.1

class baseExcel(unittest.TestCase):
    def setUp(self):
        ser = Service("C:\Drivers\chromedriver.exe")
        op = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(service=ser, options=op)

    def test_Excel(self): # 1. Es importante que se añada el "test_name" porque sino nos va a dar error
        driver = self.driver
        fg = Funciones_Globales(driver)
        fe = Funciones_Excel(driver)
        ruta = "C:/Users/alexf/Desktop/Selenium_Excel_Base/docs/Datos_Excel.xlsx"
        filas = fe.getRowCount(ruta,"Hoja1") # 2. Se almacena dentro de la variable y luego se llama a través de la función
        fg.Navegar("http://computer-database.gatling.io/computers", t)

        for r in range (2,filas+1): # 3. Se empieza desde 2 ya que la primera es de los nombres de los valores
            fg.Navegar("http://computer-database.gatling.io/computers/new", t)
            computerName = fe.readData(ruta,"Hoja1",r,1) # 4. Toma el valor del nombre del Excel y lo guarda dentro de la variable
            introduced = fe.readData(ruta,"Hoja1",r,2)
            discontuined = fe.readData(ruta,"Hoja1",r,3) # 5. Se va moviendo campos hacia la derecha y los va almacenando en las variables
            company = fe.readData(ruta,"Hoja1",r,4)

            fg.TextoMixto("ID","name",computerName,t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_1_computerName", t)
            fg.TextoMixto("ID","introduced",introduced,t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_2_introduced", t)
            fg.TextoMixto("ID", "discontinued", discontuined, t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_3_discontuined", t)
            fg.Select_Mixto_Type("ID","company","text",company,t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_4_company", t)
            fg.ClickMixto("XPATH","//input[contains(@type,'submit')]",t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_5_Validate", t)

            e = fg.FindExistMixto("XPATH","//div[contains(@class,'alert-message warning')]",t) # 6. Busca el computerName a través de su ID y si existe lo indica

            if (e == "Existe"): #7. Valida el valor que retorna la función
                print("\n Se escribió de forma satisfactoria en el archivo que se encuentra en la ruta\n-> " +ruta)
                # 8. Se especifica que en el campo 5 se debe escribir el mensaje correspondiente tras hacer una inserción
                fe.writeData(ruta,"Hoja1",r,5,"Se insertó de forma satisfactoria")

            else:
                print("\n No se insertó de forma satisfactoria")
                fe.writeData(ruta, "Hoja1", r, 5, "Error al realizar la inserción de la información")

if __name__ == '__main__':
    unittest.main()