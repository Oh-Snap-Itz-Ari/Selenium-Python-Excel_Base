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
        ruta = "C:/Users/alexf/Desktop/Automatización Selenium Python/Ejercicios/Ejercicios_Modulo_6_Excel/docs/Datos_Excel.xlsx"
        filas = fe.getRowCount(ruta,"Hoja1") # 2. Se almacena dentro de la variable y luego se llama a través de la función

        for r in range (2,filas+1): # 3. Se empieza desde 2 ya que la primera es de los nombres de los valores
            fg.Navegar("https://demoqa.com/text-box", t)
            name = fe.readData(ruta,"Hoja1",r,1) # 4. Toma el valor del nombre del Excel y lo guarda dentro de la variable
            email = fe.readData(ruta,"Hoja1",r,2)
            addr1 = fe.readData(ruta,"Hoja1",r,3) # 5. Se va moviendo campos hacia la derecha y los va almacenando en las variables
            addr2 = fe.readData(ruta,"Hoja1",r,4)

            fg.TextoMixto("XPATH","//input[@id='userName']",name,t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_1_Name", t)
            fg.TextoMixto("XPATH","//input[@id='userEmail']",email,t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_2_Email", t)
            fg.TextoMixto("XPATH", "//textarea[@id='currentAddress']", addr1, t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_3_CurrentAddress", t)
            fg.TextoMixto("XPATH", "//textarea[@id='permanentAddress']", addr2, t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_4_Permanet_address", t)
            fg.ClickMixto("XPATH","//button[@id='submit']",t)
            fg.Screenshot("TC_0"+str(r-1)+" - Campo_5_Validate", t)

            e = fg.FindExistMixto("XPATH","//p[@id='email']",t) # 6. Busca el correo a través de su XPATH y si existe lo indica

            if (e == "Existe"): #7. Valida el valor que retorna la función
                print("\n Se escribió de forma satisfactoria en el archivo que se encuentra en la ruta\n-> " +ruta)
                # 8. Se especifica que en el campo 5 se debe escribir el mensaje correspondiente tras hacer una inserción
                fe.writeData(ruta,"Hoja1",r,5,"Se insertó de forma satisfactoria")

            else:
                print("\n No se insertó de forma satisfactoria")
                fe.writeData(ruta, "Hoja1", r, 5, "Error al realizar la inserción de la información")

if __name__ == '__main__':
    unittest.main()