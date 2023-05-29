import time
import unittest

import allure
import openpyxl # 1. Importante añadir la libreria

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from allure_commons.types import AttachmentType

class Funciones_Globales (): # 2. Se crea la clase Funciones_Excel y se crea la inicialización dentro de esta
    def __init__(self,driver):
        self.driver = driver

    # Función que permite navegar a una URL en concreto
    def Navegar(self,url,tiempo):
        self.driver.get(url) # 3. Espera a que se le envie url a través de la función Navegar
        self.driver.maximize_window()
        self.driver.implicitly_wait(20)
        print("\nPágina abierta de forma satisfactoria:\n"+str(url))
        t = time.sleep(tiempo)
        return t

    def FindElementByXPATH(self, selector):
        val = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, selector)))  # 4. Busca el elemento
        val = self.driver.execute_script("arguments[0].scrollIntoView();",val)  # 5. Si lo encuentra va hacia el elemento
        val = self.driver.find_element(By.XPATH, (selector))  # 6. Selecciona el elemento a través de su XPath
        return val

    def FindElementByID(self, selector):
        val = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, selector)))  # 4. Busca el elemento
        val = self.driver.execute_script("arguments[0].scrollIntoView();",val)  # 5. Si lo encuentra va hacia el elemento
        val = self.driver.find_element(By.ID, (selector))  # 6. Selecciona el elemento a través de su XPath
        return val

    # Función que permite escribir texto, lo anterior a través de la busqueda del selector (XPATH o ID)
    def TextoMixto(self,tipo,selector,texto,tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                val.clear()  # 7. Permite limpiar el campo por si tiene algún valor escrito
                val.send_keys(texto)  # 8. Recibe el valor de texto en la función y lo escribe sobre el elemento
                print("\n- Escribiendo en el campo con {} ({}) el texto -> {}".format(tipo,selector,texto))
                t = time.sleep(tiempo)  # 9. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif(tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                val.clear()  # 13. Permite limpiar el campo por si tiene algún valor escrito
                val.send_keys(texto)  # 14. Recibe el valor de texto en la función y lo escribe sobre el elemento
                print("\n- Escribiendo en el campo con {} ({}) el texto -> {}".format(tipo,selector,texto))
                t = time.sleep(tiempo)  # 15. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite hacer clic en un elemento en concreto, lo anterior a través de la busqueda del selector (XPATH o ID)
    def ClickMixto(self,tipo,selector,tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                val.click()
                print("\n- Se da click en el campo con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                val.click()
                print("\n- Se da click en el campo con {} -> ({})".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite hacer la busqueda y selección de un elemento de una lista desplegable, lo anterior a través de
    # su elemento (XPATH o ID)
    def Select_Mixto_Text(self, tipo, selector, text, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                val = Select(val)
                val.select_by_visible_text(text) # 16. Busca y selecciona el valor de text en el listado
                print("\n- El campo seleccionado es -> ({})".format(text))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                val = Select(val)
                val.select_by_visible_text(text) # 16. Busca y selecciona el valor de text en el listado
                print("\n- El campo seleccionado es -> ({})".format(text))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite hacer la busqueda y selección de un elemento de una lista desplegable (text, index o value),
    # lo anterior a través de su elemento (XPATH o ID)
    def Select_Mixto_Type(self, tipo, selector, type, dato, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                val = Select(val)

                if(type == "text"):
                    val.select_by_visible_text(dato)  # 16. Busca y selecciona el valor de text en el listado

                elif(type == "index"):
                    val.select_by_index(dato)

                elif(type == "value"):
                    val.select_by_value(dato)

                print("\n- Se realizó la busqueda a través de ({}) y el campo seleccionado es equivalente a -> ({})".format(type,dato))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                val = Select(val)

                if (type == "text"):
                    val.select_by_visible_text(dato)  # 16. Busca y selecciona el valor de text en el listado

                elif (type == "index"):
                    val.select_by_index(dato)

                elif (type == "value"):
                    val.select_by_value(dato)

                print("\n- Se realizó la busqueda a través de ({}) y el campo seleccionado es equivalente a -> ({})".format(type,dato))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite subir un archivo especifico de una ruta del computador, lo anterior filtrando el campo a través de
    # su elemento (XPATH o ID)
    def Upload_Mixto(self, tipo, selector, ruta, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                val.send_keys(ruta)
                print("\n- Se carga el archivo de forma satisfactoria a través de la busqueda en la ruta -> ({})".format(ruta))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                val.send_keys(ruta)
                print("\n- Se carga el archivo de forma satisfactoria a través de la busqueda en la ruta -> ({})".format(ruta))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Esta función permite hacer click en varios campos de un formulario, sin importar la cantidad de argumentos que se envie
    def Clic_Multiple_Mixto(self, tipo, tiempo=.2, *args):
        if (tipo == "XPATH"):
            try:
                for num in args: # 13. Recorre la cantidad de argumentos y se almacena este valor dentro de num
                    val = self.FindElementByXPATH(num)
                    val.click()
                    print("\n- Se da click en el campo con ({}) -> ({})".format(tipo,num))
                    t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                    return t

            except TimeoutException as ex:
                for num in args:
                    print(ex.msg)
                    print("\n- No se encontró el elemento:\n- " + num)

        elif (tipo == "ID"):
            try:
                for num in args:  # 13. Recorre la cantidad de argumentos y se almacena este valor dentro de num
                    val = self.FindElementByID(num)
                    val.click()
                    print("\n- Se da click en el campo con ({}) -> ({})".format(tipo,num))
                    t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                    return t

            except TimeoutException as ex:
                for num in args:
                    print(ex.msg)
                    print("\n- No se encontró el elemento:\n- " + num)

    # Función que valida si el nombre se cargo de forma correcta tras dar clic en el botón y lo muestra en el resultado por pantalla
    def FindExistMixto(self, tipo, selector, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                print("\n- El elemento de tipo ({}) y nombre ({}) -> Existe".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return "Existe" # 17. Retorna el valor de que se encontró el elemento y por ende existe

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return "No existe"

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                print("\n- El elemento de tipo ({}) y nombre ({}) -> Existe".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return "Existe"  # 17. Retorna el valor de que se encontró el elemento y por ende existe

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return "No existe" # 18. Retorna el valor de que no se encontró el elemento y por ende NO existe

    # Función que realiza una captura de pantalla
    def Screenshot(self, name, tiempo=.2):
        self.driver.get_screenshot_as_file("C:/Users/alexf/Desktop/Automatización Selenium Python/Ejercicios/img/scr/{}.png".format(name))
        print("\nLa Captura ha sido guardada con el nombre -> ({}.png)".format(name))
        t = time.sleep(tiempo)
        return t

    #Función que permite dar doble clic al elemento seleccionado, lo anterior a través de la busqueda del XPATH o ID
    def ClickDoble(self,tipo,selector,tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                act = ActionChains(self.driver)
                act.double_click(val).perform()
                print("\n- Se da doble click en el campo con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                act = ActionChains(self.driver)
                act.double_click(val).perform()
                print("\n- Se da doble click en el campo con {} -> ({})".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite dar clic derecho al elemento seleccionado, lo anterior a través de la busqueda del XPATH o ID
    def ClickDerecho(self, tipo, selector, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                act = ActionChains(self.driver)
                act.context_click(val).perform()
                print("\n- Se da click derecho en el campo con {} -> ({})".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                act = ActionChains(self.driver)
                act.context_click(val).perform()
                print("\n- Se da click derecho en el campo con {} -> ({})".format(tipo, selector))
                t = time.sleep(tiempo)  # 16. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite agarrar un objeto y soltarlo en una ruta especifica
    def DragAndDrop(self, tipo, selector, selectordestino, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                destino = self.FindElementByXPATH(selectordestino) # 15. Se almacena el XPATH del elemento al que se va a soltar
                act = ActionChains(self.driver)
                act.drag_and_drop(val,destino).perform() # 16. Se especifica el elemento que toma y el elemento donde lo va a arrastrar
                print("\n- Se selecciono el elemento ({}) de tipo {} y se soltó en -> ({})".format(selector,tipo,selectordestino))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                destino = self.FindElementByID(selectordestino) # 15. Se almacena el ID del elemento al que se va a soltar
                act = ActionChains(self.driver)
                act.drag_and_drop(val,destino).perform() # 16. Se especifica el elemento que toma y el elemento donde lo va a arrastrar
                print("\n- Se selecciono el elemento ({}) de tipo {} y se soltó en -> ({})".format(selector,tipo,selectordestino))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite agarrar un elemento, moverlo y soltarlo en la posición que se especifica (X,Y)
    def DragAndDropXY(self, tipo, selector, x, y, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                self.driver.switch_to.frame(0) # 15. Se restablece la posición el iframe para que permita posicionarse en el X,Y seleccionado
                val = self.FindElementByXPATH(selector)
                act = ActionChains(self.driver)
                act.drag_and_drop_by_offset(val,x,y).perform() # 16. Se especifica el elemento que toma y donde lo va a soltar (X,Y)
                print("\n- Se selecciono el elemento ({}) de tipo {} y se soltó en las coordenadas (X,Y)-> ({},{})".format(selector,tipo,x,y))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                self.driver.switch_to.frame(0)
                val = self.FindElementByID(selector)
                act = ActionChains(self.driver)
                act.drag_and_drop_by_offset(val,x,y).perform() # 16. Se especifica el elemento que toma y el elemento donde lo va a arrastrar
                print("\n- Se selecciono el elemento ({}) de tipo {} y se soltó en las coordenadas (X,Y)-> ({},{})".format(selector,tipo,x,y))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite buscar un elemento (XPATH, ID), y a partir de este movernos en (X,Y) y dar clic en esta posición
    def ClickXY(self, tipo, selector, x, y, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                # self.driver.switch_to.frame(0) No es necesario a menos que tenga un iFrame (Inspeccionar y validar)
                val = self.FindElementByXPATH(selector)
                act = ActionChains(self.driver)
                act.move_to_element_with_offset(val,x,y).click().perform() # 16. Pone el puntero en la posición X,Y especificada y luego da click
                print("\n- Se dio click al elemento ({}) de tipo {}, lo anterior en las siguientes coordenadas (X,Y)\n-> ({},{})".format(selector,tipo,x,y))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                # self.driver.switch_to.frame(0) No es necesario a menos que tenga un iFrame (Inspeccionar y validar)
                val = self.FindElementByID(selector)
                act = ActionChains(self.driver)
                act.move_to_element_with_offset(val,x,y).click().perform() # 16. Pone el puntero en la posición X,Y especificada y luego da click
                print("\n- Se dio click al elemento ({}) de tipo {}, lo anterior en las siguientes coordenadas (X,Y)\n-> ({},{})".format(selector,tipo,x,y))
                t = time.sleep(tiempo)  # 17. Recibe el valor de tiempo y lo implementa una vez que se escribe el texto
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite copiar un elemento (XPATH, ID)
    def Copiar(self, tipo, selector, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                self.ClickMixto("XPATH", selector, tiempo)
                act = ActionChains(self.driver)
                act.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform() # 4. El ctrl + a permite seleccionar toda una palabra
                act.key_down(Keys.CONTROL).send_keys("c").key_up(Keys.CONTROL).perform() # 4. El ctrl + c permite copiar la palabra
                print("\n- Se copió de forma satisfactoria el contenido del elemento con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                self.ClickMixto("ID",selector,tiempo)
                act = ActionChains(self.driver)
                act.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()  # 4. El ctrl + a permite seleccionar toda una palabra
                act.key_down(Keys.CONTROL).send_keys("c").key_up(Keys.CONTROL).perform()  # 4. El ctrl + c permite copiar la palabra
                print("\n- Se copió de forma satisfactoria el contenido del elemento con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite pegar un elemento (XPATH, ID), lo anterior tras haberlo copiado con antelación
    def Pegar(self, tipo, selector, tiempo=.2):
        if (tipo == "XPATH"):
            try:
                val = self.FindElementByXPATH(selector)
                self.ClickMixto("XPATH", selector, tiempo)
                act = ActionChains(self.driver)
                act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()  # 4. El ctrl + v permite pegar
                print("\n- Se pegó de forma satisfactoria el contenido del en la ruta con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

        elif (tipo == "ID"):
            try:
                val = self.FindElementByID(selector)
                self.ClickMixto("ID", selector, tiempo)
                act = ActionChains(self.driver)
                act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()  # 4. El ctrl + v permite pegar
                print("\n- Se pegó de forma satisfactoria el contenido del en la ruta con {} -> ({})".format(tipo,selector))
                t = time.sleep(tiempo)
                return t

            except TimeoutException as ex:
                print(ex.msg)
                print("\n- No se encontró el elemento:\n- " + selector)
                return t

    # Función que permite bajar el scroll y llegar al final de la página
    def ScrollPageDown(self, tiempo=.2):
        for x in range (10):
            self.driver.find_element(By.TAG_NAME, value="body").send_keys(Keys.PAGE_DOWN)
        t = time.sleep(tiempo)
        return t

    # Función que permite subir el scroll y llegar al inicio de la página
    def ScrollPageUp(self, tiempo=.2):
        for x in range(10):
            self.driver.find_element(By.TAG_NAME, value="body").send_keys(Keys.PAGE_UP)
        t = time.sleep(tiempo)
        return t

    # Función que permite bajar el scroll y llegar al final de la página
    def ScreenshotAllure(self, name, tiempo=.2):
        allure.attach(self.driver.get_screenshot_as_png(), name=name, attachment_type=AttachmentType.PNG)
        t = time.sleep(tiempo)

    # Función que brinda un mensaje de finalización exitoso
    def Salida(self):
        print("\nLa prueba ha sido finalizada de forma satisfactoria.")