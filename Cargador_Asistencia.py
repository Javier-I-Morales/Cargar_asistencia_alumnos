import time
import xlrd
import xlwt
import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select



class cargador:

    def __init__(self):

        self.usuario = None
        self.passw = None
        self.desde = None
        self.hasta = None
        self.ruta_archivo = None
        self.ruta_save = None
        self.fecha = None
        self.columna = 0
        self.documentos_filtrados = None

        dia = str(time.localtime().tm_mday)
        mes = str(time.localtime().tm_mon)
        hora = str(time.localtime().tm_hour)
        min = str(time.localtime().tm_min)



        logging.basicConfig(filename="logapp"+dia+mes+hora+min+".log",level='DEBUG')


    def cantidad_para_procesar(self):
        return len(self.documentos_filtrados)


    def cargar_asistencias(self):

        try:

            logging.info("Iniciando la ejecución de la carga de asistencias con \nUsuario:"+self.usuario+"\nPassword:"+self.passw+"\nDesde comisión:"+self.desde+"\nHasta comisión:"+self.hasta+"\nFecha:"+self.fecha+"\n\n******************************************************************************************\n")

            # usuario = str(input("Ingrese su nombre de usuario: "))
            # passw = str(input("Ingrese su password: "))
            # desde = int(input("Ingrese desde que comisión: "))
            # hasta = int(input("Ingrese hasta que comisión: "))

            # print("IMPORTANTE, EL NUMERO DE LA HOJA DEBE COINCIDIR CON EN NUMERO DE COMISIÓN")
            # filepath = str(input("Ingrese la ruta del archivo (incluyendo al archivo con la extención .xlsx) con las hojas previamente confeccionadas: "))
            nombreHoja = "Hoja1"
            columna = 2 #int(input("Indique la posición de la columna donde se encuentran los documentos: "))
            # fecha = str(input("Ingrese la fecha (dd/mm/aaaa): "))

            posicionComisiones = {1:"0",2:"11",3:"22",4:"27",5:"28",6:"29",7:"30",8:"31",9:"32",10:"1",11:"2",12:"3",13:"4",14:"5",15:"6",16:"7",17:"8",18:"9",19:"10",20:"12",21:"13",22:"14",23:"15",24:"16",25:"17",26:"18",27:"19",28:"20",29:"21",30:"23",31:"24",32:"25",33:"26"}
            openfile = xlrd.open_workbook(self.ruta_archivo)
            libro = xlwt.Workbook()

            # chars = '/'
            # mifecha = self.fecha.translate(str.maketrans('','',chars))


            # ######################################################################################################
            # # DESDE ACA PARA ABRIR LA PAGINA Y LLEGAR HASTA LA HOJA DE ASISTENCIA DESEADA
            #
            # opcion = Options()


            opcion = webdriver.ChromeOptions()

            opcion.add_argument("start-maximized")

            # opcion.add_argument('--ignore-certificate-errors-spki-list')
            # opcion.add_argument('--ignore-ssl-errors')
            # opcion.add_argument('--allow-running-insecure-content')
            opcion.add_argument("--disable-blink-features=AutomationControlled")

            opcion.add_experimental_option('excludeSwitches', ['enable-logging'])

            opcion.accept_untrusted_certs = True




            #caps = webdriver.DesiredCapabilities.CHROME.copy()
            #caps['acceptInsecureCerts'] = True


            driver = webdriver.Chrome(executable_path='F:\\Python\\Cargar_asistencia_alumnos\\chromedriver.exe', chrome_options=opcion)#, desired_capabilities=caps
            driver.get('https://servicios.unahur.edu.ar/unahur3w/')


            logging.info("LLego hasta sector: A")

            time.sleep(3)
            user = driver.find_element("xpath",'//*[@id="usuario"]')
            user.send_keys(self.usuario)
            logging.info("LLego hasta sector: B")
            password = driver.find_element("xpath",'//*[@id="password"]')
            password.send_keys(self.passw)

            boton_ingresar = driver.find_element("xpath",'/html/body/div[6]/div/div[1]/div/form/div[3]/div/input')
            boton_ingresar.send_keys(Keys.ENTER)
            time.sleep(2)

            boton_perfil_1 = driver.find_element("xpath",'//*[@id="js-selector-perfiles"]/li/a/span')
            boton_perfil_1.click()
            time.sleep(1)
            boton_perfil_2 = driver.find_element("xpath",'//*[@id="js-select-perfil"]/li[2]/a')
            boton_perfil_2.click()
            time.sleep(2)
            boton_clase = driver.find_element("xpath",'//*[@id="zona_clases"]/a')
            boton_clase.click()
            time.sleep(2)
            boton_primer_comision = driver.find_element("xpath",'//*[@id="comisiones"]/div/fieldset/div/table/tbody/tr[1]/td[1]/a')
            boton_primer_comision.click()
            time.sleep(2)
            boton_asistencia = driver.find_element("xpath",'//*[@id="home"]/div[2]/div/table[1]/tbody/tr[1]/td[8]/a[1]')
            boton_asistencia.click()
            time.sleep(2)
            ####################------ HASTA ACÁ SE LLEGA A LA PANTALLA DE ASISTENCIAS ---------------------------------------------





            ####################################################################################################################

            for i in range(int(self.desde), int(self.hasta)+1):
                myselectComision = Select(driver.find_element("xpath",'//*[@id="zona"]/div[1]/ul[1]/select'))
                time.sleep(1)
                myselectComision.select_by_value(posicionComisiones[i])
                print("Procesando comisión: "+str(i))
                time.sleep(2)

                select = Select(driver.find_element("xpath",'//*[@id="cabecera"]/div[1]/div[1]/ul/select'))
                #select.select_by_visible_text(self.fecha)
                opciones = select.options

                for opc in opciones:
                    if self.fecha in str(opc.text):

                        texto = str(opc.text)
                        select.select_by_visible_text(texto)
                        time.sleep(2)
                        presentes = driver.find_element('id','js-asistencia-presentes').text

                        opciones.clear()

                        totalalumnos = driver.find_element('id','js-asistencia-total').text
                        print("Total alumnos en la pagina: " + totalalumnos)

                        sheet = openfile.sheet_by_name(nombreHoja)
                        documentos = []
                        otros = []
                        for j in range(sheet.nrows):
                            # dato = sheet.cell_value(j, 0)
                            # data = xlrd.xldate_as_tuple(dato,libro.dates_1904)
                            # print(str(str(data[2])+'/'+str(data[1])+'/'+str(data[0])) +"|" + str(self.fecha))

                            if sheet.cell_value(j, columna + 1) == "C" + str(i):
                                # print("entro " + data + " | " + self.fecha)
                                if not (type(sheet.cell_value(j, columna)) == str):
                                    documentos.append(str(int(sheet.cell_value(j, columna))))
                                else:
                                    otros.append(sheet.cell_value(j, columna))
                        documentosFiltrados = list(set(documentos))
                        print(otros)
                        print("Total documentos en el archivo: " + str(len(documentosFiltrados)))

                        for h in range(int(totalalumnos)):
                            doc = driver.find_element("xpath",'//*[@id="edicion_asistencias"]/form/div[' + str(h + 1) + ']/div[2]/div[2]')
                            # print(doc.text)

                            if (doc.text in documentosFiltrados):
                                check = driver.find_element("xpath",'//*[@id="edicion_asistencias"]/form/div[' + str(h + 1) + ']')
                                check.click()
                                documentosFiltrados.remove(doc.text)

                        libro1 = libro.add_sheet("Comisión " + str(i))
                        libro1.write(0, 0, "Apellido")
                        libro1.write(0, 1, "Nombre")
                        libro1.write(0, 2, "Documento")

                        for o in otros:
                            documentosFiltrados.append(o)

                        if len(documentosFiltrados) > 0:
                            cont = 1
                            for dni in documentosFiltrados:
                                do = str(dni)
                                print(do)
                                for j in range(1, sheet.nrows):
                                    if type(sheet.cell_value(j, columna)) == float:
                                        dato = str(int(sheet.cell_value(j, columna)))
                                        if do == dato:
                                            libro1.write(cont, 0, sheet.cell_value(j, 0))
                                            libro1.write(cont, 1, sheet.cell_value(j, 1))
                                            libro1.write(cont, 2, sheet.cell_value(j, 2))
                                            cont = cont + 1
                                    else:
                                        if do in str(sheet.cell_value(j, columna)):
                                            libro1.write(cont, 0, sheet.cell_value(j, 0))
                                            libro1.write(cont, 1, sheet.cell_value(j, 1))
                                            libro1.write(cont, 2, sheet.cell_value(j, 2))
                                            cont = cont + 1

                        print("Documentos de comisión n° " + str(i) + " no encontrados: " + str(
                            documentosFiltrados) + ", de la hoja: " + nombreHoja)
                        time.sleep(1)
                        boton_guardar = driver.find_element('id','js-guardar-asistencia')
                        boton_guardar.click()
                        time.sleep(2)
            libro.save(self.ruta_save+"No_Encontrados_desde_C"+self.desde+"_hasta_C"+self.hasta+".xls")
            logging.info("Proceso completado con éxito.")



        except Exception as argument:
            logging.exception(argument.args)

    ####################################  FIN DEL CODIGO  #################################################################################