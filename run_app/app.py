import os, re
from dotenv import load_dotenv
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
#Selenium imports
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import traceback
import pandas as pd
import time
from openpyxl import Workbook, load_workbook

load_dotenv()

class ConsultTrainees:

    def __init__(self, user=None, password=None) -> None:

         #Carpeta raiz
        self.current_root_path = os.getcwd()
        
        #Environ Variables
        self.default_url = os.environ.get("DEFAULT_URL")
        self.default_user = os.environ.get("DEFAULT_USER")
        self.default_pass = os.environ.get("DEFAULT_PASS")

        #Initialize webdriver
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--window-size=700,450")
        prefs = {"download.default_directory" : f"{self.current_root_path}/procesar/"}
        self.options.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
        self.wait = WebDriverWait(self.driver, 50)

       
        
        principal_folders = ["/entrada/", '/procesar/', '/entregables/']
        for folder in principal_folders: # Se crean las carpetas necesarias
            if not os.path.exists(f"{self.current_root_path}{folder}"):
                os.mkdir(f"{self.current_root_path}{folder}")
        self.path_process_folder = f"{self.current_root_path}/procesar/"
        
        self.df_consolidated = pd.DataFrame()

    def open_webdriver(self, url=None):
        try:
            if not url:
                url = self.default_url
            self.driver.get(url)
            self.driver.maximize_window()
            self.driver.switch_to.frame(0)
        except Exception as err:
            print(err)
            print(traceback.format_exc())
    

    def login_process(self, user=None, password=None):
        try:
            if not user and not password:
                user =  self.default_user
                password = self.default_pass

            try:
                element = self.wait.until(EC.element_to_be_clickable((By.ID, "username")))
                element.click()
                element.send_keys(user)
                element = self.wait.until(EC.element_to_be_clickable((By.NAME, "josso_password")))
                element.click()
                element.send_keys(password)
                element = self.wait.until(EC.element_to_be_clickable((By.NAME, "ingresar")))
                element.click()
                print('Loggeo exitoso')
                return True
            except Exception as err:
                print("Error al intentar realizar logging", err)
                print(traceback.format_exc())
                return False

        except Exception as err:
            print(err)
            print(traceback.format_exc())


    def select_role(self):
        try:
            self.driver.execute_script("return self.name || self.id")
            element = self.driver.find_element(By.XPATH, '//*[@id="seleccionRol:roles"]') 
            self.driver.execute_script("arguments[0].click();", element)
            element = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="seleccionRol:roles"]/option[6]')))
            element.click()
        except Exception as err:
            try:
                element = self.driver.find_element(By.XPATH, '//*[@id="j_id_jsp_1108549618_6:dtprogramacionesDeAmbiente:0:j_id_jsp_1108549618_10"]')
                self.driver.execute_script("arguments[0].click();", element)
            except Exception as err:
                print("ERROR AL SELECCIONAR ROL")
                print(err)
                print(traceback.format_exc())           


    def obtain_fichas_a_descargar(self, 
                               document_name="consulta.xlsx", 
                               other_path=None,  #other path se va
                               extraPath_toRoot="/entrada/"):
        
        #Cargar archivo de consulta
        #Verificar si ya existe un archivo de consulta, seleccionar si utilizar ese o reemplazarlo

        """
        Descarga de fichas...

        Attributes
            document_name: nombre archivo excel donde las fichas
            other_path (opcional): En caso de que el archivo no se encuentre en root 
            extraPath_toRoot (default): en caso de que se encuentre dentro de otras carpetas

        """
        
        try:
            consult_path  = f"{self.current_root_path}{extraPath_toRoot}{document_name}"
            if other_path:
                consult_path  = f"{other_path}{document_name}"
            if os.path.exists(consult_path):
                dfBase = pd.read_excel(consult_path)
                dfBase = pd.read_excel(consult_path)
                dfBase = pd.DataFrame(dfBase['IDENTIFICADOR_FICHA'])
                dfBase.drop_duplicates(subset=['IDENTIFICADOR_FICHA'], inplace=True)
                identificador_ficha_list = dfBase['IDENTIFICADOR_FICHA'].tolist()
                return identificador_ficha_list

        except Exception as err:
            print(err)
            print(traceback.format_exc())

    def keep_driver_open(self):
        input("Press any key to close the webdriver...")
        self.driver.quit()


    def depurate_from_existing_files(self, 
                               extraPath_toRoot="/procesar/",
                               list_fichas=None):
        

        """
        Calcular la cantidad de archivos faltantes y mostrarla

        Attributes
            other_path (opcional): En caso de que el archivo no se encuentre en root
            extraPath_toRoot (default): En caso de que se encuentre dentro de otras carpetas
            list_fichas []: Lista de las fichas obtenidas de entrada
        """
        try:
            removed_fichas = []
            procesar_path  = f"{self.current_root_path}{extraPath_toRoot}"
            file_list = os.listdir(procesar_path)
            existing_files = [file for file in file_list if file.startswith('Reporte de Aprendices Ficha')]
            # Calcular la cantidad de archivos faltantes y mostrarla
            for f in existing_files:
                existing_ficha = re.findall(r'\d+', f)
                if int(existing_ficha[0]) in list_fichas:
                    list_fichas.pop(list_fichas.index(int(existing_ficha[0])))
                    removed_fichas.append(int(existing_ficha[0]))
            #print(f"Hay {len(list_fichas)} archivos faltantes por descargar")
            print("Fichas a descargar: ", list_fichas)
            print("Fichas encontradas: ", removed_fichas)
            return list_fichas
        except Exception as err:
            print(err)
            print(traceback.format_exc())


    def download_files(self, list_fichas):
        try:
            if len(list_fichas) == 0:
                print("No hay fichas por descargar...")
                return False
            for f in list_fichas:
                self.driver.switch_to.default_content()
                element = self.wait.until(EC.element_to_be_clickable((By.XPATH, f'//body/div/div/nav/div[2]/div/div/form[2]/ul/li[6]/a')))
                #element.click()
                element = self.driver.find_element(By.XPATH, f'//body/div/div/nav/div[2]/div/div/form[2]/ul/li[6]/a')
                self.driver.execute_script("arguments[0].click();", element)
                element = self.driver.find_element(By.XPATH, f'//body/div/div/nav/div[2]/div/div/form[2]/ul/li[6]/ul/li/a')
                self.driver.execute_script("arguments[0].click();", element)
                element = self.driver.find_element(By.XPATH, f'//body/div/div/nav/div[2]/div/div/form[2]/ul/li[6]/ul/li/ul/li[3]/ul/li[3]/a')
                self.driver.execute_script("arguments[0].click();", element)
                self.driver.switch_to.frame(0)

                # Buscar la ficha de caracterización correspondiente y descargar el informe
                element = self.driver.find_element(By.XPATH, f'//body/div[2]/form/fieldset/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/a')
                time.sleep(1)
                element.click()
                self.driver.switch_to.frame(0)
                element = self.driver.find_element(By.XPATH, f'//div[2]/form/fieldset/div/table/tbody/tr/td[2]/input')
                element.send_keys(f)
                #print(len(nuevo_df_sin_duplicados), "por descargar")
                element = self.driver.find_element(By.XPATH, f'//div[2]/form/fieldset/div/div/input')
                self.driver.execute_script("arguments[0].click();", element)
                element = self.driver.find_element(By.XPATH, f'//body/div[2]/form/div/fieldset/table/tbody/tr/td/a')
                self.driver.execute_script("arguments[0].click();", element)
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(0)
                element = self.driver.find_element(By.XPATH, f'//body/div[2]/form/fieldset/div/div/input')
                self.driver.execute_script("arguments[0].click();", element)
                while os.path.exists(f"{self.current_root_path}/procesar/Reporte de Aprendices Ficha {f}.xls.crdownload") or \
                    any(file.startswith(".com.google.Chrome") for file in os.listdir(self.current_root_path)):
                    print("Waiting for file to finish downloading...")
                    time.sleep(1)
                if os.path.exists(f"{self.current_root_path}/procesar/Reporte de Aprendices Ficha {f}.xls"):
                    print(f"Descarga finalizada de {f}")
        except Exception as err:
            print(err)
            print(traceback.format_exc())


    def modified_files(self):
        try:
            print(os.listdir(self.path_process_folder))
            for filename in os.listdir(self.path_process_folder):
                if filename.endswith('.xls'):
                    print('Procesando...', filename)
                    name, extension = os.path.splitext(filename)
                    file_path_xlsx = f"{self.current_root_path}/procesar/{name}.xlsx"
                    df = pd.read_excel(f"{self.current_root_path}/procesar/{name}{extension}")
                    df.to_excel(file_path_xlsx, index=False)
                    os.remove(f"{self.current_root_path}/procesar/{filename}")
                    # Load the workbook
                    book = load_workbook(file_path_xlsx)
                    sheet = book.active

                    # Get the cell values
                    a2 = sheet.cell(row=2, column=1).value
                    c2 = sheet.cell(row=2, column=3).value
                    a3 = sheet.cell(row=3, column=1).value
                    c3 = sheet.cell(row=3, column=3).value

                    # Create a new workbook and sheet
                    new_book = Workbook()
                    new_sheet = new_book.active

                    # Write the values of the first 5 rows to the new sheet
                    for row in range(1, 6):
                        for col in range(1, sheet.max_column + 1):
                            new_sheet.cell(row=row, column=col+2).value = sheet.cell(row=row, column=col).value

                    # Write the header in the new column
                    new_sheet.cell(row=5, column=1).value = a2
                    new_sheet.cell(row=5, column=2).value = a3

                    # Write the values for the rest of the columns
                    for row in range(6, sheet.max_row + 1):
                        new_sheet.cell(row=row, column=1).value = c2
                        new_sheet.cell(row=row, column=2).value = c3
                        for col in range(1, sheet.max_column + 1):
                            new_sheet.cell(row=row, column=col+2).value = sheet.cell(row=row, column=col).value

                    # Save the modified Excel file
                    file_path_xlsx
                    new_file_path = os.path.join(self.path_process_folder, 'modified_' + name + '.xlsx')
                    new_book.save(new_file_path)

                    # Read the modified file and add the contents to the DataFrame
                    df_file = pd.read_excel(new_file_path, index_col=0)
                    self.df_consolidated = pd.concat([self.df_consolidated, df_file], ignore_index=True)

                    df_file_without_header = pd.read_excel(new_file_path)
                    df_file_without_header.to_excel(new_file_path, index=False)

                    # Print the DataFrame with the contents of the processed files
                    print(self.df_consolidated)
        except Exception as err:
            print(err)
            print(traceback.format_exc())

    def generate_consolidated_trainees(self):
        try: 
            for filename in os.listdir(self.path_process_folder):
                if filename.startswith('modified') and (filename.endswith('.xls') or filename.endswith('.xlsx')):
                    file_path = os.path.join(self.path_process_folder, filename)

                    # Leer el archivo de Excel
                    df = pd.read_excel(file_path, header=4)

                    self.df_consolidated = pd.concat([self.df_consolidated, df], ignore_index=True)
                    
            # Filtrar las filas que contengan 'formación' (en mayúsculas o minúsculas)
            df_formacion = self.df_consolidated[self.df_consolidated.apply(lambda row: 'formación' in str(row).lower(), axis=1)]

            # Guardar el DataFrame en un archivo de Excel
            output_file_path = self.path_process_folder + '/consolidado_final.xlsx'
            df_formacion.to_excel(output_file_path, index=False)

            print(f"Se ha guardado el archivo consolidado_final.xlsx en la siguiente ruta: {output_file_path}")

            file_list = [f for f in os.listdir(self.path_process_folder) if f.startswith('modified')]

            df_list = []
            for file_name in file_list:
                file_path = os.path.join(self.path_process_folder, file_name)
                print("Procesando archivo: ", file_path)
                df = pd.read_excel(file_path, skiprows=4)
                print(df.columns.tolist())
                df = df[['Ficha de Caracterización:', 'Estado:', 'Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Celular', 'Correo Electrónico', 'Estado']]
                df_list.append(df)

            df_concat = pd.concat(df_list, axis=0, ignore_index=True)

            output_path = f"{self.current_root_path}/entregables/Consolidado_Reporte_Aprendices.xlsx"
            df_concat.to_excel(output_path, index=False)
            df_concat = pd.read_excel(output_path)

            # Ver los primeros 5 registros del DataFrame
            print(df_concat.head())
            df_filtered = df_concat[df_concat['Estado'] == 'EN FORMACION']
        
            # Generar un archivo de Excel en la ruta deseada
            print("generaCION DE ARCHIVO")
            output_path = self.path_process_folder + "/Datos_formacion.xlsx"
            df_filtered.to_excel(output_path, index=False)
            # Ver los primeros 5 registros del DataFrame filtrado
            print(df_filtered.head())
            texto_df = pd.DataFrame(df_filtered.iloc[:, 0].str.split('-', expand=True)[0])
            
        except Exception as err:
            print(err)
            print(traceback.format_exc())

# BotConsulta = ConsultTrainees()
# BotConsulta.open_webdriver()
# BotConsulta.login_process()
# BotConsulta.select_role()
# list_fichas = BotConsulta.obtain_fichas_a_descargar()
# final_download_files = BotConsulta.depurate_from_existing_files(list_fichas=list_fichas)
# BotConsulta.download_files(final_download_files)
# BotConsulta.modified_files()
# BotConsulta.generate_consolidated_trainees()
# BotConsulta.keep_driver_open()

