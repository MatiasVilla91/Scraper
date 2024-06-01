import zipfile
import tempfile
import os
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from selenium.webdriver.chrome.options import Options

#SCRAPING

driver_path = r'C:\dchrome\chromedriver.exe'  #Ubicaciòn del driver
url = 'Link a la web a scrapear'     #Web solicitada
download_path = r'C:\Users\matia\Documents\prueba_tecnica'


# Configuración de opciones de Chrome
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})


# Inicializar el navegador
options = webdriver.ChromeOptions()

options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"  # Ruta Chrome

driver = webdriver.Chrome(options=chrome_options) #iniciar el navegador con la configuración de la ruta de descargas

# Acceder a la página
driver.get(url)

# Encontrar el enlace de "Documents & Downloads" y hacer clic
docs_downloads_link = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.LINK_TEXT, 'Documents & Downloads'))
)
docs_downloads_link.click()






#Encontrar el boton y elegir la opción Victims
nibrs_tables_link = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, 'dwnnibrs-download-select')) #id de boton es "dwnnibrs-download-select"
)
nibrs_tables_link.click()


menu_visible = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//nb-option[contains(text(), "Victims")]')) #La busqueda se hace con el texto "Victims" y no por el id porque el id en esta web es dinamico.
)

if menu_visible:
    # Obtener todos los elementos de opción en el menú
    opciones_menu = driver.find_elements(By.XPATH, '//nb-option[contains(text(), "Victims")]')

    # Intentar encontrar la opción que contiene "Victms" en el texto
    victims_option = next((opcion for opcion in opciones_menu if "Victims" in opcion.text), None)

    if victims_option:
        victims_option.click()
    else:
        print("No se encontró la opción 'Victims'.")
else:
    print("Error al encontrar victims")





# Encontrar segundo boton y elegir la ubicación "Florida"

nibrs_tables_link = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, 'dwnnibrsloc-select')) #id de boton es "dwnnibrs-download-select"dwnnibrsloc-select
)
nibrs_tables_link.click()

menu_visible2 = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//nb-option[contains(text(), "Florida")]'))
)

if menu_visible2:
    # Obtener todos los elementos de opción en el menú
    opciones_menu = driver.find_elements(By.XPATH, '//nb-option[contains(@id, "nb-option-")]')

    # Intentar encontrar la opción que contiene "Florida" en el texto
    florida_option = next((opcion for opcion in opciones_menu if "Florida" in opcion.text), None)

    if florida_option:
        florida_option.click()
    else:
        print("No se encontró la opción 'Florida'.")
else:
    print("El menú no está presente.")


# Descargar el archivo
boton_descarga = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "nibrs-download-button"))
)
boton_descarga.click()

# Esperar a que se complete la descarga
WebDriverWait(driver, 1200).until(
    lambda x: os.path.exists("C:/Users/matia/Documents/prueba_tecnica/Victims.zip")
)

# Cerrar el navegador después de la descarga
driver.quit()


#Abrir archivo zip

# Ruta al archivo zip
ruta_zip = 'C:/Users/matia/Documents/prueba_tecnica/victims.zip'

#Elegir archivo: Victims_Age_by_Offense_Category_2022.xlsx
nombre_archivo_excel = 'Victims_Age_by_Offense_Category_2022.xlsx'

# Crear un directorio temporal
directorio_temporal = tempfile.mkdtemp()

try:
    # Leer el archivo ZIP
    with zipfile.ZipFile(ruta_zip) as zip_ref:
        # Leer el contenido del archivo Excel directamente desde el ZIP
        with zip_ref.open(nombre_archivo_excel) as excel_file:
            # Leer el archivo Excel en un DataFrame
            datos = pd.read_excel(excel_file, engine='openpyxl', skiprows=13, skipfooter=1)

    # Filtrar datos
    # elegir categoría Crimes Against Property y generar csv sin totales, footer, ni index
    # Seleccionar las filas y columnas
    datos_filtrados = datos.iloc[0:11, 2:16]

    # Restablecer el índice del DataFrame
    datos_filtrados.reset_index(drop=True, inplace=True)

    # Guardar el DataFrame en un nuevo archivo CSV
    ruta_nuevo_csv = 'C:/Users/matia/Documents/prueba_tecnica/prueba_finalizada.csv'
    datos_filtrados.to_csv(ruta_nuevo_csv, index=False)

finally:
    # Eliminar el directorio temporal
    os.rmdir(directorio_temporal)

print('El código se ejecutó correctamente. Aqui estan los datos:\n')
print(datos_filtrados)
os.system(f'start excel "{ruta_nuevo_csv}"')  # abrir archivo Excel
