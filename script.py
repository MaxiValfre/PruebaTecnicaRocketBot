# Librerias requeridas
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Librerias adicionales
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import time

# Configuración de Selenium
driver = webdriver.Chrome()    #Si no tenes el driver en el path, ingresarlo como ruta al .exe
wait = WebDriverWait(driver, 20)
driver.maximize_window()

# Abre el archivo Excel
wb = xw.Book('Base Seguimiento Observ Auditoría al_30042021.xlsx')
sheet = wb.sheets['Hoja1'] 

# Función para enviar correo electrónico
def enviar_correo(destinatario, asunto, cuerpo, cuerpohtml):
    remitente = ''  #Ingresar usuario
    password = ''   #Ingresar Contraseña
    
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    
    msg.attach(MIMEText(cuerpohtml, 'html'))
    msg.attach(MIMEText(cuerpo, 'text'))
    
    try:
        server = smtplib.SMTP('', 25) #Ingresar servidor 
        server.starttls()
        server.login(remitente, password)
        texto = msg.as_string()
        server.sendmail(remitente, destinatario, texto)
        server.quit()
        print(f"Correo enviado a {destinatario} ubicado en la fila {row}")
    except Exception as e:
        print(f"Error al enviar correo: {e}")

# Función para seleccionar un valor del menú desplegable
def seleccionar_valor_del_menu(menu, valor):
    valor = valor.strip().lower()
    for option in menu.options:
        if option.text.strip().lower() == valor or option.get_attribute('value').strip().lower() == valor:
            menu.select_by_visible_text(option.text)
            return True
    
    print(f"Valores disponibles en el menú de proceso: {[opt.text for opt in menu.options]}")
    return False

# Función para procesar cada fila
def procesar_fila(row):
    try:
        estado = sheet.range(f'J{row}').value
        if estado == 'Regularizado':
            # Subir información al formulario usando Selenium
            driver.get('https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG')

            # Esperar a que el formulario se cargue completamente
            wait.until(EC.presence_of_element_located((By.ID, 'form')))

            # Rellenar el formulario
            try:
                # Seleccionar un valor del menú desplegable de proceso
                menu_proceso = Select(wait.until(EC.presence_of_element_located((By.NAME, 'process'))))
                if not seleccionar_valor_del_menu(menu_proceso, sheet.range(f'A{row}').value):
                    print(f"No se encontró el valor {sheet.range(f'A{row}').value} en el menú de proceso para la fila {row}")
                    return
                
                # Rellenar el campo de texto para tipo de riesgo
                campo_tipo_riesgo = wait.until(EC.presence_of_element_located((By.NAME, 'tipo_riesgo')))
                campo_tipo_riesgo.send_keys(sheet.range(f'C{row}').value)

                # Seleccionar un valor del menú desplegable para la severidad
                menu_severidad = Select(wait.until(EC.presence_of_element_located((By.NAME, 'severidad'))))
                if not seleccionar_valor_del_menu(menu_severidad, sheet.range(f'D{row}').value):
                    print(f"No se encontró el valor {sheet.range(f'D{row}').value} en el menú de severidad para la fila {row}")
                    return

                # Rellenar el campo de texto para el responsable
                campo_responsable = wait.until(EC.presence_of_element_located((By.NAME, 'res')))
                campo_responsable.send_keys(sheet.range(f'G{row}').value)

                # Rellenar el campo de fecha
                campo_fecha = wait.until(EC.presence_of_element_located((By.NAME, 'date')))
                fecha_compromiso = sheet.range(f'F{row}').value
                if isinstance(fecha_compromiso, datetime):
                    fecha_compromiso = fecha_compromiso.strftime('%d/%m/%Y')
                campo_fecha.send_keys(fecha_compromiso)

                # Rellenar el textarea para la observación
                campo_observacion = wait.until(EC.presence_of_element_located((By.NAME, 'obs')))
                campo_observacion.send_keys(sheet.range(f'B{row}').value)

                # Enviar el formulario
                boton_enviar = wait.until(EC.presence_of_element_located((By.ID, 'submit')))
                boton_enviar.click()
                print(f"Formulario enviado para la fila {row}.")
            
            except Exception as e:
                print(f"Error al procesar la fila {row}: {e}")


        # Si el estado es atrasado envia un correo al responsable
        elif estado == 'Atrasado':
            
            proceso = sheet.range(f'A{row}').value
            observacion = sheet.range(f'B{row}').value
            fecha_compromiso = sheet.range(f'F{row}').value
            responsable = sheet.range(f'G{row}').value

            asunto = f"Comunicacion del estado de proceso"

            cuerpohtml = f"""
            Proceso:{proceso}<br>
            Estado: {estado}<br>
            Observación: {observacion}<br>
            Fecha de Compromiso: {fecha_compromiso}<br>
            """
            
            cuerpo = f"""
            Proceso:{proceso}\r\n
            Estado: {estado}\r\n
            Observación: {observacion}\r\n
            Fecha de Compromiso: {fecha_compromiso}\r\n
            """
            enviar_correo(responsable, asunto, cuerpo, cuerpohtml)
            

    except Exception as e:
        print(f"Error al procesar la fila {row}: {e}")


# Itera sobre las filas del archivo Excel
row = 2
while True:
    try:
        if not sheet.range(f'J{row}').value:
            break
        procesar_fila(row)
        time.sleep(1) 
        row += 1
    except Exception as e:
        print(f"Error al procesar la fila {row}: {e}")

# Cierra el buscador y el excel
def cerrar_recursos(driver, wb):
    try:
        driver.quit()
        print("Navegador Chrome cerrado correctamente.")
    except Exception as e:
        print(f"Error al cerrar el navegador Chrome: {e}")

    try:
        wb.close()
        print("Archivo Excel cerrado correctamente.")
    except Exception as e:
        print(f"Error al cerrar el archivo Excel: {e}")
        time.sleep(5)
        try:
            wb.close()
        except Exception as e:
            print(f"Error al cerrar el archivo Excel en segundo intento: {e}")

cerrar_recursos(driver, wb)

