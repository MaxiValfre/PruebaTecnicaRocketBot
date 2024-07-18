# Proyecto de Automatización con Selenium y Excel

## Descripción

Este proyecto utiliza Python para automatizar la interacción con un formulario web y el envío de correos electrónicos basándose en los datos de un archivo Excel. El script realiza las siguientes tareas:

1. **Lee un archivo Excel** para obtener datos relevantes.
2. **Interacciona con un formulario web** utilizando Selenium.
3. **Envía correos electrónicos** a los responsables de los procesos que están marcados como "Atrasado".


## Requisitos

- Python 3.x
- `xlwings` para manipulación de archivos Excel.
- `selenium` para interacción con el navegador.
- `smtplib` para el envío de correos electrónicos.
- `datetime` para el manejo de fechas y horas.

Asegúrate de tener el [ChromeDriver](https://sites.google.com/chromium.org/driver/) en tu PATH.

## Instalación

1. **Clona el repositorio** (si aplica):
   ```bash
   git clone <URL_DEL_REPOSITORIO>
2. **Instala las dependencias**
   ```bash
   pip install xlwings selenium
   

## Uso

Asegúrate de que el archivo Excel esté en el mismo directorio que el script o proporciona la ruta correcta.
El archivo debe tener una hoja llamada Hoja1 con las columnas necesarias en las posiciones especificadas en el script.
Modifica las credenciales para el envío de correos en el script:

- Actualiza las variables remitente y password en la función enviar_correo.
- Ejecuta el script
   ```bash
      python script.py

## Funcionalidad del Script
- Interacción con el Formulario Web
- URL del formulario: https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG
- Campos del formulario:
    - process: Menú desplegable para seleccionar un proceso.
    - tipo_riesgo: Campo de texto para el tipo de riesgo.
    - severidad: Menú desplegable para la severidad.
    - res: Campo de texto para el responsable.
    - date: Campo de fecha para la fecha de compromiso.
    - obs: Área de texto para observaciones.
    - submit: Botón para enviar el formulario.

- Envío de Correos Electrónicos:
Los correos electrónicos se envían a los responsables de los procesos que están marcados como "Atrasado".
La función enviar_correo utiliza SMTP para enviar los correos. Asegúrate de que el servidor SMTP y las credenciales estén configurados correctamente.

## Consideraciones
- Maximización del navegador: La ventana del navegador se maximiza al iniciar para facilitar la visualización durante la automatización.
- Intervalo entre filas: El script espera 1 segundo entre el procesamiento de cada fila para reducir la carga en el servidor web y mejorar la estabilidad.
