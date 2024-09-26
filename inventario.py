import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import ujson

def exportar_inventario(archivo_excel, almacenes, archivo_config='cantidades_minimas.json', columna_codigo='Codigo', columna_producto='Producto', columna_cantidad='Cantidad'):
    """
    Exporta a un archivo Excel un resumen del inventario, agrupado por producto y almacen, comparando con cantidades mínimas.

    Args:
        archivo_excel (str): Ruta del archivo Excel de origen.
        almacenes (list): Lista de almacenes a filtrar.
        archivo_config (str, optional): Ruta del archivo de configuración con las cantidades mínimas. Defaults to 'cantidades_minimas.json'.
        columna_codigo (str, optional): Nombre de la columna de código de producto. Defaults to 'Codigo'.
        columna_producto (str, optional): Nombre de la columna de nombre de producto. Defaults to 'Producto'.
        columna_cantidad (str, optional): Nombre de la columna de cantidad. Defaults to 'Cantidad'.
    """

    try:
        df = pd.read_excel(archivo_excel)
    except FileNotFoundError:
        print(f"El archivo {archivo_excel} no se encontró.")
        return

    # Verificar existencia de columnas
    for columna in [columna_codigo, columna_producto, columna_cantidad]:
        if columna not in df.columns:
            print(f"La columna '{columna}' no existe en el DataFrame.")
            return

    # Cargar las cantidades mínimas desde el archivo JSON
    with open(archivo_config, 'r', encoding='utf-8') as f:
    
        cantidades_minimas = ujson.load(f)

    # Filtrar y agrupar
    df_filtrado = df[df['Almacen'].isin(almacenes)]
    cantidad_por_producto = df_filtrado.groupby([columna_codigo, columna_producto])[columna_cantidad].sum().reset_index()

    # Agregar columna para la cantidad mínima
    cantidad_por_producto['Cantidad_minima'] = cantidad_por_producto[columna_producto].map(cantidades_minimas)
    # Agregar columna para indicar si se cumple la cantidad mínima
    cantidad_por_producto['Cumple_minimo'] = cantidad_por_producto['Cantidad'] >= cantidad_por_producto['Cantidad_minima']

    # Separar los productos
    productos_bajo_cero = cantidad_por_producto[cantidad_por_producto['Cantidad'] <= 0]
    productos_bajo_minimo = cantidad_por_producto[(cantidad_por_producto['Cantidad'] > 0) & (cantidad_por_producto['Cantidad'] < cantidad_por_producto['Cantidad_minima'])]
    productos_cumplen_minimo = cantidad_por_producto[cantidad_por_producto['Cumple_minimo']]

    # Exportar a Excel
    with pd.ExcelWriter('resultado_inventario.xlsx') as writer:
        productos_bajo_cero.to_excel(writer, sheet_name='Productos agotados o negativos', index=False)
        productos_bajo_minimo.to_excel(writer, sheet_name='Productos por debajo del mínimo', index=False)
        productos_cumplen_minimo.to_excel(writer, sheet_name='Productos que cumplen el mínimo', index=False)

    print("Los resultados se han exportado a resultado_inventario.xlsx")



    # Enviar correo electrónico
def enviar_correo():
    """
    Envía un correo electrónico con el informe de inventario generado.
    """
    # Validar existencia del archivo
    if not os.path.isfile('resultado_inventario.xlsx'):
        print("No se ha encontrado el archivo resultado_inventario.xlsx.")
        return
    
    
    
    # Configuración del SMTP y datos de tu cuenta de correo (reemplaza con tus datos)
    smtp_server ='smtp.gmail.com'
    port = 587
    sender_email = 'yelkindavid1997@gmail.com'
    password = ''
    receiver_email = 'barbara@lab-ol.com'

    # Crear mensaje
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = 'Informe de inventario'
    body = "Adjunto encontrarás el informe detallado del inventario."
    message.attach(MIMEText(body, 'plain'))
    
    # Adjuntar archivo
    filename ='resultado_inventario.xlsx'
    attachment = open(filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename={filename}")
    message.attach(part)
    

    # Enviar correo
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    # Imprimir mensaje de éxito
    print("El correo se ha enviado correctamente.")

    

    
    


# Ejemplo de uso
exportar_inventario('inventario.xls', ['P502', 'MZ02'])

enviar_correo()
