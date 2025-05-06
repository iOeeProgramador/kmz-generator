import pandas as pd
import re
import zipfile
from xml.dom.minidom import Document
from datetime import datetime
import streamlit as st

def extraer_lat_lon(texto):
    match = re.search(r'Latitude\s*:\s*(-?\d+\.\d+).*Longitude\s*:\s*(-?\d+\.\d+)', texto, re.DOTALL)
    return (float(match.group(1)), float(match.group(2))) if match else (None, None)

def get_estilo(servicio, tipo):
    papeleras = ["Classic", "Classic Antiguo", "Prima Linea"]
    contenedores = ["240 L", "360 L", "770 L"]
    if servicio == "Mantencion":
        return "azul"
    elif servicio == "Hidrolavado":
        return "verde"
    elif tipo in papeleras:
        return "amarillo"
    elif tipo in contenedores:
        return "rojo"
    else:
        return "gris"

def crear_estilo_kml(doc, id_color, color_hex):
    style = doc.createElement("Style")
    style.setAttribute("id", id_color)
    icon_style = doc.createElement("IconStyle")
    color = doc.createElement("color")
    color.appendChild(doc.createTextNode(color_hex))
    icon_style.appendChild(color)
    scale = doc.createElement("scale")
    scale.appendChild(doc.createTextNode("1.2"))
    icon_style.appendChild(scale)
    icon = doc.createElement("Icon")
    href = doc.createElement("href")
    href.appendChild(doc.createTextNode("http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"))
    icon.appendChild(href)
    icon_style.appendChild(icon)
    style.appendChild(icon_style)
    return style

def generar_kmz(df, servicio, nombre_archivo):
    doc = Document()
    kml = doc.createElement("kml")
    kml.setAttribute("xmlns", "http://www.opengis.net/kml/2.2")
    doc.appendChild(kml)
    document = doc.createElement("Document")
    kml.appendChild(document)

    estilos = {
        "rojo": "ff0000ff",
        "amarillo": "ff00ffff",
        "verde": "ff00ff00",
        "azul": "ffffaa00",
        "gris": "ffaaaaaa"
    }

    for key, hex_color in estilos.items():
        style_element = crear_estilo_kml(doc, key, hex_color)
        document.appendChild(style_element)

    for _, row in df.iterrows():
        tipo = row.get('Tipo contenedor', '')
        estilo = get_estilo(servicio, tipo)
        placemark = doc.createElement("Placemark")
        description = doc.createElement("description")
        nombre_texto = str(row['Nombre']) if pd.notna(row['Nombre']) else "Sin nombre"
        description.appendChild(doc.createTextNode(nombre_texto))
        placemark.appendChild(description)
        style_url = doc.createElement("styleUrl")
        style_url.appendChild(doc.createTextNode(f"#{estilo}"))
        placemark.appendChild(style_url)
        point = doc.createElement("Point")
        coordinates = doc.createElement("coordinates")
        coordinates.appendChild(doc.createTextNode(f"{row['Longitud']},{row['Latitud']},0"))
        point.appendChild(coordinates)
        placemark.appendChild(point)
        document.appendChild(placemark)

    with open("doc.kml", "w", encoding="utf-8") as f:
        f.write(doc.toprettyxml(indent="  "))
    with zipfile.ZipFile(nombre_archivo, "w", zipfile.ZIP_DEFLATED) as kmz:
        kmz.write("doc.kml", arcname="doc.kml")

# Streamlit UI
st.title("Generador de KMZ para Servicios Municipales")
archivo = st.file_uploader("Sube el archivo Excel (.xlsx)", type=["xlsx"])
servicio = st.selectbox("Selecciona el servicio", ["Instalacion", "Hidrolavado", "Mantencion"])
from datetime import datetime

fecha_ini = st.date_input("Fecha inicio")
fecha_fin = st.date_input("Fecha fin")

# Convertir las fechas a datetime para evitar errores de comparación
fecha_ini = datetime.combine(fecha_ini, datetime.min.time())
fecha_fin = datetime.combine(fecha_fin, datetime.max.time())


if archivo and fecha_ini and fecha_fin:
    try:
        df = pd.read_excel(archivo, sheet_name=servicio)

        if servicio == "Instalacion":
            df['Fecha de respuesta'] = pd.to_datetime(df['Fecha de respuesta'], errors='coerce')
            df = df[(df['Fecha de respuesta'] >= fecha_ini) & (df['Fecha de respuesta'] <= fecha_fin)]
            df['Nombre'] = df['Fecha de respuesta'].dt.strftime('%Y-%m-%d') + ' ' + df['Tipo contenedor']
            nombre_archivo = "Instalacion.kmz"

        elif servicio == "Hidrolavado":
            df['Fecha de respuesta'] = pd.to_datetime(df['Fecha de respuesta'], errors='coerce')
            df = df[(df['Fecha de respuesta'] >= fecha_ini) & (df['Fecha de respuesta'] <= fecha_fin)]
            df['Nombre'] = df['Fecha de respuesta'].dt.strftime('%Y-%m-%d') + ' ' + df['Id de Servicio'].astype(str)
            nombre_archivo = "Hidrolavado.kmz"

        elif servicio == "Mantencion":
            df['Fecha y hora'] = pd.to_datetime(df['Fecha y hora'], errors='coerce')
            df = df[(df['Fecha y hora'] >= fecha_ini) & (df['Fecha y hora'] <= fecha_fin)]
            df['Nombre'] = df['Fecha y hora'].dt.strftime('%Y-%m-%d') + ' ' + df['Capacidad del contenedor'].astype(str)
            nombre_archivo = "Mantencion.kmz"

        df[['Latitud', 'Longitud']] = df['Ubicacion'].apply(lambda x: pd.Series(extraer_lat_lon(str(x))))
        generar_kmz(df, servicio, nombre_archivo)

        with open(nombre_archivo, "rb") as f:
            st.download_button("📥 Descargar KMZ", f, file_name=nombre_archivo)
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
