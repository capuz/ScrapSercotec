import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Función para extraer datos y guardar en Excel
def guardar_convocatorias():
    url = "https://widgetsercotec.sercotec.cl/Calendario/TraerInformacion/?region=0&tipoInstrumento=0"
    response = requests.get(url)
    if response.status_code == 200:
        json_data = response.json()
        html_content = json_data.get("data", "")
        soup = BeautifulSoup(html_content, "html.parser")
        datos = []  # Lista para almacenar todos los datos

        def extraer_convocatorias(contenedor, titulo_esperado):
            titulo_contenedor = contenedor.find("div", id="nombre").find("p").text.strip()
            if titulo_esperado in titulo_contenedor:
                print(f"\n--- {titulo_contenedor} ---")
                items = contenedor.find_all("li")
                for item in items:
                    fecha = item.find("p", class_="inicio-btn")
                    fecha_texto = fecha.text.strip() if fecha else "Sin fecha"

                    # Extraer el título (puede estar en <a>, <label>, u otro tag)
                    titulo = item.find("a") or item.find("label") or item.find("p")
                    titulo_texto = titulo.text.strip() if titulo else "Sin título"
                    enlace = titulo["href"] if titulo and titulo.name == "a" else "Sin enlace"

                    instrumento = item.find("p", class_="instrumento")
                    instrumento_texto = instrumento.text.strip() if instrumento else "Sin instrumento"
                    region = item.find("p", class_="region")
                    region_texto = region.text.strip() if region else "Sin región"

                    # Agregar los datos a la lista
                    datos.append({
                        "Título": titulo_texto,
                        "Fecha": fecha_texto,
                        "Enlace": enlace,
                        "Instrumento": instrumento_texto,
                        "Región": region_texto,
                        "Tipo": titulo_esperado
                    })
                    print(f"Fecha: {fecha_texto}")
                    print(f"Título: {titulo_texto}")
                    print(f"Enlace: {enlace}")
                    print(f"Instrumento: {instrumento_texto}")
                    print(f"Región: {region_texto}")
                    print("-" * 50)
            else:
                print(f"El contenedor no tiene el título esperado: {titulo_esperado}")

        contenedor1 = soup.find("div", id="contenedor1")
        if contenedor1:
            extraer_convocatorias(contenedor1, "CONVOCATORIAS EN POSTULACIÓN")
        else:
            print("No se encontró el contenedor1.")

        contenedor2 = soup.find("div", id="contenedor2")
        if contenedor2:
            extraer_convocatorias(contenedor2, "PRÓXIMAS CONVOCATORIAS")
        else:
            print("No se encontró el contenedor2.")

        # Crear un DataFrame con los datos
        df = pd.DataFrame(datos)

        # Crear una fila adicional con el mensaje
        mensaje = {
            "Título": "Si quieres verificar estos resultados, revísalos en www.sercotec.cl",
            "Fecha": "",
            "Enlace": "www.sercotec.cl",
            "Instrumento": "",
            "Región": "",
            "Tipo": ""
        }
        df_mensaje = pd.DataFrame([mensaje])

        # Concatenar el mensaje al DataFrame principal
        df_final = pd.concat([df, df_mensaje], ignore_index=True)

        # Obtener la fecha actual en el formato YYYYMMDD
        fecha_actual = datetime.now().strftime("%Y%m%d")
        nombre_por_defecto = f"Extraccion{fecha_actual}"

        # Abrir un cuadro de diálogo para seleccionar la ubicación del archivo
        nombre_archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar archivo Excel",
            initialfile=nombre_por_defecto,
        )
        if nombre_archivo:
            # Guardar el DataFrame en un archivo Excel
            with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Convocatorias")
                # Acceder al libro y la hoja de trabajo
                workbook = writer.book
                worksheet = writer.sheets["Convocatorias"]
                # Ajustar el ancho de las columnas
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    worksheet.column_dimensions[column].width = adjusted_width
            # Mostrar mensaje de éxito
            messagebox.showinfo("Éxito", f"Archivo guardado en: {nombre_archivo}")
            # Abrir el archivo Excel automáticamente
            try:
                os.startfile(nombre_archivo)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
    else:
        messagebox.showerror("Error", f"Error al acceder a la URL: {response.status_code}")

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Extraer Convocatorias SERCOTEC")
root.geometry("350x190")
root.resizable(False, False)
btn_guardar = tk.Button(root, text="Guardar Convocatorias", command=guardar_convocatorias)
btn_guardar.pack(pady=50, padx=20, expand=True)
root.mainloop()
