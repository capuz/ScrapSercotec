import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Función para capitalizar correctamente un título
def capitalizar_titulo(texto):
    excepciones = {"de", "en", "y", "a", "con", "sin", "por", "para", "del"}
    palabras = texto.lower().split()
    palabras_capitalizadas = [
        palabra.capitalize() if palabra not in excepciones else palabra
        for palabra in palabras
    ]
    return " ".join(palabras_capitalizadas)

# Función para extraer datos de un contenedor
def extraer_datos_contenedor(contenedor, tipo):
    items = contenedor.find_all("li")
    datos = []
    for item in items:
        fecha_tag = item.find("p", class_="inicio-btn")
        titulo_tag = item.find(["a", "label"]) or item.find("p")
        instrumento_tag = item.find("p", class_="instrumento")
        region_tag = item.find("p", class_="region")

        datos.append({
            "Título": capitalizar_titulo(titulo_tag.text.strip()) if titulo_tag else "Sin título",
            "Fecha": fecha_tag.text.strip() if fecha_tag else "Sin fecha",
            "Enlace": titulo_tag.get("href", "Sin enlace") if titulo_tag and titulo_tag.name == "a" else "Sin enlace",
            "Instrumento": instrumento_tag.text.strip() if instrumento_tag else "Sin instrumento",
            "Región": region_tag.text.strip() if region_tag else "Sin región",
            "Tipo": tipo
        })
    return datos

# Función principal para extraer datos y guardar en Excel
def guardar_convocatorias():
    url = "https://widgetsercotec.sercotec.cl/Calendario/TraerInformacion/?region=0&tipoInstrumento=0"
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.json().get("data", ""), "html.parser")

        datos = []
        for contenedor_id, tipo in [("contenedor1", "CONVOCATORIAS EN POSTULACIÓN"), 
                                    ("contenedor2", "PRÓXIMAS CONVOCATORIAS")]:
            contenedor = soup.find("div", id=contenedor_id)
            if contenedor:
                datos.extend(extraer_datos_contenedor(contenedor, tipo))

        df = pd.DataFrame(datos)
        mensaje = {
            "Título": "Si quieres verificar estos resultados, revísalos en www.sercotec.cl",
            "Fecha": "",
            "Enlace": "https://www.sercotec.cl",
            "Instrumento": "",
            "Región": "",
            "Tipo": ""
        }
        df_final = pd.concat([df, pd.DataFrame([mensaje])], ignore_index=True)

        # Guardar en Excel con enlaces clicables
        fecha_actual = datetime.now().strftime("%Y%m%d")
        nombre_por_defecto = f"Extraccion{fecha_actual}"
        nombre_archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar archivo Excel",
            initialfile=nombre_por_defecto,
        )
        if nombre_archivo:
            with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Convocatorias")
                worksheet = writer.sheets["Convocatorias"]
                for col in worksheet.columns:
                    max_length = max(len(str(cell.value)) for cell in col if cell.value)
                    adjusted_width = (max_length + 2) * 1.2
                    worksheet.column_dimensions[col[0].column_letter].width = adjusted_width

                # Convertir la columna "Enlace" a hipervínculos
                for idx, row in df_final.iterrows():
                    if row["Enlace"] != "Sin enlace" and row["Enlace"] != "www.sercotec.cl":
                        worksheet.cell(row=idx + 2, column=3).hyperlink = row["Enlace"]
                        worksheet.cell(row=idx + 2, column=3).style = "Hyperlink"

            messagebox.showinfo("Éxito", f"Archivo guardado en: {nombre_archivo}")
            try:
                os.startfile(nombre_archivo)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

    except requests.RequestException as e:
        messagebox.showerror("Error", f"Error al acceder a la URL: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")

# Interfaz gráfica
root = tk.Tk()
root.title("Extraer Convocatorias SERCOTEC")
root.geometry("350x190")
root.resizable(False, False)

btn_guardar = tk.Button(root, text="Guardar Convocatorias", command=guardar_convocatorias)
btn_guardar.pack(pady=50, padx=20, expand=True)

root.mainloop()