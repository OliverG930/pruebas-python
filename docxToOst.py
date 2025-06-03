import os
import subprocess

# Directorios
origen = "C:\\Users\\Desktop\\informes\\DOCX"
destino = "C:\\Users\\Desktop\\informes\\OST"
os.makedirs(destino, exist_ok=True)

# Ruta al ejecutable de LibreOffice
soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

def convertir_docx_a_odt(origen, destino):
    for archivo in os.listdir(origen):
        if archivo.lower().endswith(".docx"):
            ruta_docx = os.path.join(origen, archivo)
            print(f"Convirtiendo: {archivo}")
            try:
                subprocess.run([
                    soffice_path,
                    "--headless",  # sin interfaz gráfica
                    "--convert-to", "odt",
                    "--outdir", destino,
                    ruta_docx
                ], check=True)
            except subprocess.CalledProcessError as e:
                print(f"❌ Error al convertir {archivo}: {e}")
    print("✔ Conversión finalizada.")

convertir_docx_a_odt(origen, destino)
