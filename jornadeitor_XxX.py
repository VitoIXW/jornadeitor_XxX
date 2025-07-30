import os
import sys
import yaml
from docx import Document
from docx.shared import Inches

def rellenar_registro(archivo_entrada, config_path="config.yaml"):
    # Verificar si el archivo existe
    if not os.path.exists(archivo_entrada):
        print(f"❌ El archivo {archivo_entrada} no existe.")
        return
    
    # Verificar si el config existe
    if not os.path.exists(config_path):
        print(f"❌ El archivo de configuración {config_path} no existe.")
        return
    
    # Cargar datos del YAML
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    
    nombre_apellidos = config.get("nombre_apellidos", "")
    nif = config.get("nif", "")
    naf = config.get("naf", "")
    hora_entrada = config.get("hora_entrada", "9:00")
    hora_salida = config.get("hora_salida", "17:00")
    firma_path = config.get("firma", "sign/sign.png")

    # Verificar si la firma existe
    if not os.path.exists(firma_path):
        print(f"❌ La firma {firma_path} no existe.")
        return

    # Cargar el documento
    doc = Document(archivo_entrada)

    # 1️⃣ Rellenar datos personales
    for p in doc.paragraphs:
        if "Nombre y apellidos:" in p.text:
            p.text = f"Nombre y apellidos: {nombre_apellidos}"
        if "NIF:" in p.text:
            p.text = f"NIF: {nif}"
        if "NAF:" in p.text:
            p.text = f"NAF: {naf}"

    # 2️⃣ Rellenar horas y firmas en la tabla
    for table in doc.tables:
        for row in table.rows:
            # Ignorar filas de encabezados o finales
            if any(palabra in row.cells[0].text.lower() for palabra in ["día", "recibido", "firma"]):
                continue
            
            # Verificar que tiene al menos 6 columnas (día + 4 horas + firma)
            if len(row.cells) >= 6:
                # Rellenar Entrada Mañana
                if row.cells[1].text.strip() == "":
                    row.cells[1].text = hora_entrada
                # Salida Mañana en blanco
                row.cells[2].text = ""
                # Entrada Tarde en blanco
                row.cells[3].text = ""
                # Rellenar Salida Tarde
                if row.cells[4].text.strip() == "":
                    row.cells[4].text = hora_salida
                
                # Vaciar celda de firma antes de insertar la imagen
                firma_cell = row.cells[5]
                firma_cell.text = ""
                run = firma_cell.paragraphs[0].add_run()
                run.add_picture(firma_path, width=Inches(1))

    # Crear carpeta output si no existe
    os.makedirs("output", exist_ok=True)
    
    # Nombre de salida
    nombre_salida = os.path.join("output", os.path.basename(archivo_entrada))
    doc.save(nombre_salida)
    print(f"✅ Archivo generado con datos personales, horas y firmas: {nombre_salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python rellenar_registro.py <archivo_entrada.docx> [config.yaml]")
    else:
        archivo = sys.argv[1]
        config = sys.argv[2] if len(sys.argv) > 2 else "config.yaml"
        rellenar_registro(archivo, config)
