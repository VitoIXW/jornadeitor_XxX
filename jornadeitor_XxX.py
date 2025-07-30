import os
import sys
from docx import Document

def rellenar_registro(archivo_entrada, hora_entrada="9:00", hora_salida="17:00"):
    # Verificar si el archivo existe
    if not os.path.exists(archivo_entrada):
        print(f"❌ El archivo {archivo_entrada} no existe.")
        return
    
    # Cargar el documento
    doc = Document(archivo_entrada)

    # Rellenar las celdas vacías con las horas indicadas
    for table in doc.tables:
        for row in table.rows:
            # Ignorar filas de encabezados o finales
            if any(palabra in row.cells[0].text.lower() for palabra in ["día", "recibido", "firma"]):
                continue

            # Las columnas están en orden: Día, Entrada Mañana, Salida Mañana, Entrada Tarde, Salida Tarde
            if len(row.cells) >= 5:
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

    # Crear carpeta output si no existe
    os.makedirs("output", exist_ok=True)
    
    # Nombre de salida
    nombre_salida = os.path.join("output", os.path.basename(archivo_entrada))
    doc.save(nombre_salida)
    print(f"✅ Archivo generado: {nombre_salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python rellenar_registro.py <archivo_entrada.docx> [hora_entrada] [hora_salida]")
    else:
        archivo = sys.argv[1]
        hora_entrada = sys.argv[2] if len(sys.argv) > 2 else "9:00"
        hora_salida = sys.argv[3] if len(sys.argv) > 3 else "17:00"
        rellenar_registro(archivo, hora_entrada, hora_salida)
