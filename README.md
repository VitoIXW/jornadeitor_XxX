## 📝 Jornadeitor\_XxX

Este script permite **rellenar automáticamente un registro de jornada laboral en formato `.docx`**, añadiendo los datos de **empleado**, **fechas**, **horas de entrada/salida** y **firmas** en cada día trabajado.

---

### **📦 Requisitos**

* Python **3.8+**
* Librerías:

  ```bash
  pip install python-docx pyyaml
  ```
* Plantilla de registro en formato **.docx**
* Archivo de configuración **`config.yaml`** (ver más abajo).

---

### **▶️ Uso**

```bash
python rellenar_registro.py <archivo_entrada.docx> [config.yaml]
```

**Ejemplo:**

```bash
python rellenar_registro.py "input/PLANTILLA REGISTRO FEBRERO.docx"
```

* Si no se especifica el segundo parámetro, por defecto usará `config.yaml`.
* El archivo procesado se guardará en la carpeta `output` con el **mismo nombre** que el original.

---

### **🔹 Ayuda**

Puedes mostrar la ayuda con:

```bash
python rellenar_registro.py -h
```

Esto mostrará:

```
Uso:
    python rellenar_registro.py <archivo_entrada.docx> [config.yaml]

Opciones:
    -h, --help      Muestra esta ayuda

Descripción:
    Este script rellena automáticamente un registro de jornada en formato .docx,
    actualizando el nombre, NIF, NAF, fecha y las horas de entrada/salida.
    También añade la firma diaria en cada fila de la tabla principal.
```

---

### **⚙️ Archivo de configuración (`config.yaml`)**

Puedes usar el archivo de ejemplo incluido: [`config.example.yaml`](config.example.yaml)
Solo cópialo y modifícalo según tus datos:

```bash
cp config.example.yaml config.yaml
```

Ejemplo de contenido:

```yaml
nombre_apellidos: "Juan Pérez García"
nif: "12345678A"
naf: "1122334455"
hora_entrada_mañana: "09:00"
hora_salida_mañana: ""        # dejar vacío si no se quiere rellenar
hora_entrada_tarde: ""        # dejar vacío si no se quiere rellenar
hora_salida_tarde: "17:00"
firma: "sign/sign.png"        # ruta al archivo de la firma
```

* Si dejas un campo vacío (por ejemplo `hora_entrada_tarde`), **no se rellenará** en la tabla.
* La firma debe estar en formato **.png**.

---

### **📄 Salida**

* Se genera un archivo **.docx** con los datos completados en la carpeta `output/`.
* El nombre será el mismo que el archivo de entrada.
