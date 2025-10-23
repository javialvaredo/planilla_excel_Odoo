# 🧾 Modificador de Documentos de Odoo (Excel)

## 📖 Descripción
Este proyecto es una herramienta en **Python** que permite **abrir, visualizar y modificar archivos Excel exportados desde Odoo**.  
Usa una interfaz gráfica desarrollada con **Tkinter** y aplica transformaciones automáticas sobre columnas específicas, según reglas predefinidas (por ejemplo, tipo de moneda o tipo de transacción).

---

## ⚙️ Funcionalidades principales

- 📂 **Abrir archivo Excel** desde una ventana de diálogo.  
- ✏️ **Modificar valores de la columna F** según el tipo de moneda:
  - 💵 `USD` → `=valor*TC`
  - 💶 `EUR` → `=valor*EUR`
  - 🇦🇷 `ARS` → deja el valor sin cambios  
- 💼 **Procesar registros con "Despachos" o "Tr_exterior"** en la columna B:
  - Si contiene `"Despachos"` → escribe en columna **E** `=valor*0.79*TC`
  - Si contiene `"Tr_exterior"` → escribe en columna **E** `=valor*TC`
- 🎨 **Aplica formato numérico automático** con separador de miles y sin decimales.  
- 💾 **Guarda automáticamente** el archivo modificado.  
- 📊 **Permite abrir el Excel resultante** desde la interfaz.

---

## 🧩 Estructura del código

El proyecto está organizado con **Programación Orientada a Objetos (POO)**:

| Clase | Descripción |
|-------|--------------|
| `ExcelManager` | Gestiona la lectura, escritura y modificación del archivo Excel usando `openpyxl`. |
| `AplicacionTkinter` | Crea la interfaz gráfica con Tkinter y los botones para ejecutar las distintas acciones. |

---

## 🚀 Requisitos

Instalar las dependencias necesarias:

```bash
pip install openpyxl

python main.py
| Columna B   | Columna F | Columna G | Resultado                     |
| ----------- | --------- | --------- | ----------------------------- |
| Factura     | 1000      | USD       | `=1000*TC`                    |
| Despachos   | 500       | USD       | (en columna E) `=500*0.79*TC` |
| Tr_exterior | 750       | USD       | (en columna E) `=750*TC`      |
| Local       | 200       | ARS       | `200`                         |


💡 Tecnologías utilizadas

🐍 Python 3

📊 openpyxl

🪟 Tkinter

✍️ Autor

Desarrollado por Javier Alvaredo
💼 Automatización de procesos para planillas Excel de Odoo

🧠 Licencia

Este proyecto se distribuye bajo licencia MIT.
Podés usarlo, modificarlo y compartirlo libremente, citando la fuente.