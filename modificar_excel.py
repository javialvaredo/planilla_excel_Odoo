import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
from openpyxl import load_workbook

class ExcelManager:
    """Clase encargada de manejar y modificar archivos Excel."""
    def __init__(self):
        self.workbook = None
        self.file_path = None
        self.modificado_path = None

    def abrir_archivo(self, file_path):
        """Carga un archivo Excel existente."""
        try:
            self.workbook = load_workbook(file_path)
            self.file_path = file_path
            # definir nombre del archivo modificado
            base, ext = os.path.splitext(file_path)
            self.modificado_path = f"{base}_modificado{ext}"
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
            return False

    def limpiar_columna(self, hoja, columna="E"):
        """Deja la columna indicada sin valores."""
        ws = self.workbook[hoja]
        for fila in range(1, ws.max_row + 1):
            ws[f"{columna}{fila}"].value = None

    def modificar_columna_condicional(self, hoja):
        """
        Modifica columna F según condiciones:
        - Si columna B contiene 'Despachos' → =valor*0.79*TC
        - Si columna G = 'USD' → =valor*TC
        - Si columna G = 'EUR' → =valor*EUR
        - Si columna G = 'ARS' → deja valor sin cambios
        Aplica formato con separador de miles y sin decimales.
        """
        ws = self.workbook[hoja]
        for fila in range(1, ws.max_row + 1):
            celda_b = ws[f"B{fila}"]
            celda_f = ws[f"F{fila}"]
            celda_g = ws[f"G{fila}"]

            if celda_f.value is None or celda_g.value is None:
                continue

            moneda = str(celda_g.value).strip().upper()
            descripcion = str(celda_b.value).strip().upper() if celda_b.value else ""

            try:
                valor = float(celda_f.value)
            except (ValueError, TypeError):
                continue  # si no es numérico, lo salta

            # Condicional especial por columna B
            if "DESPACHOS" in descripcion:
                celda_f.value = f"={valor}*0.79*TC"

            elif moneda == "USD":
                celda_f.value = f"={valor}*TC"

            elif moneda == "EUR":
                celda_f.value = f"={valor}*EUR"

            # Si es ARS o cualquier otro → sin cambios

            # Formato con separador de miles, sin decimales
        celda_f.number_format = '#,##0'
        

    
    def guardar_modificado(self):
        """Guarda el archivo modificado con sufijo _modificado."""
        self.workbook.save(self.modificado_path)
        return self.modificado_path

    def limpiar_y_modificar(self, hoja="Hoja1"):
        """Ejecuta todo el proceso: limpiar E y modificar F según G."""
        self.limpiar_columna(hoja, "E")
        self.modificar_columna_condicional(hoja)
        return self.guardar_modificado()


class ExcelApp(tk.Tk):
    """Clase principal que maneja la interfaz Tkinter."""
    def __init__(self):
        super().__init__()
        self.title("Editor de Excel Condicional")
        self.geometry("500x250")

        self.excel_manager = ExcelManager()
        self._crear_widgets()

    def _crear_widgets(self):
        """Crea los elementos de la interfaz."""
        self.label_archivo = tk.Label(self, text="Ningún archivo seleccionado")
        self.label_archivo.pack(pady=10)

        # Botón abrir archivo
        btn_abrir = tk.Button(self, text="Abrir Excel", command=self.abrir_excel)
        btn_abrir.pack(pady=5)

        # Botón modificar (deshabilitado hasta abrir archivo)
        self.btn_modificar = tk.Button(
            self, text="Limpiar E y Modificar F según G", command=self.ejecutar_proceso
        )
        self.btn_modificar.pack(pady=10)
        self.btn_modificar.config(state="disabled")

        # Botón abrir archivo modificado
        self.btn_abrir_modificado = tk.Button(
            self, text="Abrir Excel Modificado", command=self.abrir_modificado
        )
        self.btn_abrir_modificado.pack(pady=10)
        self.btn_abrir_modificado.config(state="disabled")

    def abrir_excel(self):
        """Permite seleccionar y abrir un archivo Excel."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm *.xltx *.xltm")]
        )
        if file_path:
            if self.excel_manager.abrir_archivo(file_path):
                self.label_archivo.config(text=f"Archivo: {os.path.basename(file_path)}")
                # Habilitar botón de modificar
                self.btn_modificar.config(state="normal")

    def ejecutar_proceso(self):
        """Ejecuta el proceso completo: limpiar E y modificar F según G."""
        if not self.excel_manager.file_path:
            messagebox.showwarning("Atención", "Primero abra un archivo Excel.")
            return

        modificado_path = self.excel_manager.limpiar_y_modificar("Hoja1")
        messagebox.showinfo("Proceso completado", f"Archivo guardado como:\n{modificado_path}")

        # Habilitar botón para abrir archivo modificado
        self.btn_abrir_modificado.config(state="normal")

    def abrir_modificado(self):
        """Abre el archivo modificado en Excel (Windows)."""
        if self.excel_manager.modificado_path and os.path.exists(self.excel_manager.modificado_path):
            try:
                # Windows
                if os.name == 'nt':
                    os.startfile(self.excel_manager.modificado_path)
                else:
                    # Mac/Linux
                    subprocess.call(['open', self.excel_manager.modificado_path])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showwarning("Atención", "No existe el archivo modificado.")


if __name__ == "__main__":
    app = ExcelApp()
    app.mainloop()
