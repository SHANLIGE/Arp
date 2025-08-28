from flask import Flask, request, render_template, send_file, redirect, url_for
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import os
import threading
import webview
from werkzeug.utils import secure_filename
import time
import ctypes
import pygetwindow as gw

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER




# Word generation
def generar_docx(df, hoja):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    for _, fila in df.iterrows():
        doc.add_paragraph(hoja, style='Heading 1').alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        for campo in df.columns:
            contenido = fila.get(campo, "N/A")
            texto = "N/A" if pd.isna(contenido) else str(contenido).strip()
            doc.add_paragraph(f"{campo}:", style='Heading 3')
            doc.add_paragraph(texto)
            doc.add_paragraph()
        doc.add_page_break()
    output_path = os.path.join(OUTPUT_FOLDER, 'documento.docx')
    doc.save(output_path)
    return output_path

# PDF generation
class PPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='Letter')
        self.set_auto_page_break(auto=True, margin=20)
        self.set_margins(left=20, top=20, right=20)
        self.set_font("Helvetica", "", 11)

    def add_pqr(self, pqr, campos, hoja):
        self.set_font("Helvetica", "B", 13)
        try:
            self.multi_cell(175, 10, hoja, align="C")
        except:
            self.multi_cell(175, 10, hoja.encode('latin-1', 'replace').decode('latin-1'), align="C")
        self.ln(3)

        for campo in campos:
            self.set_font("Helvetica", "B", 11)
            try:
                self.multi_cell(175, 8, f"{campo}:", align="L")
            except:
                self.multi_cell(175, 8, f"{campo}:".encode('latin-1', 'replace').decode('latin-1'), align="L")

            contenido = pqr.get(campo, "")
            self.set_font("Helvetica", "", 11)

            try:
                texto = "N/A" if pd.isna(contenido) else str(contenido).strip()
                texto = texto.encode('latin-1', 'replace').decode('latin-1')

                # Si el texto tiene palabras larguísimas (como URLs), inserta espacios
                if len(texto) > 100 and ' ' not in texto:
                    texto = '\n'.join([texto[i:i+100] for i in range(0, len(texto), 100)])

                if not texto.strip():
                    texto = " "

                self.multi_cell(175, 8, texto, align="L")
            except:
                self.multi_cell(175, 8, "[Error al mostrar este campo]", align="L")

            self.ln(2)




def generar_pdf(df, hoja):
    pdf = PPDF()
    for _, fila in df.iterrows():
        pdf.add_page()
        pdf.add_pqr(fila, df.columns, hoja)
    output_path = os.path.join(OUTPUT_FOLDER, 'documento.pdf')
    pdf.output(output_path)
    return output_path

@app.route('/', methods=['GET', 'POST'])
def index():

 

    global archivo_generado
    if request.method == 'POST':
        file = request.files.get('archivo')
        formato = request.form.get('formato')
        hoja = request.form.get('hoja')
        if file and formato and hoja:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            df = pd.read_excel(filepath, sheet_name=hoja)

            if formato.lower() == 'word':
                archivo_generado = generar_docx(df, hoja)
            else:
                archivo_generado = generar_pdf(df, hoja)

            return redirect(url_for('descargar'))

    return render_template('index.html', mensaje=None)


@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    file = request.files['archivo']
    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        xl = pd.ExcelFile(filepath)
        return {'sheets': xl.sheet_names}
    return {'sheets': []}


def run_flask():
    app.run(debug=True, use_reloader=False)


if __name__ == '__main__':
    threading.Thread(target=run_flask, daemon=True).start()
    time.sleep(1)

    

    class API:
        def __init__(self):
            self.maximized = False
            self.hwnd = None  # Manejador de la ventana

        def _get_hwnd(self):
            if not self.hwnd:
                time.sleep(0.2)  # Espera a que se cree la ventana
                self.hwnd = ctypes.windll.user32.FindWindowW(None, "Arp")
            return self.hwnd

        def close(self):
            hwnd = self._get_hwnd()
            if hwnd:
                ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE

        def minimize(self):
            hwnd = self._get_hwnd()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE

        def toggle_maximize(self):
            hwnd = self._get_hwnd()
            if hwnd:
                if self.maximized:
                    ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                else:
                    ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                self.maximized = not self.maximized

        def move(self, dx, dy):
            win = webview.windows[0]
            win.move(win.x + dx, win.y + dy)


            
  # --- RUTAS FLASK ---

@app.route('/loading')
def loading():
    return render_template('loading.html')

@app.route('/descargar')
def descargar():
    global archivo_generado
    return render_template('index.html', mensaje="✅ ¡Archivo generado con éxito! Puedes volver a convertir otro si lo deseas.", descarga=True)


# --- FLUJO DE VENTANAS ---
def show_main():
    time.sleep(5) 
    splash_window.destroy()
    time.sleep(0.1)
    webview.create_window(
        "Arp",
        "http://127.0.0.1:5000",
        width=1250,
        height=800,
        resizable=True,
        frameless=True,
        js_api=API()
    )


# --- ARRANQUE ---
if __name__ == '__main__':
    threading.Thread(target=app.run, kwargs={'debug': False, 'use_reloader': False}).start()
    time.sleep(1)

    splash_window = webview.create_window(
        "Cargando...",
        "http://127.0.0.1:5000/loading",
        width=600,
        height=400,
        frameless=True,
    )

    threading.Thread(target=show_main, daemon=True).start()
    webview.start(http_server=True)


    class API:
        def close(self):
            webview.windows[0].destroy()
            os._exit(0)  