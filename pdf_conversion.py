"""
pdf_conversion.py — Conversión de libros Excel (.xlsx) a PDF.

Utiliza Microsoft Excel instalado en el equipo (automatización COM) para
exportar cada hoja a PDF respetando el área de impresión y la configuración
de página de la plantilla, de modo que el PDF conserve la estructura de
páginas del SISCA y quede listo para imprimir o editar.
"""

import os

import pythoncom
from win32com.client import dynamic


def _iniciar_com():
    try:
        pythoncom.CoInitialize()
        return True
    except Exception:
        return False


def _cerrar_com():
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def xlsx_a_pdf(ruta_xlsx, ruta_pdf):
    """Convierte ``ruta_xlsx`` a PDF en ``ruta_pdf``.

    Retorna ``(True, None)`` en éxito o ``(False, mensaje_error)``.
    """
    if not os.path.exists(ruta_xlsx):
        return False, "El archivo Excel no existe."

    _iniciar_com()
    excel = None
    try:
        excel = dynamic.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        ruta_xlsx_abs = os.path.abspath(ruta_xlsx)
        ruta_pdf_abs = os.path.abspath(ruta_pdf)

        try:
            libro = excel.Workbooks.Open(ruta_xlsx_abs, False, True)
        except Exception:
            libro = excel.Workbooks.Open(ruta_xlsx_abs, False, False)

        try:
            libro.ExportAsFixedFormat(0, ruta_pdf_abs)
        finally:
            try:
                libro.Close(SaveChanges=False)
            except Exception:
                pass
    except Exception as exc:
        return False, f"No se pudo convertir a PDF (verifica que Excel esté instalado): {exc}"
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        _cerrar_com()

    if os.path.exists(ruta_pdf) and os.path.getsize(ruta_pdf) > 0:
        return True, None
    return False, "La conversión a PDF no produjo un archivo válido."
