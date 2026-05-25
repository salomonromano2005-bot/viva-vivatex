import sys
import json
import traceback
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, FormulaRule
from openpyxl.styles.differential import DifferentialStyle
import re

# ─── PALETA VIVATEX ───────────────────────────────────────────────────────────
V_DARK_GREEN   = "1E6B2E"
V_MID_GREEN    = "4A9A20"
V_LIGHT_GREEN  = "6BBF3E"
V_PALE_GREEN   = "EAF5E0"
V_PALE_GREEN2  = "D4EDBA"
V_WHITE        = "FFFFFF"
V_DARK_GRAY    = "2D2D2D"
V_MID_GRAY     = "666666"
V_LIGHT_GRAY   = "F5F5F5"
V_BORDER_GRAY  = "D0D0D0"
V_RED          = "C0392B"
V_ORANGE       = "E67E22"
V_YELLOW       = "F1C40F"
V_BLUE         = "2980B9"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=V_DARK_GRAY, italic=False, name="Calibri"):
    return Font(bold=bold, size=size, color=color, italic=italic, name=name)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin(sides="all"):
    s = Side(style="thin", color=V_BORDER_GRAY)
    n = Side(style=None)
    if sides == "all":
        return Border(left=s, right=s, top=s, bottom=s)
    if sides == "bottom":
        return Border(bottom=s)
    if sides == "top_bottom":
        return Border(top=s, bottom=s)
    return Border(left=s, right=s, top=s, bottom=s)

def border_medium():
    s = Side(style="medium", color=V_DARK_GREEN)
    return Border(left=s, right=s, top=s, bottom=s)

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def merge_and_write(ws, cell_range, value, fnt=None, fll=None, aln=None, brd=None):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = value
    if fnt: cell.font = fnt
    if fll: cell.fill = fll
    if aln: cell.alignment = aln
    if brd: cell.border = brd

def write_header_row(ws, row, headers, fill_color=V_DARK_GREEN, font_color=V_WHITE, start_col=1, row_height=22):
    ws.row_dimensions[row].height = row_height
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(row=row, column=i)
        c.value = h
        c.font = Font(bold=True, size=10, color=font_color, name="Calibri")
        c.fill = fill(fill_color)
        c.alignment = align("center")
        c.border = border_thin()

def write_data_row(ws, row, values, even=False, start_col=1, row_height=18, formats=None):
    ws.row_dimensions[row].height = row_height
    bg = V_PALE_GREEN if even else V_WHITE
    for i, v in enumerate(values, start=start_col):
        c = ws.cell(row=row, column=i)
        c.value = v
        c.fill = fill(bg)
        c.alignment = align("center")
        c.border = border_thin()
        if formats and i - start_col < len(formats) and formats[i - start_col]:
            c.number_format = formats[i - start_col]

def semaforo_color(estatus):
    e = str(estatus).upper()
    if "CRÍTICO" in e or "CRITICO" in e: return V_RED
    if "ALERTA" in e: return V_ORANGE
    if "VIGILANCIA" in e: return V_YELLOW
    if "META" in e or "✔" in e: return V_LIGHT_GREEN
    return V_LIGHT_GRAY

def add_portada(wb, titulo, subtitulo, empresa, periodo, confidencial):
    ws = wb.active
    ws.title = "Portada"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 20
    ws.row_dimensions[1].height = 8

    # Franja superior verde oscuro
    for r in range(2, 8):
        ws.row_dimensions[r].height = 18
        for col in range(1, 4):
            ws.cell(row=r, column=col).fill = fill(V_DARK_GREEN)

    ws.merge_cells("B2:C7")
    c = ws["B2"]
    c.value = empresa.upper()
    c.font = Font(bold=True, size=22, color=V_WHITE, name="Calibri")
    c.alignment = align("left", "center")
    c.fill = fill(V_DARK_GREEN)

    # Línea decorativa verde claro
    for col in range(1, 4):
        ws.cell(row=8, column=col).fill = fill(V_LIGHT_GREEN)
    ws.row_dimensions[8].height = 6

    # Título principal
    ws.row_dimensions[10].height = 10
    ws.merge_cells("B11:C11")
    c = ws["B11"]
    c.value = titulo.upper()
    c.font = Font(bold=True, size=18, color=V_DARK_GREEN, name="Calibri")
    c.alignment = align("left", "center")

    ws.merge_cells("B12:C12")
    c = ws["B12"]
    c.value = subtitulo
    c.font = Font(size=13, color=V_MID_GREEN, name="Calibri", italic=True)
    c.alignment = align("left", "center")
    ws.row_dimensions[12].height = 20

    ws.row_dimensions[14].height = 6
    for col in range(1, 4):
        ws.cell(row=14, column=col).fill = fill(V_PALE_GREEN2)

    # Datos del reporte
    datos = [
        ("📅 PERÍODO:", periodo),
        ("🏢 EMPRESA:", empresa),
        ("🔒 CLASIFICACIÓN:", confidencial),
    ]
    for i, (label, val) in enumerate(datos, start=16):
        ws.row_dimensions[i].height = 22
        c = ws.cell(row=i, column=2)
        c.value = label
        c.font = Font(bold=True, size=11, color=V_DARK_GREEN, name="Calibri")
        c.alignment = align("left")
        d = ws.cell(row=i, column=3)
        d.value = val
        d.font = Font(size=11, color=V_DARK_GRAY, name="Calibri")
        d.alignment = align("left")

    # Línea inferior
    for col in range(1, 4):
        ws.cell(row=22, column=col).fill = fill(V_DARK_GREEN)
    ws.row_dimensions[22].height = 5

    ws.merge_cells("B24:C24")
    c = ws["B24"]
    c.value = "Generado por AVIVA · Inteligencia Artificial de Grupo Vivatex S.A. de C.V."
    c.font = Font(size=9, color=V_MID_GRAY, italic=True, name="Calibri")
    c.alignment = align("center")

    ws.print_area = "A1:C30"
    return ws

def add_resumen_sheet(wb, data):
    ws = wb.create_sheet("📊 Resumen Ejecutivo")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 22

    # Título de sección
    ws.merge_cells("B1:F1")
    ws.row_dimensions[1].height = 8

    ws.merge_cells("B2:F2")
    c = ws["B2"]
    c.value = "📊 RESUMEN EJECUTIVO"
    c.font = Font(bold=True, size=14, color=V_WHITE, name="Calibri")
    c.fill = fill(V_DARK_GREEN)
    c.alignment = align("center")
    ws.row_dimensions[2].height = 28

    # Subtítulo
    ws.merge_cells("B3:F3")
    c = ws["B3"]
    c.value = data.get("subtitulo", "Indicadores Clave de Desempeño")
    c.font = Font(size=10, color=V_MID_GRAY, italic=True, name="Calibri")
    c.alignment = align("center")
    ws.row_dimensions[3].height = 18

    row = 5
    for seccion in data.get("secciones", []):
        # Encabezado de sección
        ws.merge_cells(f"B{row}:F{row}")
        c = ws[f"B{row}"]
        c.value = seccion.get("titulo", "")
        c.font = Font(bold=True, size=11, color=V_WHITE, name="Calibri")
        c.fill = fill(V_MID_GREEN)
        c.alignment = align("center")
        ws.row_dimensions[row].height = 22
        row += 1

        cols = seccion.get("columnas", [])
        if cols:
            write_header_row(ws, row, cols, start_col=2)
            row += 1

            for j, fila in enumerate(seccion.get("filas", [])):
                is_even = j % 2 == 0
                for k, val in enumerate(fila, start=2):
                    c = ws.cell(row=row, column=k)
                    c.value = val
                    bg = V_PALE_GREEN if is_even else V_WHITE

                    # Semáforo en columna de estatus
                    if "estatus" in (cols[k-2].lower() if k-2 < len(cols) else ""):
                        color = semaforo_color(str(val))
                        c.fill = fill(color)
                        c.font = Font(bold=True, size=9, color=V_WHITE if color in [V_RED, V_DARK_GREEN, V_MID_GREEN] else V_DARK_GRAY, name="Calibri")
                    else:
                        c.fill = fill(bg)
                        c.font = Font(size=10, color=V_DARK_GRAY, name="Calibri")

                    c.alignment = align("center")
                    c.border = border_thin()
                    ws.row_dimensions[row].height = 18
                row += 1

        row += 2

    return ws

def add_data_sheet(wb, hoja_data):
    nombre = hoja_data.get("nombre", "Datos")
    ws = wb.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    columnas = hoja_data.get("columnas", [])
    filas = hoja_data.get("filas", [])
    totales = hoja_data.get("totales", [])
    tipo_grafica = hoja_data.get("tipo_grafica", None)

    # Widths dinámicos
    ws.column_dimensions['A'].width = 2
    for i in range(len(columnas)):
        col_letter = get_column_letter(i + 2)
        ws.column_dimensions[col_letter].width = max(16, len(str(columnas[i])) + 4)

    ws.row_dimensions[1].height = 8

    # Título de hoja
    last_col = get_column_letter(len(columnas) + 1)
    ws.merge_cells(f"B2:{last_col}2")
    c = ws["B2"]
    c.value = nombre.replace("📈", "").replace("📊", "").strip().upper()
    c.font = Font(bold=True, size=13, color=V_WHITE, name="Calibri")
    c.fill = fill(V_DARK_GREEN)
    c.alignment = align("center")
    ws.row_dimensions[2].height = 26

    ws.row_dimensions[3].height = 6
    for col in range(2, len(columnas) + 2):
        ws.cell(row=3, column=col).fill = fill(V_LIGHT_GREEN)

    # Encabezados
    write_header_row(ws, 4, columnas, start_col=2)

    # Datos
    for j, fila in enumerate(filas):
        row = j + 5
        is_even = j % 2 == 0
        for k, val in enumerate(fila, start=2):
            c = ws.cell(row=row, column=k)
            try:
                num = float(str(val).replace(",", "").replace("$", "").replace("%", ""))
                c.value = num
                if "%" in str(val):
                    c.number_format = '0.00%'
                elif "$" in str(val) or (k > 2 and columnas[k-2] and "$" in str(columnas[k-2])):
                    c.number_format = '$#,##0.00'
                else:
                    c.number_format = '#,##0.00'
            except:
                c.value = val
            c.fill = fill(V_PALE_GREEN if is_even else V_WHITE)
            c.font = Font(size=10, color=V_DARK_GRAY, name="Calibri")
            c.alignment = align("center")
            c.border = border_thin()
            ws.row_dimensions[row].height = 18

    # Fila de totales
    if totales and len(totales) == len(columnas):
        tot_row = len(filas) + 5
        ws.row_dimensions[tot_row].height = 22
        for k, val in enumerate(totales, start=2):
            c = ws.cell(row=tot_row, column=k)
            c.value = val
            c.font = Font(bold=True, size=10, color=V_WHITE, name="Calibri")
            c.fill = fill(V_DARK_GREEN)
            c.alignment = align("center")
            c.border = border_thin()

    # Gráfica
    if tipo_grafica and filas and len(columnas) >= 2:
        data_start = 5
        data_end = len(filas) + 4
        num_cols = len(columnas)

        if tipo_grafica == "bar":
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
            chart.overlap = -10
        elif tipo_grafica == "line":
            chart = LineChart()
        elif tipo_grafica == "pie":
            chart = PieChart()
        elif tipo_grafica == "bar_stacked":
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "stacked"
        else:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"

        chart.title = nombre.replace("📈", "").replace("📊", "").strip()
        chart.style = 10
        chart.height = 14
        chart.width = 24

        # Categorías (primera columna de datos)
        cats = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)

        # Series numéricas
        added = 0
        for col_idx in range(3, num_cols + 2):
            try:
                data_ref = Reference(ws, min_col=col_idx, min_row=4, max_row=data_end)
                series = chart.series.__class__()
                chart.add_data(data_ref, titles_from_data=True)
                added += 1
                if tipo_grafica == "pie":
                    break
            except:
                pass

        if added == 0:
            data_ref = Reference(ws, min_col=3, min_row=4, max_row=data_end)
            chart.add_data(data_ref, titles_from_data=True)

        chart.set_categories(cats)

        if tipo_grafica not in ("pie",):
            chart.x_axis.title = columnas[0] if columnas else ""
            chart.y_axis.title = columnas[1] if len(columnas) > 1 else "Valor"
            chart.y_axis.numFmt = '#,##0'
            chart.x_axis.tickLblSkip = 1

        # Colores verde Vivatex para las series
        vivatex_colors = ["1E6B2E", "6BBF3E", "4A9A20", "A3D977", "0D4A1E", "8ED050"]
        for i, s in enumerate(chart.series):
            color = vivatex_colors[i % len(vivatex_colors)]
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.solidFill = color

        chart_row = data_end + 3
        ws.add_chart(chart, f"B{chart_row}")

    return ws

def generate_report(data_json):
    data = json.loads(data_json) if isinstance(data_json, str) else data_json
    wb = Workbook()

    titulo = data.get("titulo", "Reporte Vivatex")
    subtitulo = data.get("subtitulo", "Reporte Ejecutivo")
    empresa = data.get("empresa", "Grupo Vivatex S.A. de C.V.")
    periodo = data.get("periodo", "2026")
    confidencial = data.get("confidencial", "Confidencial · Uso Exclusivo de Dirección")

    # Portada
    add_portada(wb, titulo, subtitulo, empresa, periodo, confidencial)

    # Resumen ejecutivo si viene
    if "resumen" in data:
        add_resumen_sheet(wb, data["resumen"])

    # Hojas de datos
    for hoja in data.get("hojas", []):
        add_data_sheet(wb, hoja)

    return wb

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Uso: generate_excel.py <json_data> <output_path>"}))
        sys.exit(1)
    try:
        data_json = sys.argv[1]
        output_path = sys.argv[2]
        wb = generate_report(data_json)
        wb.save(output_path)
        print(json.dumps({"success": True, "path": output_path}))
    except Exception as e:
        print(json.dumps({"error": str(e), "trace": traceback.format_exc()}))
        sys.exit(1)
