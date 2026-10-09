import openpyxl

#cargar el archivo de Excel
wb = openpyxl.load_workbook('ventas_tecnologia_v2.xlsx', data_only=True)

#seleccionar la hoja de trabajo
hoja = wb.active

# print(f'La hoja activa es: {hoja.title}')

#crear una lista vacía para almacenar los datos de ventas
datos = []

#recorrer desde la fila 2
for fila in hoja.iter_rows(min_row=2, values_only=True):
    # # Verificar si la celda de la columna D (índice 3) es mayor a 100
    # if fila[3] is not None and fila[3] > 100:
        datos.append(list(fila))

# print(f'Total de registros leídos en memoria: {len(datos)}')

# suma_precio = sum(fila[3] for fila in datos)  # Sumar los precios de todos los productos
# promedio_precio = suma_precio / len(datos)  # Calcular el promedio de precios
# suma_ventas = sum(fila[4] for fila in datos)  # Sumar las ventas de todos los productos
# fila_final = ["TOTAL", "", "", suma_precio, suma_ventas]  # Crear la fila final con los totales
# datos.append(fila_final)  # Agregar la fila final a la lista de datos
from tabulate import tabulate

#nombres de columnas
encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]

#alineación por columna
alineacion = ("center", "left", "center", "right", "right")

id_producto = int(input("Ingrese el ID del producto a buscar: "))
for fila in datos:
    if fila[0] == id_producto:
       dato = [fila]  # Guardar la fila encontrada en una nueva lista
#mostrar tabla formateada en consola
print(f"\n=== REPORTE DE VENTAS EN CONSOLA ({len(dato)} PRODUCTOS) ===")
print(tabulate(dato, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

#Crear doc base
doc = SimpleDocTemplate("reporte_ejecutivo_ventas_reto_3.pdf", pagesize=letter)
estilos = getSampleStyleSheet()

#Estilos personalizados
estilo_titulo = ParagraphStyle(
    'EstiloTitulo', parent=estilos['Title'],
    fontName='Helvetica-Bold', fontSize=18,
    textColor=colors.HexColor("#1A365D"), spaceAfter=120
)
estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
estilo_encabezado = ParagraphStyle(
    'EstiloEncabezado', parent=estilos['Normal'],
    fontName='Helvetica-Bold', fontSize=10,
    textColor=colors.whitesmoke, alignment=1
)

#Lista story
story = [
    Paragraph("REPORTE EJECUTIVO DE VENTAS", estilo_titulo),
    Paragraph(f"<b>Total de productos procesados:</b> {len(dato)}", estilo_cuerpo),
    Spacer(1, 12)
]

#Convertir celdas a Paragraph
encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]
filas_pdf = [
    [
        Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
        Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
        Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
    ]
    for r in dato
]

#Configurar anchos y tabla
anchos_cols = [40, 160, 100, 80, 80]
tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths=anchos_cols,repeatRows=1)

#Aplicar estilo visual con TableStyle
tabla_pdf.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")])
]))

story.append(tabla_pdf) 

#Generar PDF
doc.build(story)