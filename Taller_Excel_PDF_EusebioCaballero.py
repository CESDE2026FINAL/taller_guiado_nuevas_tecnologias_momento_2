import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

ARCHIVO_EXCEL = "ventas_tecnologia_v2.xlsx"
ARCHIVO_PDF = "reporte_ejecutivo_ventas_reto_3.pdf"
ENCABEZADOS = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
ALINEACION = ("center", "left", "center", "right", "right")


# ---------- RETO 1: lectura de datos ----------
def cargar_datos(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    hoja = wb.active
    print(f"La hoja activa es: {hoja.title}")
    datos = [list(fila) for fila in hoja.iter_rows(min_row=2, values_only=True)]
    print(f"Total de registros leídos en memoria: {len(datos)}")
    return datos


# ---------- RETO 2: consola ----------
def mostrar_tabla(filas, titulo):
    print(f"\n=== {titulo} ({len(filas)} PRODUCTOS) ===")
    if not filas:
        print("No hay resultados para mostrar.")
        return
    print(tabulate(filas, headers=ENCABEZADOS, tablefmt="fancy_grid", colalign=ALINEACION))


def filtrar_por_precio(datos, minimo=100):
    return [f for f in datos if f[3] is not None and f[3] > minimo]


def buscar_por_id(datos, id_producto):
    return [f for f in datos if f[0] == id_producto]


# ---------- RETO 3: PDF ----------
def exportar_pdf(datos, ruta):
    doc = SimpleDocTemplate(ruta, pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=20
    )
    estilo_cuerpo = ParagraphStyle("EstiloCuerpo", parent=estilos["Normal"], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        "EstiloEncabezado", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.whitesmoke, alignment=1
    )
    centro = ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)
    izquierda = ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)
    derecha = ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)

    story = [
        Paragraph("REPORTE EJECUTIVO DE VENTAS", estilo_titulo),
        Paragraph(f"<b>Total de productos procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12),
    ]

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in ENCABEZADOS]
    filas_pdf = [
        [
            Paragraph(str(r[0]), centro),
            Paragraph(str(r[1]), izquierda),
            Paragraph(str(r[2]), centro),
            Paragraph(f"${r[3]:,.2f}", derecha),
            Paragraph(str(r[4]), derecha),
        ]
        for r in datos
    ]

    tabla = Table([encabezados_pdf] + filas_pdf,
                  colWidths=[40, 160, 100, 80, 80], repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
    ]))
    story.append(tabla)

    doc.build(story)
    print(f"\n✔ Informe PDF generado: {ruta}")


# ---------- MENÚ ----------
def mostrar_menu():
    print("\n" + "=" * 35)
    print("        MENÚ DE VENTAS")
    print("=" * 35)
    print("1) Ver catálogo completo")
    print("2) Filtrar por precio (> $100)")
    print("3) Buscar por ID")
    print("4) Exportar informe PDF")
    print("5) Salir")


def main():
    datos = cargar_datos(ARCHIVO_EXCEL)

    while True:
        mostrar_menu()
        opcion = input("Elige una opción (1-5): ").strip()

        if opcion == "1":
            mostrar_tabla(datos, "CATÁLOGO COMPLETO")

        elif opcion == "2":
            mostrar_tabla(filtrar_por_precio(datos), "PRODUCTOS CON PRECIO > $100")

        elif opcion == "3":
            try:
                id_producto = int(input("Ingrese el ID del producto a buscar: "))
            except ValueError:
                print("⚠ El ID debe ser un número entero.")
                continue
            resultado = buscar_por_id(datos, id_producto)
            if resultado:
                mostrar_tabla(resultado, f"RESULTADO PARA ID {id_producto}")
            else:
                print(f"No se encontró ningún producto con ID {id_producto}.")

        elif opcion == "4":
            exportar_pdf(datos, ARCHIVO_PDF)

        elif opcion == "5":
            print("Saliendo del programa. ¡Hasta pronto!")
            break

        else:
            print("⚠ Opción no válida. Intenta de nuevo.")


if __name__ == "__main__":
    main()