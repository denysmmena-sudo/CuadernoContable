# -*- coding: utf-8 -*-
import datetime, os, sqlite3, webbrowser, zipfile, base64
import flet as ft
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import pandas as pd

NOMBRE_COMERCIO = "ComercialPablitto"
ARCHIVO_LOGO = "logo.png"  # Asegúrate de que se llame exactamente logo.png
ultimo_pdf_generado = None
LOG_ENVIADOS = "ventas_enviadas.log"


def obtener_logo_base64():
    if os.path.exists(ARCHIVO_LOGO):
        try:
            with open(ARCHIVO_LOGO, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode("utf-8")
        except Exception:
            return None
    return None


LOGO_BASE64 = obtener_logo_base64()


def inicializar_db_con_base():
    conn = sqlite3.connect("comercio_ventas.db")
    cursor = conn.cursor()

    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS inventario
                   (
                       codigo
                       TEXT
                       PRIMARY
                       KEY,
                       descripcion
                       TEXT,
                       unidad
                       TEXT,
                       precio
                       REAL,
                       stock
                       REAL
                   )
                   """)

    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS ventas
                   (
                       id
                       INTEGER
                       PRIMARY
                       KEY
                       AUTOINCREMENT,
                       cliente
                       TEXT,
                       fecha
                       TEXT,
                       codigo
                       TEXT,
                       cantidad
                       REAL,
                       total
                       REAL
                   )
                   """)

    archivo_excel = "CatalogoProductos_ComercialPablitto.xlsx"

    if os.path.exists(archivo_excel):
        try:
            df = pd.read_excel(archivo_excel, skiprows=2)
            for _, fila in df.iterrows():
                codigo = str(fila.get("CODIGO", "")).strip()
                if codigo and codigo.startswith("PR"):
                    descripcion = str(fila.get("DESCRIPCION", ""))
                    unidad = str(fila.get("UNID/MED", "UND"))
                    precio = float(fila.get("MONTO VENTA", 0.0))

                    cursor.execute("SELECT stock FROM inventario WHERE codigo = ?", (codigo,))
                    resultado = cursor.fetchone()
                    stock_inicial = 100.0 if resultado is None else resultado[0]

                    cursor.execute("""
                        INSERT OR REPLACE INTO inventario (codigo, descripcion, unidad, precio, stock)
                        VALUES (?, ?, ?, ?, ?)
                    """, (codigo, descripcion, unidad, precio, stock_inicial))

            conn.commit()
            print("¡Catálogo sincronizado correctamente!")
        except Exception as e:
            print(f"Error al procesar el archivo Excel: {e}")
    conn.close()


def validar_cliente(v):
    if not v: return "El nombre del cliente no puede estar vacío."
    if v.isdigit():
        if len(v) not in (8, 11): return "DNI debe tener 8 dígitos o RUC 11 dígitos."
    elif not all(c.isalpha() or c.isspace() for c in v):
        return "Error: El nombre del cliente no debe contener números."
    return None


def convertir_unidad(u):
    u = u.strip().upper()
    return "UNIDAD" if u in ("NIU", "MTR") else ("GALON" if u == "GLL" else u)


def generar_pdf_nota_multiples(cliente, items, venta_id):
    global ultimo_pdf_generado
    fecha_str = datetime.datetime.now().strftime("%d.%m.%Y")
    num_fmt = f"{venta_id:06d}"
    filename = f"NV_ComercialPablitto_{num_fmt}_{fecha_str}.pdf"
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    fecha_emision = datetime.datetime.now().strftime("%d/%m/%Y")

    c.line(40, height - 20, width - 40, height - 20)

    # Se omite el dibujo del logo en la esquina superior izquierda

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2.0, height - 35, NOMBRE_COMERCIO)
    c.line(40, height - 44, width - 40, height - 44)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, height - 64, f"NOTA DE VENTA_{num_fmt}")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, height - 89, "Fecha de Emisión")
    c.drawString(160, height - 89, f": {fecha_emision}")
    c.drawString(40, height - 104, "Señor(es)")

    # Si es "SIN NOMBRE" o está vacío, se muestra un guión "-"
    cliente_str = cliente if cliente and cliente != "SIN NOMBRE" else "-"
    c.drawString(160, height - 104, f": {cliente_str}")

    c.drawString(40, height - 119, "Tipo de Moneda")
    c.drawString(160, height - 119, ": SOLES")
    c.drawString(40, height - 134, "Observación")
    c.drawString(160, height - 134, ": -")
    c.line(40, height - 149, width - 40, height - 149)

    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, height - 164, "N°")
    c.drawString(65, height - 164, "Cantidad")
    c.drawString(125, height - 159, "Unidad")
    c.drawString(125, height - 170, "Medida")
    c.drawString(175, height - 164, "Código")
    c.drawString(225, height - 164, "Descripción")
    c.drawCentredString(488, height - 161, "Valor (*)")
    c.drawCentredString(488, height - 172, "Unitario")
    c.drawCentredString(540, height - 161, "Importe (**)")
    c.drawString(517, height - 172, "de Venta")
    c.line(40, height - 179, width - 40, height - 179)

    y_pos, subtotal_general, descuento_total = 194, 0.0, 0.0
    for idx, item in enumerate(items, start=1):
        codigo, desc_base, unidad = item["codigo"], item["descripcion"], convertir_unidad(item["unidad"])
        cant, precio_cat, dcto = item["cantidad"], item["precio"], item["descuento"]
        importe_bruto = (precio_cat * 1.18) * cant
        subtotal_general += importe_bruto
        descuento_total += dcto

        c.setFont("Helvetica", 9)
        c.drawString(40, height - y_pos, f"{idx}")
        c.drawRightString(100, height - y_pos, f"{cant:.2f}")
        c.drawString(125, height - y_pos, unidad)
        c.drawString(175, height - y_pos, codigo)
        c.drawString(225, height - y_pos, desc_base)
        c.drawRightString(505, height - y_pos, f"{precio_cat:.4f}")
        c.drawRightString(550, height - y_pos, f"{importe_bruto:.2f}")
        if dcto > 0:
            c.setFont("Helvetica-Oblique", 7)
            c.setFillColorRGB(0.3, 0.3, 0.3)
            c.drawString(552, height - y_pos, f"(dcto. {dcto:.2f})")
            c.setFillColorRGB(0, 0, 0)
        y_pos += 20

    total_general = subtotal_general - descuento_total
    y_box = height - y_pos + 5
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(455, y_box, "Subtotal :")
    c.drawString(467, y_box, "S/")
    c.drawRightString(550, y_box, f"{subtotal_general:.2f}")
    y_box -= 15
    c.drawRightString(455, y_box, "Descuento Total :")
    c.drawString(467, y_box, "S/")
    c.drawRightString(550, y_box, f"{descuento_total:.2f}")
    y_box -= 12
    c.line(40, y_box, width - 40, y_box)

    y_tributos = y_box - 25
    c.setFont("Helvetica", 8)
    c.drawString(40, y_tributos, "(*) Sin impuestos.")
    c.drawString(40, y_tributos - 12, "(**) Con IGV.")
    op_gravada, igv_calc = total_general / 1.18, total_general - (total_general / 1.18)
    valores_tributos = [f"{op_gravada:.2f}", f"{igv_calc:.2f}"]
    labels_tributos = ["Op. Gravada", "IGV"]

    current_box_y = y_tributos - 5
    for i, label in enumerate(labels_tributos):
        cy = current_box_y - (i * 20)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(365, cy, f"{label} :")
        c.rect(455, cy - 3, 105, 14)
        c.drawString(467, cy, "S/")
        c.setFont("Helvetica", 9)
        c.drawRightString(550, cy, valores_tributos[i])

    final_y_line = current_box_y - 28
    c.line(345, final_y_line, 570, final_y_line)
    final_rect_bottom = current_box_y - 47
    final_y = final_rect_bottom + 3
    c.setFont("Helvetica-Bold", 9)
    c.drawString(365, final_y, "Importe Total :")
    c.rect(455, final_rect_bottom, 105, 14)
    c.drawString(467, final_y, "S/")
    c.drawRightString(550, final_y, f"{total_general:.2f}")

    # Línea colocada justo debajo del bloque de Importe Total
    c.line(40, final_rect_bottom - 20, width - 40, final_rect_bottom - 20)

    c.save()
    ultimo_pdf_generado = filename
    return filename


def main(page: ft.Page):
    page.title = f"{NOMBRE_COMERCIO} - Cuaderno de Ventas"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = "#2F3133"

    inicializar_db_con_base()
    carrito, inputs_stock_dict = [], {}

    chk_sin_nombre = ft.Checkbox(
        label="Sin nombre",
        value=True,
        label_style=ft.TextStyle(color=ft.Colors.BROWN_900, size=13, weight=ft.FontWeight.BOLD)
    )

    def on_change_cliente(e):
        if txt_cliente.value and txt_cliente.value.strip() != "":
            if chk_sin_nombre.value:
                chk_sin_nombre.value = False
                page.update()

    txt_cliente = ft.TextField(
        label="Nombre del Cliente o DNI/RUC",
        width=290,
        capitalization=ft.TextCapitalization.CHARACTERS,
        bgcolor=ft.Colors.WHITE,
        value="",
        on_change=on_change_cliente
    )

    def cambiar_estado_sin_nombre(e):
        if chk_sin_nombre.value:
            txt_cliente.value = ""
        page.update()

    chk_sin_nombre.on_change = cambiar_estado_sin_nombre

    txt_codigo = ft.TextField(label="Código", width=140, prefix_text="PR", value="",
                              keyboard_type=ft.KeyboardType.NUMBER, bgcolor=ft.Colors.WHITE)
    txt_cantidad = ft.TextField(label="Cantidad", width=95, keyboard_type=ft.KeyboardType.NUMBER, value="1.0",
                                bgcolor=ft.Colors.WHITE)
    txt_descuento = ft.TextField(label="Descuento", width=100, keyboard_type=ft.KeyboardType.NUMBER, value="0.0",
                                 bgcolor=ft.Colors.WHITE)
    lbl_resultado = ft.Text("", color=ft.Colors.GREEN_800, weight=ft.FontWeight.BOLD)
    tabla_carrito = ft.Column(controls=[], spacing=0, scroll=ft.ScrollMode.ALWAYS, expand=True)
    lbl_total_acumulado = ft.Text("Total de venta: S/ 0.00", weight=ft.FontWeight.BOLD, size=15,
                                  color=ft.Colors.BROWN_900)
    txt_info_pdf = ft.Text("", size=13, color=ft.Colors.BLACK87)
    lbl_resultado_stock = ft.Text("", color=ft.Colors.GREEN_800, weight=ft.FontWeight.BOLD)
    contenedor_lista_stock = ft.Column(controls=[], spacing=0, scroll=ft.ScrollMode.ALWAYS, expand=True)
    hoja_con_renglones = ft.Column(expand=True)

    def actualizar_vista_carrito():
        tabla_carrito.controls.clear()
        total_acum = 0.0
        for index, item in enumerate(carrito):
            subtotal_item = ((item["precio"] * 1.18) * item["cantidad"]) - item["descuento"]
            total_acum += subtotal_item
            desc_text = f" (dcto. S/ {item['descuento']:.2f})" if item["descuento"] > 0 else ""

            def crear_accion(idx):
                return lambda e: eliminar_item_carrito(idx)

            tabla_carrito.controls.append(
                ft.Container(
                    content=ft.Row([
                        ft.Text(f"{index + 1}. [{item['codigo']}] {item['descripcion']}{desc_text}", width=240,
                                color=ft.Colors.BLACK87),
                        ft.Text(f"Cant: {item['cantidad']} {item['unidad']}", width=80, color=ft.Colors.BLACK54),
                        ft.Text(f"S/ {subtotal_item:.2f}", width=65, weight=ft.FontWeight.BOLD,
                                color=ft.Colors.BROWN_900),
                        ft.IconButton(icon=ft.Icons.CLOSE, icon_color=ft.Colors.RED, tooltip="Eliminar producto",
                                      on_click=crear_accion(index)),
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    padding=ft.padding.symmetric(horizontal=8, vertical=6),
                    border=ft.border.only(bottom=ft.border.BorderSide(width=1, color=ft.Colors.GREY_300)),
                )
            )
        lbl_total_acumulado.value = f"Total de venta: S/ {total_acum:.2f}"
        page.update()

    def eliminar_item_carrito(index):
        if 0 <= index < len(carrito):
            eliminado = carrito.pop(index)
            lbl_resultado.value = f"Se eliminó el producto {eliminado['codigo']} del carrito."
            lbl_resultado.color = ft.Colors.ORANGE_800
            actualizar_vista_carrito()

    def anotar_producto_al_carrito(e):
        if not chk_sin_nombre.value:
            cliente_val = txt_cliente.value.strip().upper() if txt_cliente.value else ""
            err = validar_cliente(cliente_val)
            if err:
                lbl_resultado.value = err
                lbl_resultado.color = ft.Colors.RED
                page.update()
                return
        if not txt_codigo.value or not txt_cantidad.value:
            lbl_resultado.value = "Por favor completa el código y la cantidad."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return
        try:
            cant, desc = float(txt_cantidad.value), float(txt_descuento.value) if txt_descuento.value else 0.0
        except ValueError:
            lbl_resultado.value = "Cantidad o descuento deben ser números válidos."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return

        codigo_completo = f"PR{txt_codigo.value.strip().zfill(3)}"
        conn = sqlite3.connect("comercio_ventas.db")
        cursor = conn.cursor()
        cursor.execute("SELECT descripcion, unidad, precio, stock FROM inventario WHERE codigo = ?", (codigo_completo,))
        prod = cursor.fetchone()
        conn.close()

        if not prod:
            lbl_resultado.value = f"Error: Producto '{codigo_completo}' no encontrado."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return
        descripcion, unidad, precio, stock = prod
        if desc > (precio * cant):
            lbl_resultado.value = "El descuento no puede ser mayor al valor total."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return

        cant_en_carrito = sum(i["cantidad"] for i in carrito if i["codigo"] == codigo_completo)
        if stock < (cant_en_carrito + cant):
            lbl_resultado.value = f"Stock insuficiente. Disponible: {stock}"
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return

        carrito.append({"codigo": codigo_completo, "descripcion": descripcion, "unidad": unidad, "cantidad": cant,
                        "precio": precio, "descuento": desc})
        lbl_resultado.value = f"Producto {codigo_completo} ({descripcion}) anotado."
        lbl_resultado.color = ft.Colors.GREEN_700
        txt_codigo.value = ""
        txt_cantidad.value = "1.0"
        txt_descuento.value = "0.0"
        actualizar_vista_carrito()

    def procesar_venta_final(e):
        if chk_sin_nombre.value or not txt_cliente.value.strip():
            cliente_val = "SIN NOMBRE"
        else:
            cliente_val = txt_cliente.value.strip().upper()
            err = validar_cliente(cliente_val)
            if err:
                lbl_resultado.value = err
                lbl_resultado.color = ft.Colors.RED
                page.update()
                return

        if not carrito:
            lbl_resultado.value = "El carrito está vacío."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return

        conn = sqlite3.connect("comercio_ventas.db")
        cursor = conn.cursor()
        try:
            primer_total = (carrito[0]["precio"] * carrito[0]["cantidad"]) - carrito[0]["descuento"]
            cursor.execute(
                "INSERT INTO ventas (cliente, fecha, codigo, cantidad, total) VALUES (?, date('now'), ?, ?, ?)",
                (cliente_val, carrito[0]["codigo"], carrito[0]["cantidad"], primer_total))
            venta_id = cursor.lastrowid

            for item in carrito:
                total_item = (item["precio"] * item["cantidad"]) - item["descuento"]
                if item != carrito[0]:
                    cursor.execute(
                        "INSERT INTO ventas (cliente, fecha, codigo, cantidad, total) VALUES (?, date('now'), ?, ?, ?)",
                        (cliente_val, item["codigo"], item["cantidad"], total_item))
                cursor.execute("UPDATE inventario SET stock = stock - ? WHERE codigo = ?",
                               (item["cantidad"], item["codigo"]))

            conn.commit()
            pdf_nombre = generar_pdf_nota_multiples(cliente_val, carrito, venta_id)
            lbl_resultado.value = f"¡Venta registrada y PDF generado ({pdf_nombre})!"
            lbl_resultado.color = ft.Colors.GREEN_700
            carrito.clear()
            actualizar_vista_carrito()
        except Exception as ex:
            conn.rollback()
            lbl_resultado.value = f"Error al procesar la venta: {str(ex)}"
            lbl_resultado.color = ft.Colors.RED
        finally:
            conn.close()
            page.update()

    def enviar_ventas_zip(e):
        carpeta_destino = "Ventas_Comprimidas"
        os.makedirs(carpeta_destino, exist_ok=True)
        enviados_set = set()
        if os.path.exists(LOG_ENVIADOS):
            with open(LOG_ENVIADOS, "r", encoding="utf-8") as f:
                enviados_set = set(line.strip() for line in f if line.strip())

        todos_pdf = [f for f in os.listdir(".") if f.startswith("NV_ComercialPablitto_") and f.endswith(".pdf")]
        nuevos_pdf = [f for f in todos_pdf if f not in enviados_set]
        if not nuevos_pdf:
            lbl_resultado.value = "No hay nuevas ventas por enviar."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return

        fecha_hora_str = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        nombre_zip = os.path.join(carpeta_destino, f"Ventas_Nuevas_{fecha_hora_str}.zip")
        try:
            with zipfile.ZipFile(nombre_zip, "w") as zipf:
                for pdf in nuevos_pdf: zipf.write(pdf, os.path.basename(pdf))
            with open(LOG_ENVIADOS, "a", encoding="utf-8") as f:
                for pdf in nuevos_pdf: f.write(pdf + "\n")
            lbl_resultado.value = f"¡Ventas empaquetadas en: {nombre_zip}!"
            lbl_resultado.color = ft.Colors.BLUE_800
        except Exception as ex:
            lbl_resultado.value = f"Error comprimiendo: {str(ex)}"
            lbl_resultado.color = ft.Colors.RED
        page.update()

    def abrir_carpeta_pdf(e):
        ruta = os.path.abspath(".")
        try:
            if os.name == "nt":
                os.startfile(ruta)
            elif os.name == "posix":
                import subprocess, sys
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", ruta])
            lbl_resultado.value = "Carpeta abierta."
            lbl_resultado.color = ft.Colors.BLUE_800
        except Exception as ex:
            lbl_resultado.value = f"Error: {str(ex)}"
            lbl_resultado.color = ft.Colors.RED
        page.update()

    def abrir_visor_externo(e):
        if ultimo_pdf_generado and os.path.exists(ultimo_pdf_generado):
            webbrowser.open(os.path.abspath(ultimo_pdf_generado))

    dlg_ver_pdf = ft.AlertDialog(
        title=ft.Row([ft.Icon(ft.Icons.PICTURE_AS_PDF, color=ft.Colors.RED_700),
                      ft.Text("Previsualización de Nota PDF", weight=ft.FontWeight.BOLD)]),
        content=ft.Column(
            [ft.Text("El documento se ha generado correctamente:", size=13), txt_info_pdf, ft.Container(height=10),
             ft.Text("Ábrelo abajo:", size=12, italic=True)], tight=True, width=380, height=130),
        actions=[
            ft.TextButton("Cerrar", on_click=lambda e: setattr(dlg_ver_pdf, "open", False) or page.update()),
            ft.ElevatedButton("Abrir Visor PDF", icon=ft.Icons.LAUNCH, on_click=abrir_visor_externo,
                              bgcolor=ft.Colors.BLUE_GREY_800, color=ft.Colors.WHITE),
        ],
    )

    def abrir_dialog_pdf(e):
        if not ultimo_pdf_generado or not os.path.exists(ultimo_pdf_generado):
            lbl_resultado.value = "No hay ningún PDF disponible."
            lbl_resultado.color = ft.Colors.RED
            page.update()
            return
        txt_info_pdf.value = f"Archivo: {os.path.abspath(ultimo_pdf_generado)}"
        if dlg_ver_pdf not in page.overlay: page.overlay.append(dlg_ver_pdf)
        dlg_ver_pdf.open = True
        page.update()

    def construir_vista_stock():
        inputs_stock_dict.clear()
        conn = sqlite3.connect("comercio_ventas.db")
        cursor = conn.cursor()
        cursor.execute("SELECT codigo, descripcion, stock FROM inventario ORDER BY codigo ASC")
        filas = cursor.fetchall()
        conn.close()

        controles = []
        for codigo, descripcion, stock_val in filas:
            txt_stk = ft.TextField(value=str(stock_val), width=90, height=35, text_size=12,
                                   keyboard_type=ft.KeyboardType.NUMBER, bgcolor=ft.Colors.WHITE, content_padding=10,
                                   read_only=True)
            inputs_stock_dict[codigo] = txt_stk
            controles.append(
                ft.Container(
                    content=ft.Row([
                        ft.Text(codigo, width=70, weight=ft.FontWeight.BOLD, color=ft.Colors.BLACK87),
                        ft.Text(descripcion, width=250, color=ft.Colors.BLACK87),
                        txt_stk
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    padding=ft.padding.symmetric(horizontal=8, vertical=4),
                    border=ft.border.only(bottom=ft.border.BorderSide(width=1, color=ft.Colors.GREY_300)),
                )
            )
        return controles

    def habilitar_edicion_stock(e):
        for t in inputs_stock_dict.values(): t.read_only = False
        lbl_resultado_stock.value = "Modo edición activado."
        lbl_resultado_stock.color = ft.Colors.BLUE_800
        page.update()

    def actualizar_stock_db(e):
        conn = sqlite3.connect("comercio_ventas.db")
        cursor = conn.cursor()
        try:
            for cod, t in inputs_stock_dict.items():
                cursor.execute("UPDATE inventario SET stock = ? WHERE codigo = ?", (float(t.value), cod))
                t.read_only = True
            conn.commit()
            lbl_resultado_stock.value = "¡Stock actualizado y bloqueado!"
            lbl_resultado_stock.color = ft.Colors.GREEN_700
        except Exception as ex:
            conn.rollback()
            lbl_resultado_stock.value = f"Error: {str(ex)}"
            lbl_resultado_stock.color = ft.Colors.RED
        finally:
            conn.close()
            page.update()

    btn_anotar = ft.ElevatedButton("Anotar", icon=ft.Icons.EDIT, on_click=anotar_producto_al_carrito,
                                   color=ft.Colors.WHITE, bgcolor="#363839")
    btn_registrar_pdf = ft.ElevatedButton("Registrar y Generar Nota PDF", on_click=procesar_venta_final,
                                          color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN_700)
    btn_ver_pdf = ft.ElevatedButton("Ver PDF", icon=ft.Icons.PICTURE_AS_PDF, on_click=abrir_dialog_pdf,
                                    color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_GREY_800)
    btn_enviar_ventas = ft.ElevatedButton("Enviar Ventas", icon=ft.Icons.FOLDER_ZIP, on_click=enviar_ventas_zip,
                                          color=ft.Colors.WHITE, bgcolor=ft.Colors.INDIGO_700)
    btn_home_carpeta = ft.IconButton(icon=ft.Icons.HOME, icon_color=ft.Colors.WHITE, bgcolor="#4A2E18",
                                     tooltip="Ver notas", on_click=abrir_carpeta_pdf, icon_size=28)

    def cambiar_a_vista_ventas(e):
        hoja_con_renglones.controls.clear()
        hoja_con_renglones.controls.append(contenido_venta_vista)
        page.update()

    def cambiar_a_vista_stock(e):
        contenedor_lista_stock.controls = construir_vista_stock()
        lbl_resultado_stock.value = ""
        hoja_con_renglones.controls.clear()
        hoja_con_renglones.controls.append(contenido_stock_vista)
        page.update()

    btn_stock = ft.ElevatedButton("Stock", icon=ft.Icons.INVENTORY, on_click=cambiar_a_vista_stock,
                                  color=ft.Colors.WHITE, bgcolor=ft.Colors.TEAL_800)

    header_content = ft.Container(
        content=ft.Row([
            ft.Container(
                content=ft.Image(src_base64=LOGO_BASE64, width=40, height=40,
                                 fit=ft.ImageFit.CONTAIN) if LOGO_BASE64 else ft.Icon(ft.Icons.STORE,
                                                                                      color=ft.Colors.WHITE, size=30),
                bgcolor=ft.Colors.WHITE, border_radius=4, padding=2
            ),
            ft.Column([
                ft.Text(NOMBRE_COMERCIO, size=18, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                ft.Text("Registro Diario de Ventas", size=13, weight=ft.FontWeight.W_600, color=ft.Colors.WHITE70),
            ], spacing=0, alignment=ft.MainAxisAlignment.CENTER)
        ], alignment=ft.MainAxisAlignment.CENTER, spacing=15),
        bgcolor="#363839", border_radius=6, padding=ft.padding.symmetric(horizontal=16, vertical=10)
    )

    logo_inferior = ft.Container(
        content=ft.Column([
            ft.Image(src_base64=LOGO_BASE64, width=150, height=40,
                     fit=ft.ImageFit.CONTAIN) if LOGO_BASE64 else ft.Container(),
        ], alignment=ft.MainAxisAlignment.START, spacing=4),
        alignment=ft.alignment.center_left,
        padding=ft.padding.only(left=5, top=20)
    )

    contenido_venta_vista = ft.Column([
        header_content,
        ft.Row([txt_cliente, chk_sin_nombre], alignment=ft.MainAxisAlignment.START, spacing=10),
        ft.Row([txt_codigo, txt_cantidad, txt_descuento, btn_anotar], alignment=ft.MainAxisAlignment.START, spacing=15),
        ft.Text("Productos en la Nota de Venta actual:", weight=ft.FontWeight.BOLD, color=ft.Colors.BROWN_800),
        ft.Container(content=tabla_carrito, bgcolor=ft.Colors.WHITE, border=ft.border.all(1, ft.Colors.BROWN_300),
                     border_radius=4, height=140, padding=ft.padding.only(right=2)),
        lbl_total_acumulado,
        ft.Row([btn_registrar_pdf, btn_ver_pdf], alignment=ft.MainAxisAlignment.START, spacing=10),
        lbl_resultado,
        logo_inferior,
    ], spacing=8, expand=True)

    header_stock_content = ft.Container(
        content=ft.Row([
            ft.Container(
                content=ft.Image(src_base64=LOGO_BASE64, width=35, height=35,
                                 fit=ft.ImageFit.CONTAIN) if LOGO_BASE64 else ft.Icon(ft.Icons.INVENTORY,
                                                                                      color=ft.Colors.WHITE, size=28),
                bgcolor=ft.Colors.WHITE, border_radius=4, padding=2
            ),
            ft.Text("Gestión de Inventario (100 Productos)", size=15, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE)
        ], alignment=ft.MainAxisAlignment.CENTER, spacing=15),
        bgcolor="#363839", border_radius=6, padding=ft.padding.symmetric(horizontal=16, vertical=10)
    )

    contenido_stock_vista = ft.Column([
        header_stock_content,
        ft.Row([ft.Text("Código", width=70, weight=ft.FontWeight.BOLD, color=ft.Colors.BROWN_900),
                ft.Text("Descripción", width=250, weight=ft.FontWeight.BOLD, color=ft.Colors.BROWN_900),
                ft.Text("Cantidad Existente", width=100, weight=ft.FontWeight.BOLD, color=ft.Colors.BROWN_900)],
               alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        ft.Container(content=contenedor_lista_stock, bgcolor=ft.Colors.WHITE,
                     border=ft.border.all(1, ft.Colors.BROWN_300), border_radius=4, height=320,
                     padding=ft.padding.only(right=2)),
        ft.Row([ft.ElevatedButton("Añadir / Editar Stock", icon=ft.Icons.EDIT_NOTE, on_click=habilitar_edicion_stock,
                                  color=ft.Colors.WHITE, bgcolor=ft.Colors.AMBER_900),
                ft.ElevatedButton("Actualizar Stock", icon=ft.Icons.SAVE, on_click=actualizar_stock_db,
                                  color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN_700)],
               alignment=ft.MainAxisAlignment.START, spacing=10),
        lbl_resultado_stock,
    ], spacing=10, expand=True)

    hoja_con_renglones.controls.append(contenido_venta_vista)

    def crear_argolla():
        return ft.Container(content=ft.Row([
            ft.Container(bgcolor="#B0B0B0", width=4, height=12, border_radius=2),
            ft.Container(bgcolor="#E0E0E0", border=ft.border.all(1, "#888888"), border_radius=6, width=28, height=12,
                         gradient=ft.LinearGradient(begin=ft.alignment.top_center, end=ft.alignment.bottom_center,
                                                    colors=["#FFFFFF", "#CCCCCC", "#888888", "#DDDDDD"]),
                         shadow=ft.BoxShadow(blur_radius=2, color=ft.Colors.BLACK45, offset=ft.Offset(0, 2))),
        ], spacing=0), width=32, height=16)

    agenda_marco = ft.Container(
        content=ft.Row([
            ft.Stack([
                ft.Container(bgcolor="#6B4423", border_radius=ft.border_radius.horizontal(left=6, right=0), width=25,
                             height=620,
                             shadow=ft.BoxShadow(blur_radius=4, color=ft.Colors.BLACK54, offset=ft.Offset(-2, 0))),
                ft.Container(
                    content=ft.Column([crear_argolla() for _ in range(9)], alignment=ft.MainAxisAlignment.SPACE_AROUND,
                                      spacing=0, width=32), padding=ft.padding.only(left=2, top=10)),
            ]),
            ft.Container(content=hoja_con_renglones, padding=15, expand=True),
            ft.Container(content=ft.Column([
                ft.Text("Acciones", weight=ft.FontWeight.BOLD, color=ft.Colors.BROWN_900),
                btn_enviar_ventas, btn_stock, ft.Container(height=5),
                ft.Column([ft.Text("Ver Ventas", size=11, color=ft.Colors.BROWN_800, weight=ft.FontWeight.W_500),
                           btn_home_carpeta], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=4),
                ft.TextButton("Volver a Venta", icon=ft.Icons.ARROW_BACK, on_click=cambiar_a_vista_ventas),
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8),
                padding=20, alignment=ft.alignment.center),
        ], alignment=ft.MainAxisAlignment.SPACE_AROUND, spacing=0),
        bgcolor="#E8EBC9", border=ft.border.all(2, "#4A2E18"),
        border_radius=ft.border_radius.only(top_right=8, bottom_right=8, top_left=4, bottom_left=4),
        width=880, height=620,
        shadow=ft.BoxShadow(spread_radius=4, blur_radius=18, color=ft.Colors.BLACK54, offset=ft.Offset(0, 8))
    )
    page.add(agenda_marco)


if __name__ == "__main__":
    ft.app(target=main, view=ft.AppView.WEB_BROWSER)