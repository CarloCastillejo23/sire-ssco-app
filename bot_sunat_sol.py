import json
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

NS_UBL = {
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}

URL_SOL_MENU = "https://e-menu.sunat.gob.pe/cl-ti-itmenu/MenuInternet.htm"


def log(msg: str):
    print(msg, flush=True)


def extraer_descripcion_xml(ruta_xml: Path) -> str:
    try:
        raw = ruta_xml.read_bytes()
        # Algunos XML bajados de SUNAT vienen dentro de un ZIP
        if raw[:2] == b"PK":
            with zipfile.ZipFile(ruta_xml, "r") as z:
                xml_inner = [
                    f for f in z.namelist()
                    if f.lower().endswith(".xml") and not f.startswith("__MACOSX")
                ]
                if not xml_inner:
                    return ""
                raw = z.read(xml_inner[0])
                ruta_xml.write_bytes(raw)

        root = ET.fromstring(raw)
        lineas = (
            root.findall(".//cac:InvoiceLine", NS_UBL)
            or root.findall(".//cac:CreditNoteLine", NS_UBL)
            or root.findall(".//cac:DebitNoteLine", NS_UBL)
        )
        descripciones = []
        for linea in lineas:
            partes = [
                d.text.strip()
                for d in linea.findall(".//cac:Item/cbc:Description", NS_UBL)
                if d.text and d.text.strip()
            ]
            if partes:
                texto = " ".join(partes)
                if texto not in descripciones:
                    descripciones.append(texto)

        if not descripciones:
            notas = [
                n.text.strip()
                for n in root.findall("./cbc:Note", NS_UBL)
                if n.text and n.text.strip()
            ]
            return " ; ".join(notas[:2])

        return " ; ".join(descripciones)
    except Exception as e:
        log(f"[AVISO] No se pudo leer XML {ruta_xml.name}: {e}")
        return ""


def cerrar_modales_sol(page):
    selectores = [
        'button:has-text("Continuar sin confirmar")',
        'button:has-text("Finalizar")',
        'button:has-text("Cerrar")',
        'button:has-text("Aceptar")',
        "#btnCerrar",
        ".modal.in button.close",
        ".modal.show button.close",
    ]
    for _ in range(3):
        cerro_algo = False
        for sel in selectores:
            try:
                loc = page.locator(sel).first
                if loc.is_visible(timeout=600):
                    loc.click(timeout=1500)
                    cerro_algo = True
                    time.sleep(0.5)
            except Exception:
                pass
        if not cerro_algo:
            break


def iniciar_sesion_sol(page, ruc: str, usuario: str, clave: str):
    log("Abriendo Portal SUNAT SOL...")
    page.goto(URL_SOL_MENU, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("#txtRuc", timeout=30000)

    page.fill("#txtRuc", ruc)
    page.fill("#txtUsuario", usuario)
    page.fill("#txtContrasena", clave)
    page.click("#btnAceptar")

    page.wait_for_selector("#txtBusca, #divOpcionServicio2, .menu-principal", timeout=45000)
    time.sleep(2)
    cerrar_modales_sol(page)
    log("Sesión iniciada correctamente en SUNAT SOL.")


def navegar_a_consulta_cpe(page):
    log("Navegando a Consulta de Comprobantes de Pago...")

    try:
        tab_empresas = page.locator(
            '#divOpcionServicio2, a:has-text("Empresas"), button:has-text("Empresas")'
        ).first
        if tab_empresas.is_visible(timeout=2000):
            tab_empresas.click()
            time.sleep(1)
    except Exception:
        pass

    rutas_busqueda = [
        ("Nueva Consulta de comprobantes de pago", "Nueva Consulta de comprobantes de pago"),
        ("Consulta de Comprobantes de Pago", "Consulta de Comprobantes de Pago"),
        ("Consulta de Facturas y Notas", "Consulta de Facturas y Notas"),
    ]

    navegado = False
    for texto_buscar, texto_click in rutas_busqueda:
        try:
            buscador = page.locator("#txtBusca").first
            if buscador.is_visible(timeout=2500):
                buscador.fill("")
                buscador.fill(texto_buscar)
                time.sleep(1.2)
                opcion = page.locator(
                    f'a:has-text("{texto_click}"), span:has-text("{texto_click}"), li:has-text("{texto_click}")'
                ).last
                if opcion.is_visible(timeout=2500):
                    opcion.click()
                    navegado = True
                    log(f"Opción abierta: {texto_click}")
                    break
        except Exception:
            continue

    if not navegado:
        for texto in ["Comprobantes de Pago", "SEE - SOL", "Factura Electrónica", "Consultar Factura y Nota"]:
            try:
                loc = page.locator(f'text="{texto}"').first
                if loc.is_visible(timeout=2000):
                    loc.click()
                    time.sleep(0.8)
            except Exception:
                pass

    page.wait_for_selector('iframe[name="iframeApplication"], #iframeApplication', timeout=30000)
    time.sleep(3)

    frame = None
    for f in page.frames:
        if f.name == "iframeApplication" or "iframeApplication" in (f.name or ""):
            frame = f
            break
    if frame is None:
        iframe_el = page.locator('iframe[name="iframeApplication"], #iframeApplication').first
        frame = iframe_el.content_frame()

    if frame is None:
        raise RuntimeError("No se pudo acceder al iframe de Consulta de Comprobantes de Pago.")
    return frame


def guardar_diagnostico_iframe(frame, carpeta_base: Path):
    try:
        info = frame.evaluate("""() => {
            const els = Array.from(document.querySelectorAll('input, select, button, a')).map(el => ({
                tag: el.tagName,
                type: el.type || '',
                id: el.id || '',
                name: el.name || '',
                placeholder: el.placeholder || '',
                text: (el.innerText || el.value || '').trim().slice(0, 60)
            }));
            return { url: location.href, count: els.length, elements: els };
        }""")
        ruta = carpeta_base / "diagnostico_sunat_iframe.json"
        ruta.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"[INFO] Diagnóstico del formulario guardado en: {ruta.name}")
    except Exception:
        pass


def buscar_y_descargar_por_comprobante(
    page, frame, item: dict, dir_xml: Path, dir_pdf: Path
) -> tuple[Path | None, Path | None, str]:
    ruc_emisor = item["ruc_emisor"]
    tipo_doc = item["tipo_doc"]
    serie = item["serie"]
    numero = item["numero"]
    fecha_emision = item.get("fecha_emision", "")

    nombre_base = f"{ruc_emisor}-{tipo_doc}-{serie}-{numero}"
    ruta_xml = dir_xml / f"{nombre_base}.xml"
    ruta_pdf = dir_pdf / f"{nombre_base}.pdf"

    if ruta_xml.exists() and ruta_xml.stat().st_size > 50 and ruta_pdf.exists() and ruta_pdf.stat().st_size > 100:
        return ruta_xml, ruta_pdf, "YA_DESCARGADO"

    # 1) Marcar opción "Recibido" si está visible
    for sel_rec in [
        'input[type="radio"][value*="R"]',
        'label:has-text("Recibido")',
        'label:has-text("Recibidos")',
    ]:
        try:
            loc = frame.locator(sel_rec).first
            if loc.is_visible(timeout=400):
                loc.click(timeout=1000)
                break
        except Exception:
            pass

    # Seleccionar tipo de comprobante en combo desplegable
    mapa_tipo_consulta = {
        "01": ["FE Recibidas", "Factura", "01"],
        "03": ["BV Recibidas", "Boleta", "03"],
        "07": ["NC Recibidas", "Nota de Crédito", "Nota de Credito", "07"],
        "08": ["ND Recibidas", "Nota de Débito", "Nota de Debito", "08"],
    }
    try:
        for sel in frame.locator("select").all():
            if not sel.is_visible(timeout=300):
                continue
            opciones_u = [o.upper() for o in sel.locator("option").all_inner_texts()]
            if any("FACTURA" in o or "FE RECIBIDA" in o or "COMPROBANTE" in o for o in opciones_u):
                for candidato in mapa_tipo_consulta.get(tipo_doc, []):
                    idx = next((i for i, o in enumerate(opciones_u) if candidato.upper() in o), None)
                    if idx is not None:
                        sel.select_option(index=idx)
                        time.sleep(0.3)
                        break
    except Exception:
        pass

    # 2) Llenar campos de búsqueda
    def llenar_si_existe(selectores: list[str], valor: str) -> bool:
        if not valor:
            return False
        for s in selectores:
            try:
                loc = frame.locator(s).first
                if loc.is_visible(timeout=400) and loc.is_enabled(timeout=400):
                    loc.fill("")
                    loc.fill(valor)
                    return True
            except Exception:
                continue
        return False

    llenar_si_existe(
        ['input[id*="ruc" i]', 'input[name*="ruc" i]', 'input[placeholder*="RUC" i]'],
        ruc_emisor,
    )
    llenar_si_existe(
        ['input[id*="serie" i]', 'input[name*="serie" i]', 'input[placeholder*="Serie" i]'],
        serie,
    )
    llenar_si_existe(
        [
            'input[id*="numero" i]', 'input[id*="numDoc" i]', 'input[id*="nro" i]',
            'input[name*="numero" i]', 'input[name*="nro" i]',
            'input[placeholder*="Número" i]', 'input[placeholder*="Numero" i]',
        ],
        numero,
    )
    if fecha_emision:
        llenar_si_existe(
            [
                'input[id*="fecIni" i]', 'input[id*="fechaInicio" i]', 'input[id*="desde" i]',
                'input[name*="fecIni" i]', 'input[name*="desde" i]',
            ],
            fecha_emision,
        )
        llenar_si_existe(
            [
                'input[id*="fecFin" i]', 'input[id*="fechaFin" i]', 'input[id*="hasta" i]',
                'input[name*="fecFin" i]', 'input[name*="hasta" i]',
            ],
            fecha_emision,
        )

    # 3) Clic en Buscar / Consultar
    clic_buscar = False
    for sel_btn in [
        'button:has-text("Buscar")', 'input[value*="Buscar" i]',
        'button:has-text("Consultar")', 'input[value*="Consultar" i]',
        'button:has-text("Aceptar")', "#btnBuscar",
    ]:
        try:
            btn = frame.locator(sel_btn).first
            if btn.is_visible(timeout=500):
                btn.click(timeout=2500)
                clic_buscar = True
                break
        except Exception:
            continue

    if not clic_buscar:
        return None, None, "NO_SE_ENCONTRO_BOTON_BUSCAR"

    time.sleep(2.2)

    # 4) Descargar XML y PDF
    xml_ok = ruta_xml.exists() and ruta_xml.stat().st_size > 50
    pdf_ok = ruta_pdf.exists() and ruta_pdf.stat().st_size > 100

    def intentar_descarga(selectores: list[str], destino: Path) -> bool:
        for sel in selectores:
            try:
                loc = frame.locator(sel).first
                if loc.is_visible(timeout=800):
                    with page.expect_download(timeout=15000) as dl_info:
                        loc.click(timeout=3000)
                    dl = dl_info.value
                    dl.save_as(str(destino))
                    if destino.exists() and destino.stat().st_size > 50:
                        return True
            except PlaywrightTimeout:
                continue
            except Exception:
                continue
        return False

    if not xml_ok:
        xml_ok = intentar_descarga(
            [
                'a:has-text("XML")', 'button:has-text("XML")',
                'a[title*="XML" i]', 'button[title*="XML" i]',
                'img[alt*="XML" i]', 'img[title*="XML" i]',
                '[id*="xml" i]', '[onclick*="xml" i]',
            ],
            ruta_xml,
        )

    if not pdf_ok:
        pdf_ok = intentar_descarga(
            [
                'a:has-text("PDF")', 'button:has-text("PDF")',
                'a[title*="PDF" i]', 'button[title*="PDF" i]',
                'img[alt*="PDF" i]', 'img[title*="PDF" i]',
                '[id*="pdf" i]', '[onclick*="pdf" i]',
            ],
            ruta_pdf,
        )

    estado = ("XML:OK" if xml_ok else "XML:NO_DISP") + " | " + ("PDF:OK" if pdf_ok else "PDF:NO_DISP")
    return (ruta_xml if xml_ok else None), (ruta_pdf if pdf_ok else None), estado


def ejecutar_bot(ruta_tarea_json: str):
    tarea = json.loads(Path(ruta_tarea_json).read_text(encoding="utf-8"))
    ruc = tarea["ruc"]
    usuario = tarea["usuario_sol"]
    clave = tarea["clave_sol"]
    comprobantes = tarea["comprobantes"]
    carpeta_salida = Path(tarea["carpeta_salida"])
    ruta_resultado = Path(tarea["ruta_resultado"])
    headless = bool(tarea.get("headless", False))

    dir_xml = carpeta_salida / "xml"
    dir_pdf = carpeta_salida / "pdf"
    dir_xml.mkdir(parents=True, exist_ok=True)
    dir_pdf.mkdir(parents=True, exist_ok=True)

    resultados = {}
    total = len(comprobantes)
    log(f"Iniciando bot para {total} comprobante(s). Carpeta: {carpeta_salida}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=120)
        context = browser.new_context(accept_downloads=True, locale="es-PE")
        page = context.new_page()

        try:
            iniciar_sesion_sol(page, ruc, usuario, clave)
            frame = navegar_a_consulta_cpe(page)
            guardar_diagnostico_iframe(frame, carpeta_salida)

            for idx, item in enumerate(comprobantes, start=1):
                clave_cruce = item["clave_cruce"]
                etiqueta = f"{item['ruc_emisor']} {item['serie']}-{item['numero']}"
                log(f"[{idx}/{total}] Procesando {etiqueta}...")

                try:
                    ruta_xml, ruta_pdf, estado = buscar_y_descargar_por_comprobante(
                        page, frame, item, dir_xml, dir_pdf
                    )
                    glosa = extraer_descripcion_xml(ruta_xml) if ruta_xml else ""
                    resultados[clave_cruce] = {
                        "glosa": glosa,
                        "xml_path": str(ruta_xml) if ruta_xml else "",
                        "pdf_path": str(ruta_pdf) if ruta_pdf else "",
                        "estado": estado,
                    }
                    log(f"    -> {estado}" + (f" | Glosa: {glosa[:65]}" if glosa else ""))
                except Exception as e:
                    resultados[clave_cruce] = {
                        "glosa": "",
                        "xml_path": "",
                        "pdf_path": "",
                        "estado": f"ERROR: {e}",
                    }
                    log(f"    -> ERROR en {etiqueta}: {e}")

                ruta_resultado.write_text(
                    json.dumps(resultados, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                time.sleep(1.2)
        finally:
            browser.close()

    log("Proceso de descarga SOL finalizado.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python bot_sunat_sol.py <ruta_tarea.json>")
        sys.exit(1)
    ejecutar_bot(sys.argv[1])
