import io
import re
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Auditor SIRE + CompraAPP + SSCO", layout="wide")
st.title("Extractor SIRE Compras (Formato CompraAPP), XML y Control SSCO")

# --- CONSTANTES ---
URLS_SSCO = {
    "SSCO_DIRECTO": "https://www.sunat.gob.pe/padronesnotificaciones/sujeSinCapacidadOperativa.html",
    "SSCO_VINCULADO_ART12": "https://www.sunat.gob.pe/padronesnotificaciones/relacion-ssco.html",
}

NS_UBL = {
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}

COLUMNAS_COMPRAAPP_61 = [
    "Vou.Origen", "Vou.Numero", "Vou.Fecha", "Doc", "Numero", "Fec.Doc", "Fec.Venc.",
    "Codigo", "B.I.O.G y E. (A)", "B.I.O.G.y E. y NO GRA. (B)", "B.I.O.G.sin D.C.FIS(C)",
    "AD. NO GRAV.", "I.S.C.", "IGV (A)", "IGV (B)", "IGV (C)", "OTROS TRIB.", "IMP. BOLSA",
    "Moneda", "TC", "Glosa", "Cta Gastos", "Cta IGV", "Cta O. Trib.", "Cta x Pagar",
    "C.Costo", "Presupuesto", "R.Doc", "R.numero", "R.Fecha", "D.Numero", "D.Fecha",
    "RUC", "R.Social", "Tipo", "Tip.Doc.Iden", "Medio de Pago", "Apellido 1", "Apellido 2",
    "Nombre", "T.Bien", "P.origen", "P.vou", "P.fecha", "P.fecha D.", "P.fecha V.",
    "P.cta cob", "P.m.pago", "P.doc", "P.num doc", "P.moneda", "P.tc", "P.monto",
    "P.glosa", "P.fe", "Retencion 0/1", "PDB ndes", "CodTasa", "Ind.Ret", "B.Imp", "IGV",
]

COLUMNAS_SSCO_EXTRA = ["Alerta SSCO", "Fecha Pub. SSCO"]
COLUMNAS_COMPRAAPP_TOTAL = COLUMNAS_COMPRAAPP_61 + COLUMNAS_SSCO_EXTRA

COLS_FECHA = {"Vou.Fecha", "Fec.Doc", "Fec.Venc.", "R.Fecha", "D.Fecha"}
COLS_MONTO = {
    "B.I.O.G y E. (A)", "B.I.O.G.y E. y NO GRA. (B)", "B.I.O.G.sin D.C.FIS(C)",
    "AD. NO GRAV.", "I.S.C.", "IGV (A)", "IGV (B)", "IGV (C)", "OTROS TRIB.", "IMP. BOLSA",
}
COLS_TC = {"TC"}

ANCHOS_BASE = {
    "Vou.Fecha": 12.5, "Doc": 9.5, "Numero": 16.0, "Fec.Doc": 12.5, "Fec.Venc.": 13.0,
    "Codigo": 14.0, "B.I.O.G y E. (A)": 15.5, "B.I.O.G.y E. y NO GRA. (B)": 24.5,
    "B.I.O.G.sin D.C.FIS(C)": 20.5, "AD. NO GRAV.": 14.5, "I.S.C.": 10.0,
    "IGV (A)": 13.5, "IGV (B)": 13.5, "IGV (C)": 13.5, "OTROS TRIB.": 13.5,
    "IMP. BOLSA": 12.5, "Moneda": 9.5, "TC": 11.0, "Glosa": 65.0,
    "R.Doc": 9.5, "R.numero": 15.0, "R.Fecha": 12.5, "RUC": 14.0,
    "R.Social": 45.0, "Tip.Doc.Iden": 13.0, "Alerta SSCO": 22.0, "Fecha Pub. SSCO": 16.5,
}


# --- UTILIDADES DE TEXTO Y PARSEO ---
def norm_txt(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def buscar_col(df: pd.DataFrame, patrones: list[str], excluir: list[str] | None = None) -> str | None:
    excluir_norm = [norm_txt(e) for e in (excluir or [])]
    norm_map = {col: norm_txt(col) for col in df.columns}

    for pat in patrones:
        p = norm_txt(pat)
        for col, ncol in norm_map.items():
            if any(ex in ncol for ex in excluir_norm):
                continue
            if p == ncol:
                return col

    for pat in patrones:
        p = norm_txt(pat)
        for col, ncol in norm_map.items():
            if any(ex in ncol for ex in excluir_norm):
                continue
            if p in ncol:
                return col
    return None


def parse_fecha(val):
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "01/01/0001"):
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt)
        except ValueError:
            continue
    return None


def parse_monto(val, tc=None, es_usd=False, es_nc=False):
    if val is None or pd.isna(val):
        return None
    s = str(val).strip().replace(",", "")
    if not s or s.lower() in ("none", "nan"):
        return None
    try:
        num = float(s)
    except ValueError:
        return None
    if abs(num) < 1e-9:
        return None
    if es_nc:
        num = abs(num)
    if es_usd and tc and tc > 0:
        num = round(num / tc, 2)
    else:
        num = round(num, 2)
    return num if abs(num) >= 1e-9 else None


def limpiar_num_doc(val) -> str:
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan"):
        return ""
    limpio = s.lstrip("0")
    return limpio if limpio else "0"


def limpiar_nombre_archivo(texto: str) -> str:
    limpio = re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()
    return re.sub(r"\s+", " ", limpio)


# --- 1. CLIENTE API SIRE ---
class SunatSireClient:
    def __init__(self, ruc, usuario_sol, clave_sol, client_id, client_secret):
        self.ruc = ruc.strip()
        self.usuario_completo = f"{self.ruc}{usuario_sol.strip()}"
        self.clave_sol = clave_sol.strip()
        self.client_id = client_id.strip()
        self.client_secret = client_secret.strip()
        self.token = None

    def autenticar(self):
        url = f"https://api-seguridad.sunat.gob.pe/v1/clientessol/{self.client_id}/oauth2/token/"
        payload = {
            "grant_type": "password",
            "scope": "https://api-sire.sunat.gob.pe",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "username": self.usuario_completo,
            "password": self.clave_sol,
        }
        r = requests.post(
            url,
            data=payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
        if not r.ok:
            raise RuntimeError(f"Error {r.status_code} al autenticar con SUNAT: {r.text}")
        self.token = r.json()["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}", "Accept": "application/json"}

    def descargar_propuesta_rce(self, periodo: str) -> pd.DataFrame:
        url_ticket = (
            f"https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rce/"
            f"propuesta/web/propuesta/{periodo}/exportacioncomprobantepropuesta"
            f"?codTipoArchivo=0&codOrigenEnvio=1"
        )
        r_tk = requests.get(url_ticket, headers=self._headers(), timeout=30)
        if not r_tk.ok:
            raise RuntimeError(f"Error {r_tk.status_code} al solicitar ticket: {r_tk.text}")
        num_ticket = r_tk.json()["numTicket"]

        url_estado = (
            f"https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rvierce/"
            f"gestionprocesosmasivos/web/masivo/consultaestadotickets"
            f"?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={num_ticket}"
        )
        nom_archivo, cod_tipo, cod_proceso = None, None, "10"
        for _ in range(20):
            r_est = requests.get(url_estado, headers=self._headers(), timeout=30)
            if not r_est.ok:
                raise RuntimeError(f"Error {r_est.status_code} consultando ticket: {r_est.text}")
            registros = r_est.json().get("registros", [])
            if registros and registros[0].get("codEstadoProceso") == "06":
                reg = registros[0]
                archivos = reg.get("archivoReporte") or []
                if not archivos:
                    raise ValueError("El ticket terminó sin archivos (verifica si hay compras en este periodo).")
                info = archivos[0]
                nom_archivo = info.get("nomArchivoReporte")
                cod_tipo = info.get("codTipoAchivoReporte") or info.get("codTipoArchivoReporte") or "01"
                cod_proceso = reg.get("codProceso") or "10"
                break
            time.sleep(3)

        if not nom_archivo:
            raise TimeoutError("SUNAT tardó demasiado en generar el archivo del ticket.")

        url_zip = (
            "https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rvierce/"
            "gestionprocesosmasivos/web/masivo/archivoreporte"
        )
        params_zip = {
            "nomArchivoReporte": nom_archivo,
            "codTipoArchivoReporte": cod_tipo,
            "codLibro": "080000",
            "perTributario": periodo,
            "codProceso": cod_proceso,
            "numTicket": num_ticket,
        }
        r_zip = requests.get(url_zip, headers=self._headers(), params=params_zip, timeout=60)

        if r_zip.status_code == 422 and cod_tipo == "00":
            params_zip["codTipoArchivoReporte"] = "01"
            r_zip = requests.get(url_zip, headers=self._headers(), params=params_zip, timeout=60)

        if not r_zip.ok:
            raise RuntimeError(f"Error {r_zip.status_code} de SUNAT al descargar ZIP: {r_zip.text}")

        with zipfile.ZipFile(io.BytesIO(r_zip.content)) as z:
            archivos_datos = [f for f in z.namelist() if f.lower().endswith((".txt", ".csv"))]
            if not archivos_datos:
                raise ValueError("El ZIP descargado de SUNAT no contiene archivos TXT o CSV.")
            raw_bytes = z.read(archivos_datos[0])

        try:
            df = pd.read_csv(io.BytesIO(raw_bytes), sep="|", dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(raw_bytes), sep="|", dtype=str, encoding="latin-1")
        except pd.errors.EmptyDataError:
            df = pd.DataFrame()

        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        return df


# --- 2. PADRÓN SSCO (CON CACHÉ DE 6 HORAS) ---
@st.cache_data(ttl=21600, show_spinner=False)
def obtener_padron_ssco() -> tuple[pd.DataFrame, list[str]]:
    registros = []
    errores = []
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    for tipo_lista, url in URLS_SSCO.items():
        try:
            resp = requests.get(url, headers=headers, timeout=25)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding or "utf-8"

            try:
                tablas = pd.read_html(io.StringIO(resp.text))
            except ValueError:
                continue

            for df in tablas:
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [" ".join(str(x) for x in col).strip() for col in df.columns]

                if all(isinstance(c, int) for c in df.columns) and len(df) > 1:
                    primera_fila = [str(x).strip() for x in df.iloc[0].tolist()]
                    if any("RUC" in x.upper() for x in primera_fila):
                        df = df.iloc[1:].copy()
                        df.columns = primera_fila

                col_ruc = next((c for c in df.columns if str(c).strip().upper() == "RUC"), None)
                if col_ruc is None:
                    col_ruc = next((c for c in df.columns if "RUC" in str(c).upper()), None)

                if col_ruc is not None:
                    col_fecha = next(
                        (c for c in df.columns if "PUBLICACI" in str(c).upper()),
                        None,
                    )
                    rucs = (
                        df[col_ruc]
                        .astype(str)
                        .str.replace(r"\.0$", "", regex=True)
                        .str.strip()
                    )
                    fechas = df[col_fecha].astype(str).str.strip() if col_fecha else ""

                    tmp = pd.DataFrame({
                        "ruc_ssco": rucs,
                        "alerta_ssco": tipo_lista,
                        "fecha_publicacion_ssco": fechas,
                    })
                    tmp = tmp[tmp["ruc_ssco"].str.fullmatch(r"\d{11}", na=False)]
                    if not tmp.empty:
                        registros.append(tmp)
        except Exception as e:
            errores.append(f"{tipo_lista}: {e}")

    if not registros:
        return pd.DataFrame(columns=["ruc_ssco", "alerta_ssco", "fecha_publicacion_ssco"]), errores

    padron = pd.concat(registros, ignore_index=True).drop_duplicates(subset=["ruc_ssco"])
    return padron, errores


# --- 3. PARSER DE XML UBL 2.1 ---
def leer_xml_bytes(contenido: bytes) -> dict | None:
    try:
        root = ET.fromstring(contenido)
        ruc_emisor = root.findtext(
            ".//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID",
            default="",
            namespaces=NS_UBL,
        ).strip()
        id_doc = root.findtext("./cbc:ID", default="", namespaces=NS_UBL).strip()

        serie, numero = "", ""
        if "-" in id_doc:
            partes = id_doc.split("-", 1)
            serie = partes[0].strip().upper()
            numero = limpiar_num_doc(partes[1])

        lineas = (
            root.findall(".//cac:InvoiceLine", NS_UBL)
            or root.findall(".//cac:CreditNoteLine", NS_UBL)
            or root.findall(".//cac:DebitNoteLine", NS_UBL)
        )
        descripciones = []
        for linea in lineas:
            desc = " ".join(
                d.text.strip()
                for d in linea.findall(".//cac:Item/cbc:Description", NS_UBL)
                if d.text
            )
            if desc:
                descripciones.append(desc)

        return {
            "clave_cruce": f"{ruc_emisor}|{serie}|{numero}",
            "ruc_xml": ruc_emisor,
            "comprobante_xml": id_doc,
            "descripcion_items": " ; ".join(descripciones),
        }
    except Exception:
        return None


def procesar_archivos_xml(archivos_subidos) -> pd.DataFrame:
    filas = []
    for archivo in archivos_subidos:
        nombre = archivo.name.lower()
        if nombre.endswith(".xml"):
            data = leer_xml_bytes(archivo.read())
            if data:
                filas.append(data)
        elif nombre.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(archivo.read())) as z:
                for item in z.namelist():
                    if item.lower().endswith(".xml") and not item.startswith("__MACOSX"):
                        data = leer_xml_bytes(z.read(item))
                        if data:
                            filas.append(data)
    if not filas:
        return pd.DataFrame(columns=["clave_cruce", "ruc_xml", "comprobante_xml", "descripcion_items"])
    return pd.DataFrame(filas).drop_duplicates(subset=["clave_cruce"])


# --- 4. TRANSFORMADOR A FORMATO COMPRAAPP + FILTRO INTELIGENTE ---
def construir_compraapp(df_sire: pd.DataFrame) -> pd.DataFrame:
    c_fec_em = buscar_col(df_sire, ["fecha de emision", "fecha emision"], excluir=["modificado"])
    c_fec_vc = buscar_col(df_sire, ["fecha vcto/pago", "fecha vcto", "vencimiento"])
    c_tipo_cp = buscar_col(df_sire, ["tipo cp/doc.", "tipo cp/doc"], excluir=["modificado"])
    c_serie = buscar_col(df_sire, ["serie del cdp", "serie"], excluir=["modificado"])
    c_nro = buscar_col(
        df_sire,
        ["nro cp o doc. nro inicial (rango)", "nro cp o doc", "nro inicial"],
        excluir=["modificado", "final"],
    )
    c_tip_id = buscar_col(df_sire, ["tipo doc identidad", "tipo doc. identidad"])
    c_ruc_pr = buscar_col(df_sire, ["nro doc identidad", "nro doc. identidad"], excluir=["tipo"])
    c_rsoc_pr = buscar_col(
        df_sire,
        ["apellidos nombres/ razon social", "apellidos nombres/razon social", "razon social"],
        excluir=["o razon social"],
    )
    if not c_rsoc_pr:
        cands = [c for c in df_sire.columns if "razon" in norm_txt(c)]
        c_rsoc_pr = cands[-1] if len(cands) > 1 else (cands[0] if cands else None)

    c_bi_a = buscar_col(df_sire, ["bi gravado dg"])
    c_igv_a = buscar_col(df_sire, ["igv / ipm dg", "igv/ipm dg"])
    c_bi_b = buscar_col(df_sire, ["bi gravado dgng"])
    c_igv_b = buscar_col(df_sire, ["igv / ipm dgng", "igv/ipm dgng"])
    c_bi_c = buscar_col(df_sire, ["bi gravado dng"])
    c_igv_c = buscar_col(df_sire, ["igv / ipm dng", "igv/ipm dng"])
    c_ng = buscar_col(df_sire, ["valor adq. ng", "valor adq ng"])
    c_isc = buscar_col(df_sire, ["isc"])
    c_icbper = buscar_col(df_sire, ["icbper"])
    c_otros = buscar_col(df_sire, ["otros trib/ cargos", "otros trib"])
    c_mon = buscar_col(df_sire, ["moneda"])
    c_tc = buscar_col(df_sire, ["tipo de cambio"])
    c_r_fec = buscar_col(df_sire, ["fecha emision doc modificado"])
    c_r_doc = buscar_col(df_sire, ["tipo cp modificado"])
    c_r_ser = buscar_col(df_sire, ["serie cp modificado"])
    c_r_nro = buscar_col(df_sire, ["nro cp modificado"])

    filas = []
    for _, row in df_sire.iterrows():
        def g(col):
            return row[col] if (col and col in df_sire.columns) else None

        fec_doc = parse_fecha(g(c_fec_em))
        fec_venc = parse_fecha(g(c_fec_vc))

        doc_raw = str(g(c_tipo_cp) or "").strip()
        doc = doc_raw.zfill(2) if doc_raw.isdigit() else doc_raw
        es_nc = (doc == "07")

        serie = str(g(c_serie) or "").strip().upper()
        if serie.lower() in ("none", "nan"):
            serie = ""
        nro = limpiar_num_doc(g(c_nro))
        numero = f"{serie}-{nro}" if (serie and nro) else (serie or nro)

        ruc_prov = str(g(c_ruc_pr) or "").strip()
        if ruc_prov.lower() in ("none", "nan"):
            ruc_prov = ""
        rsoc_prov = str(g(c_rsoc_pr) or "").strip()
        if rsoc_prov.lower() in ("none", "nan"):
            rsoc_prov = ""
        tip_id = str(g(c_tip_id) or "").strip()
        if tip_id.lower() in ("none", "nan"):
            tip_id = ""

        mon_raw = str(g(c_mon) or "PEN").strip().upper()
        es_usd = mon_raw in ("USD", "D", "US$")
        moneda = "D" if es_usd else "S"

        tc_val = None
        if es_usd:
            try:
                t = float(str(g(c_tc) or "").strip().replace(",", ""))
                if t > 0:
                    tc_val = round(t, 3)
            except ValueError:
                tc_val = None

        r_doc_raw = str(g(c_r_doc) or "").strip()
        r_doc = (
            r_doc_raw.zfill(2)
            if r_doc_raw.isdigit()
            else ("" if r_doc_raw.lower() in ("none", "nan") else r_doc_raw)
        )
        r_ser = str(g(c_r_ser) or "").strip().upper()
        if r_ser.lower() in ("none", "nan"):
            r_ser = ""
        r_nro = limpiar_num_doc(g(c_r_nro))
        r_numero = f"{r_ser}-{r_nro}" if (r_ser and r_nro) else (r_ser or r_nro)
        r_fecha = parse_fecha(g(c_r_fec))

        glosa = (
            str(row["descripcion_items"]).strip()
            if ("descripcion_items" in df_sire.columns and pd.notna(row["descripcion_items"]))
            else ""
        )
        alerta_ssco = (
            str(row["alerta_ssco"]).strip()
            if ("alerta_ssco" in df_sire.columns and pd.notna(row["alerta_ssco"]))
            else ""
        )
        fecha_ssco = (
            str(row["fecha_publicacion_ssco"]).strip()
            if ("fecha_publicacion_ssco" in df_sire.columns and pd.notna(row["fecha_publicacion_ssco"]))
            else ""
        )

        registro = {col: None for col in COLUMNAS_COMPRAAPP_TOTAL}
        registro.update({
            "Vou.Fecha": fec_doc,
            "Doc": doc or None,
            "Numero": numero or None,
            "Fec.Doc": fec_doc,
            "Fec.Venc.": fec_venc,
            "Codigo": ruc_prov or None,
            "B.I.O.G y E. (A)": parse_monto(g(c_bi_a), tc_val, es_usd, es_nc),
            "B.I.O.G.y E. y NO GRA. (B)": parse_monto(g(c_bi_b), tc_val, es_usd, es_nc),
            "B.I.O.G.sin D.C.FIS(C)": parse_monto(g(c_bi_c), tc_val, es_usd, es_nc),
            "AD. NO GRAV.": parse_monto(g(c_ng), tc_val, es_usd, es_nc),
            "I.S.C.": parse_monto(g(c_isc), tc_val, es_usd, es_nc),
            "IGV (A)": parse_monto(g(c_igv_a), tc_val, es_usd, es_nc),
            "IGV (B)": parse_monto(g(c_igv_b), tc_val, es_usd, es_nc),
            "IGV (C)": parse_monto(g(c_igv_c), tc_val, es_usd, es_nc),
            "OTROS TRIB.": parse_monto(g(c_otros), tc_val, es_usd, es_nc),
            "IMP. BOLSA": parse_monto(g(c_icbper), tc_val, es_usd, es_nc),
            "Moneda": moneda,
            "TC": tc_val,
            "Glosa": glosa or None,
            "R.Doc": r_doc or None,
            "R.numero": r_numero or None,
            "R.Fecha": r_fecha,
            "RUC": ruc_prov or None,
            "R.Social": rsoc_prov or None,
            "Tip.Doc.Iden": tip_id or None,
            "Alerta SSCO": alerta_ssco or None,
            "Fecha Pub. SSCO": fecha_ssco or None,
        })
        filas.append(registro)

    df_out = pd.DataFrame(filas, columns=COLUMNAS_COMPRAAPP_TOTAL)
    df_out = df_out.sort_values(
        by=["Doc", "Fec.Doc"],
        ascending=[True, True],
        na_position="last",
    ).reset_index(drop=True)
    return df_out


def filtrar_columnas_con_datos(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    cols_con_datos = []
    cols_vacias = []
    for col in df.columns:
        tiene_dato = False
        for v in df[col]:
            if v is None or pd.isna(v):
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            if col in COLS_MONTO and isinstance(v, (int, float)) and abs(v) < 1e-9:
                continue
            tiene_dato = True
            break
        if tiene_dato:
            cols_con_datos.append(col)
        else:
            cols_vacias.append(col)
    return df[cols_con_datos].copy(), cols_vacias


# --- 5. GENERADOR EXCEL ESTILO COMPRAAPP ---
def escribir_hoja_compraapp(ws, df_hoja: pd.DataFrame):
    font_hdr = Font(name="Calibri", size=11, bold=True)
    font_data = Font(name="Calibri", size=11, bold=False)
    font_alerta = Font(name="Calibri", size=11, bold=True, color="9C0006")
    align_hdr = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="000000")
    border_hdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_alerta = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    cols = list(df_hoja.columns)

    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font = font_hdr
        cell.alignment = align_hdr
        cell.border = border_hdr

    for r_idx, (_, row) in enumerate(df_hoja.iterrows(), start=2):
        tiene_ssco = (
            "Alerta SSCO" in df_hoja.columns
            and pd.notna(row.get("Alerta SSCO"))
            and str(row.get("Alerta SSCO")).strip() != ""
        )
        for c_idx, col_name in enumerate(cols, start=1):
            val = row[col_name]
            if val is None or pd.isna(val):
                clean_val = None
            elif isinstance(val, pd.Timestamp):
                clean_val = val.to_pydatetime()
            else:
                clean_val = val

            cell = ws.cell(row=r_idx, column=c_idx, value=clean_val)
            cell.font = font_alerta if (tiene_ssco and col_name in ("Alerta SSCO", "RUC", "R.Social")) else font_data
            if tiene_ssco and col_name in ("Alerta SSCO", "Fecha Pub. SSCO"):
                cell.fill = fill_alerta

            if col_name in COLS_FECHA:
                cell.number_format = "dd/mm/yyyy"
            elif col_name in COLS_MONTO:
                cell.number_format = "#,###,##0.00"
            elif col_name in COLS_TC:
                cell.number_format = "#,###,##0.000"
            else:
                cell.number_format = "@"

    for c_idx, col_name in enumerate(cols, start=1):
        col_letter = get_column_letter(c_idx)
        base_w = ANCHOS_BASE.get(col_name, max(len(col_name) + 3, 12.0))
        max_len = len(col_name)
        for r_idx in range(2, len(df_hoja) + 2):
            v = ws.cell(row=r_idx, column=c_idx).value
            if v is not None:
                s = v.strftime("%d/%m/%Y") if isinstance(v, datetime) else str(v)
                if len(s) > max_len:
                    max_len = len(s)
        cap = 100.0 if col_name == "Glosa" else (55.0 if col_name == "R.Social" else 30.0)
        ws.column_dimensions[col_letter].width = min(max(base_w, max_len + 2.5), cap)


def generar_excel_compraapp(df_formato: pd.DataFrame, omitir_vacias: bool = True) -> tuple[bytes, list[str], list[str]]:
    if omitir_vacias:
        df_export, cols_omitidas = filtrar_columnas_con_datos(df_formato)
    else:
        hay_ssco = "Alerta SSCO" in df_formato.columns and df_formato["Alerta SSCO"].notna().any()
        cols_fijas = COLUMNAS_COMPRAAPP_TOTAL if hay_ssco else COLUMNAS_COMPRAAPP_61
        df_export = df_formato[cols_fijas].copy()
        cols_omitidas = []

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hoja 1"
    escribir_hoja_compraapp(ws, df_export)

    # Crear segunda pestaña solo si realmente se detectaron proveedores SSCO
    if "Alerta SSCO" in df_formato.columns:
        df_ssco = df_formato[df_formato["Alerta SSCO"].notna()]
        if not df_ssco.empty:
            df_ssco_clean, _ = filtrar_columnas_con_datos(df_ssco)
            ws_ssco = wb.create_sheet(title="Alertas_SSCO")
            escribir_hoja_compraapp(ws_ssco, df_ssco_clean)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), list(df_export.columns), cols_omitidas


# --- INTERFAZ LATERAL ---
with st.sidebar:
    st.header("Credenciales SUNAT")
    ruc = st.text_input("RUC", max_chars=11)
    usuario_sol = st.text_input("Usuario SOL")
    clave_sol = st.text_input("Clave SOL", type="password")
    client_id = st.text_input("API Client ID (SIRE)")
    client_secret = st.text_input("API Client Secret (SIRE)", type="password")
    periodo = st.text_input("Periodo (AAAAMM)", value="202611", max_chars=6)
    btn_descargar = st.button("1. Descargar Propuesta SIRE", type="primary", use_container_width=True)

if "df_sire" not in st.session_state:
    st.session_state.df_sire = None
if "total_padron_ssco" not in st.session_state:
    st.session_state.total_padron_ssco = 0
if "razon_contribuyente" not in st.session_state:
    st.session_state.razon_contribuyente = ""

# --- FLUJO PRINCIPAL ---
if btn_descargar:
    if not all([ruc, usuario_sol, clave_sol, client_id, client_secret, periodo]):
        st.error("Completa todas las credenciales y el periodo antes de consultar.")
    else:
        with st.spinner("Consultando API SIRE de SUNAT y cruzando padrón SSCO..."):
            try:
                cliente = SunatSireClient(ruc, usuario_sol, clave_sol, client_id, client_secret)
                cliente.autenticar()
                df_sire = cliente.descargar_propuesta_rce(periodo)

                if df_sire.empty:
                    st.warning("La propuesta del periodo consultado está vacía (0 comprobantes).")
                    st.session_state.df_sire = None
                else:
                    col_razon_contrib = buscar_col(df_sire, ["apellidos y nombres o razon social"])
                    if col_razon_contrib and not df_sire[col_razon_contrib].dropna().empty:
                        st.session_state.razon_contribuyente = str(df_sire[col_razon_contrib].dropna().iloc[0]).strip()
                    else:
                        st.session_state.razon_contribuyente = ruc.strip()

                    col_ruc_prov = buscar_col(df_sire, ["nro doc identidad", "nro doc. identidad"], excluir=["tipo"])
                    col_serie = buscar_col(df_sire, ["serie del cdp", "serie"], excluir=["modificado"])
                    col_num = buscar_col(
                        df_sire,
                        ["nro cp o doc. nro inicial (rango)", "nro cp o doc", "nro inicial"],
                        excluir=["modificado", "final"],
                    )

                    padron_ssco, errores_ssco = obtener_padron_ssco()
                    st.session_state.total_padron_ssco = len(padron_ssco)
                    if errores_ssco:
                        st.warning("Aviso al consultar listas SSCO: " + " | ".join(errores_ssco))

                    if col_ruc_prov:
                        df_sire[col_ruc_prov] = df_sire[col_ruc_prov].astype(str).str.strip()
                        df_sire = df_sire.merge(
                            padron_ssco,
                            left_on=col_ruc_prov,
                            right_on="ruc_ssco",
                            how="left",
                        ).drop(columns=["ruc_ssco"], errors="ignore")

                        if col_serie and col_num:
                            df_sire["clave_cruce"] = (
                                df_sire[col_ruc_prov].str.strip()
                                + "|"
                                + df_sire[col_serie].astype(str).str.strip().str.upper()
                                + "|"
                                + df_sire[col_num].apply(limpiar_num_doc)
                            )

                    st.session_state.df_sire = df_sire
                    st.success(f"Propuesta cargada: {len(df_sire)} comprobantes encontrados.")
            except Exception as e:
                st.error(f"Error en la consulta: {e}")

# --- VISUALIZACIÓN, CRUCE CON XML Y EXPORTACIÓN COMPRAAPP ---
if st.session_state.df_sire is not None:
    df_trabajo = st.session_state.df_sire.copy()

    total_docs = len(df_trabajo)
    total_ssco = int(df_trabajo["alerta_ssco"].notna().sum()) if "alerta_ssco" in df_trabajo.columns else 0
    c1, c2, c3 = st.columns(3)
    c1.metric("Comprobantes en Propuesta RCE", total_docs)
    c2.metric("RUCs verificados en Padrón SSCO", st.session_state.total_padron_ssco)
    c3.metric("Proveedores con Alerta SSCO", total_ssco)

    if total_ssco > 0:
        st.error(f"Se detectaron {total_ssco} comprobante(s) de proveedores en la lista SSCO de SUNAT.")

    st.divider()
    st.subheader("2. Agregar descripciones (Glosa) desde archivos XML")
    archivos_xml = st.file_uploader(
        "Sube uno o varios archivos .xml o un .zip con los comprobantes del periodo",
        type=["xml", "zip"],
        accept_multiple_files=True,
    )

    if archivos_xml:
        df_xml = procesar_archivos_xml(archivos_xml)
        if not df_xml.empty and "clave_cruce" in df_trabajo.columns:
            df_trabajo = df_trabajo.merge(
                df_xml[["clave_cruce", "descripcion_items"]],
                on="clave_cruce",
                how="left",
            )
            coincidencias = int(df_trabajo["descripcion_items"].notna().sum())
            st.info(f"Se vincularon descripciones (Glosa) en {coincidencias} de {total_docs} comprobantes.")

    # Transformar al formato CompraAPP
    df_compraapp = construir_compraapp(df_trabajo)

    st.divider()
    st.subheader("3. Vista previa formato CompraAPP")
    omitir_vacias = st.checkbox(
        "Omitir automáticamente columnas vacías en el Excel",
        value=True,
        help="Analiza cada columna y elimina automáticamente las que no tengan datos o estén en cero.",
    )

    excel_bytes, cols_incluidas, cols_omitidas = generar_excel_compraapp(
        df_compraapp,
        omitir_vacias=omitir_vacias,
    )

    st.caption(
        f"Columnas activas en el Excel ({len(cols_incluidas)}): {', '.join(cols_incluidas)}"
        + (f" — Omitidas por estar vacías: {len(cols_omitidas)}" if omitir_vacias else "")
    )

    # Preparar vista previa con fechas legibles dd/mm/yyyy
    df_preview = df_compraapp[cols_incluidas].copy()
    for col_f in COLS_FECHA.intersection(df_preview.columns):
        df_preview[col_f] = df_preview[col_f].apply(
            lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
        )

    st.dataframe(df_preview, use_container_width=True)

    periodo_guion = f"{periodo[:4]}-{periodo[4:6]}" if len(periodo) == 6 else periodo
    razon_limpia = limpiar_nombre_archivo(st.session_state.razon_contribuyente or ruc)
    nombre_xlsx = f"CompraAPP {periodo_guion} {razon_limpia}.xlsx"

    st.download_button(
        label=f"Descargar {nombre_xlsx}",
        data=excel_bytes,
        file_name=nombre_xlsx,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
