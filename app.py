import io
import time
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Auditor SIRE + SSCO", layout="wide")
st.title("Extractor SIRE Compras, Descripciones XML y Control SSCO")

# --- CONSTANTES ---
URLS_SSCO = {
    "SSCO_DIRECTO": "https://www.sunat.gob.pe/padronesnotificaciones/sujeSinCapacidadOperativa.html",
    "SSCO_VINCULADO_ART12": "https://www.sunat.gob.pe/padronesnotificaciones/relacion-ssco.html",
}

NS_UBL = {
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}


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
        r.raise_for_status()
        self.token = r.json()["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}", "Accept": "application/json"}

    def descargar_propuesta_rce(self, periodo: str) -> pd.DataFrame:
        # 1) Solicitar ticket (codTipoArchivo=0 devuelve TXT delimitado por '|')
        url_ticket = (
            f"https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rce/"
            f"propuesta/web/propuesta/{periodo}/exportacioncomprobantepropuesta"
            f"?codTipoArchivo=0&codOrigenEnvio=1"
        )
        r_tk = requests.get(url_ticket, headers=self._headers(), timeout=30)
        r_tk.raise_for_status()
        num_ticket = r_tk.json()["numTicket"]

        # 2) Consultar estado del ticket
        url_estado = (
            f"https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rvierce/"
            f"gestionprocesosmasivos/web/masivo/consultaestadotickets"
            f"?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={num_ticket}"
        )
        nom_archivo, cod_tipo = None, None
        for _ in range(20):
            r_est = requests.get(url_estado, headers=self._headers(), timeout=30).json()
            registros = r_est.get("registros", [])
            if registros and registros[0].get("codEstadoProceso") == "06":
                info = registros[0]["archivoReporte"][0]
                nom_archivo = info["nomArchivoReporte"]
                cod_tipo = info["codTipoAchivoReporte"]
                break
            time.sleep(3)

        if not nom_archivo:
            raise TimeoutError("SUNAT tardó demasiado en generar el archivo del ticket.")

        # 3) Descargar ZIP en memoria
        url_zip = (
            f"https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros/rvierce/"
            f"gestionprocesosmasivos/web/masivo/archivoreporte"
            f"?nomArchivoReporte={nom_archivo}&codTipoArchivoReporte={cod_tipo}&codLibro=080000"
        )
        r_zip = requests.get(url_zip, headers=self._headers(), timeout=60)
        r_zip.raise_for_status()

        with zipfile.ZipFile(io.BytesIO(r_zip.content)) as z:
            archivo_txt = [f for f in z.namelist() if f.lower().endswith((".txt", ".csv"))][0]
            with z.open(archivo_txt) as f:
                df = pd.read_csv(f, sep="|", dtype=str, encoding="latin-1")
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
            tablas = pd.read_html(io.StringIO(resp.text), dtype=str)
            for df in tablas:
                cols_ruc = [c for c in df.columns if "RUC" in str(c).upper()]
                if cols_ruc:
                    tmp = pd.DataFrame({
                        "ruc_ssco": df[cols_ruc[0]].astype(str).str.strip(),
                        "alerta_ssco": tipo_lista,
                    })
                    registros.append(tmp)
        except Exception as e:
            errores.append(f"{tipo_lista}: {e}")

    if not registros:
        return pd.DataFrame(columns=["ruc_ssco", "alerta_ssco"]), errores

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
            numero = partes[1].strip().lstrip("0")

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


# --- 4. EXPORTADOR A EXCEL ---
def generar_excel(df_final: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="SIRE_Consolidado")
        if "alerta_ssco" in df_final.columns:
            df_ssco = df_final[df_final["alerta_ssco"].notna()]
            df_ssco.to_excel(writer, index=False, sheet_name="Alertas_SSCO")
    return buffer.getvalue()


# --- INTERFAZ LATERAL (CREDENCIALES EN SESIÓN, NUNCA EN GITHUB) ---
with st.sidebar:
    st.header("Credenciales SUNAT")
    ruc = st.text_input("RUC", max_chars=11)
    usuario_sol = st.text_input("Usuario SOL")
    clave_sol = st.text_input("Clave SOL", type="password")
    client_id = st.text_input("API Client ID (SIRE)")
    client_secret = st.text_input("API Client Secret (SIRE)", type="password")
    periodo = st.text_input("Periodo (AAAAMM)", value="202610", max_chars=6)
    btn_descargar = st.button("1. Descargar Propuesta SIRE", type="primary", use_container_width=True)

if "df_sire" not in st.session_state:
    st.session_state.df_sire = None

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

                # Detectar columnas clave del archivo TXT de propuesta RCE
                col_ruc_prov = next((c for c in df_sire.columns if "Nro Doc Identidad" in c), None)
                col_serie = next((c for c in df_sire.columns if "Serie" in c), None)
                col_num = next((c for c in df_sire.columns if "Nro CP" in c), None)

                padron_ssco, errores_ssco = obtener_padron_ssco()
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
                            + df_sire[col_num].astype(str).str.strip().str.lstrip("0")
                        )

                st.session_state.df_sire = df_sire
                st.success(f"Propuesta cargada: {len(df_sire)} comprobantes encontrados.")
            except Exception as e:
                st.error(f"Error en la consulta: {e}")

# --- VISUALIZACIÓN Y CRUCE CON XML ---
if st.session_state.df_sire is not None:
    df_trabajo = st.session_state.df_sire.copy()

    # Métricas rápidas
    total_docs = len(df_trabajo)
    total_ssco = int(df_trabajo["alerta_ssco"].notna().sum()) if "alerta_ssco" in df_trabajo.columns else 0
    c1, c2 = st.columns(2)
    c1.metric("Comprobantes en Propuesta RCE", total_docs)
    c2.metric("Proveedores con Alerta SSCO", total_ssco)

    st.divider()
    st.subheader("2. Agregar descripciones desde archivos XML")
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
            st.info(f"Se vincularon descripciones en {coincidencias} de {total_docs} comprobantes.")

    columnas_ocultas = ["clave_cruce"]
    df_mostrar = df_trabajo.drop(columns=[c for c in columnas_ocultas if c in df_trabajo.columns])
    st.dataframe(df_mostrar, use_container_width=True)

    st.download_button(
        label="Descargar Excel Consolidado (SIRE + Descripciones + SSCO)",
        data=generar_excel(df_mostrar),
        file_name=f"SIRE_RCE_{ruc}_{periodo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
