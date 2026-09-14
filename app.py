
import streamlit as st
import pandas as pd
import re, io, os, tempfile
from pathlib import Path
import fitz

st.set_page_config(page_title="Monitor COI – SDM Bogotá", page_icon="📊", layout="wide")

COMPANIES = {
    "CONSORCIO SEÑALIZAR BOGOTÁ 2025": [
        "CONSORCIO SEÑALIZAR BOGOTA 2025",
        "CONSORCIO SEÑALIZAR BOGOTÁ 2025",
        "CONSORCIO SEÑALIZAR",
        "SEÑALIZAR BOGOTA 2025",
    ],
    "CONSORCIO SEGURVIAL BOGOTÁ": [
        "CONSORCIO SEGURVIAL BOGOTA",
        "CONSORCIO SEGURVIAL BOGOTÁ",
        "CONSORCIO SEGURVIAL",
        "CONSORCIO SEGURIDAD VIAL BOGOTA 2025",
        "CONSORCIO SEGURIDAD VIAL",
    ],
    "SOCINTER S.A.S.": [
        "SOCINTER SAS",
        "SOCINTER S.A.S.",
        "SOCINTER S A S",
        "SOCINTER",
    ],
}

def norm(s):
    s = s.upper()
    repl = str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN")
    s = s.translate(repl)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def find_company(text):
    n = norm(text)
    hits = []
    for company, aliases in COMPANIES.items():
        if any(norm(a) in n for a in aliases):
            hits.append(company)
    return hits

def classify(text):
    n = norm(text)
    if "NO AUTORIZA PMT" in n or "NO AUTORIZA" in n:
        return "NO AUTORIZADO"
    if "AUTORIZA PMT" in n or "AUTORIZA" in n:
        if any(x in n for x in ["OBSERV", "CONDICION", "REQUER", "DEBE"]):
            return "AUTORIZADO CON OBSERVACIONES"
        return "AUTORIZADO"
    if "EMERGENCIA" in n:
        return "FORMALIZACIÓN DE EMERGENCIA"
    return "OTRO"

def extract_fields(text):
    def grab(patterns):
        for p in patterns:
            m = re.search(p, text, re.I | re.S)
            if m: return re.sub(r"\s+", " ", m.group(1)).strip(" .:-")
        return ""
    return {
        "CIV": grab([r"\bCIV\s*[:\-]?\s*([0-9A-Z\-]+)"]),
        "Dirección / tramo": grab([r"(?:DIRECCI[ÓO]N|DIRECCION)\s*[:\-]?\s*(.{1,180}?)(?:\n|CIV|CONTRATISTA|FECHA)"]),
        "Contrato": grab([r"(?:CONTRATO|CONTRATO DE OBRA)\s*[:\-]?\s*([A-Z0-9\-.]+)"]),
        "Radicado SDM": grab([r"(?:RADICADO(?:\s+SDM)?|RADICADO)\s*[:\-]?\s*([A-Z0-9\-.]+)"]),
        "Localidad": grab([r"LOCALIDAD\s*[:\-]?\s*([A-ZÁÉÍÓÚÜÑ \-]+)"]),
        "Responsable": grab([r"(?:RESPONSABLE|INGENIERO RESPONSABLE)\s*[:\-]?\s*(.{2,100}?)(?:\n|RADICADO|LOCALIDAD)"]),
    }

def parse_pdf(data):
    doc = fitz.open(stream=data, filetype="pdf")
    rows=[]
    page_hits={}
    for pno, page in enumerate(doc, start=1):
        text=page.get_text("text")
        companies=find_company(text)
        if not companies: continue
        for company in companies:
            f=extract_fields(text)
            f.update({"Empresa":company, "Estado":classify(text), "Página PDF":pno})
            f["Observaciones"] = re.sub(r"\s+", " ", text).strip()
            rows.append(f)
            page_hits.setdefault(company, []).append(pno)
    return doc, pd.DataFrame(rows), page_hits

def filtered_pdf(doc, pages):
    out=fitz.open()
    for p in sorted(set(pages)):
        out.insert_pdf(doc, from_page=p-1, to_page=p-1)
    return out.tobytes()

st.title("📊 Monitor COI – Secretaría Distrital de Movilidad")
st.caption("Analizador de COI para CONSORCIO SEÑALIZAR BOGOTÁ 2025, CONSORCIO SEGURVIAL BOGOTÁ y SOCINTER S.A.S.")

with st.sidebar:
    st.header("Empresas")
    selected = st.multiselect("Mostrar", list(COMPANIES), default=list(COMPANIES))
    st.markdown("---")
    st.info("Carga un COI en PDF. El programa identifica las páginas donde aparecen las empresas y organiza los resultados.")

pdf_file = st.file_uploader("📄 Cargar COI en PDF", type=["pdf"])

if pdf_file:
    data=pdf_file.getvalue()
    with st.spinner("Analizando el COI..."):
        doc, df, hits=parse_pdf(data)
    if df.empty:
        st.warning("No se encontraron coincidencias con las empresas configuradas.")
    else:
        df=df[df["Empresa"].isin(selected)].copy()
        st.success(f"Análisis terminado. {len(df)} coincidencia(s) de página/empresa encontradas.")
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Registros", len(df))
        c2.metric("Autorizados", int((df["Estado"]=="AUTORIZADO").sum()))
        c3.metric("No autorizados", int((df["Estado"]=="NO AUTORIZADO").sum()))
        c4.metric("Con observaciones", int((df["Estado"]=="AUTORIZADO CON OBSERVACIONES").sum()))

        st.subheader("Resultados")
        cols=["Empresa","CIV","Dirección / tramo","Contrato","Estado","Localidad","Radicado SDM","Página PDF"]
        st.dataframe(df[cols], use_container_width=True, hide_index=True)

        st.subheader("Descargas")
        excel=io.BytesIO()
        with pd.ExcelWriter(excel, engine="openpyxl") as writer:
            df.to_excel(writer,index=False,sheet_name="Resultados")
            resumen=df.groupby(["Empresa","Estado"]).size().unstack(fill_value=0)
            resumen.to_excel(writer,sheet_name="Resumen")
        st.download_button("📊 Descargar Excel", excel.getvalue(), "Resultado_COI.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        for company in selected:
            pages=hits.get(company,[])
            if not pages: continue
            pdf_bytes=filtered_pdf(doc,pages)
            safe=re.sub(r"[^A-Za-z0-9]+","_",company).strip("_")
            st.download_button(f"📕 PDF – {company} ({len(pages)} páginas)",
                               pdf_bytes, f"{safe}.pdf","application/pdf")
else:
    st.markdown("""
    ### ¿Qué hace esta versión?
    - Lee el PDF del COI.
    - Busca las tres empresas configuradas.
    - Identifica la página donde aparecen.
    - Clasifica preliminarmente la autorización.
    - Exporta los resultados a Excel.
    - Genera un PDF con las páginas encontradas para cada empresa.

    **Próxima fase:** separar varios registros/CIV que aparezcan en una misma página y conectar el botón **“Buscar COI más reciente”** con la publicación oficial de SDM.
    """)
