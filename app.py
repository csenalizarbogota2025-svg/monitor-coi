import streamlit as st
import pandas as pd
import re, io
import fitz
import plotly.express as px

st.set_page_config(page_title="Monitor COI – SDM Bogotá", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

COMPANIES = {
    "CONSORCIO SEÑALIZAR BOGOTÁ 2025": {
        "aliases": ["CONSORCIO SEÑALIZAR BOGOTA 2025", "CONSORCIO SEÑALIZAR BOGOTÁ 2025", "CONSORCIO SEÑALIZAR BOGOTA", "CONSORCIO SEÑALIZAR BOGOTÁ", "CONSORCIO SEÑALIZAR"],
        "contract": "2024-3652",
        "strict_contract": True,
    },
    "CONSORCIO SEGURVIAL BOGOTÁ": {
        "aliases": ["CONSORCIO SEGURVIAL BOGOTA", "CONSORCIO SEGURVIAL BOGOTÁ", "CONSORCIO SEG VIAL BOGOTA", "CONSORCIO SEG VIAL BOGOTÁ", "CONSORCIO SEGURIDAD VIAL BOGOTA 2025", "CONSORCIO SEGURIDAD VIAL BOGOTA", "CONSORCIO SEGURIDAD VIAL", "SEGURVIAL BOGOTA"],
        "contract": "2024-3651",
        "strict_contract": True,
    },
    "SOCINTER S.A.S.": {
        "aliases": ["SOCINTER SAS", "SOCINTER S.A.S.", "SOCINTER S A S", "SOCINTER"],
        "contract": None,
        "strict_contract": False,
    },
}

# Las 16 columnas oficiales del formato COI PM02-PR01-F04.
COLS = [
    "No.", "CIV INICIO", "CIV FIN", "DIRECCIÓN DE LA OBRA INICIO", "DIRECCIÓN DE LA OBRA FIN",
    "CONTRATISTA", "FECHA INICIO", "FECHA FIN", "HORARIO DE TRABAJO", "HORARIO DE CIERRE",
    "No. CONTRATO", "OBSERVACIONES", "AUTORIZADO", "LOCALIDAD", "ING. RESPONSABLE", "No RADICADO SDM"
]

# Posiciones de las columnas en el PDF COI. Se mantienen alineadas al formato oficial observado.
X = [75, 125, 195, 265, 355, 470, 590, 655, 720, 800, 875, 950, 1600, 1695, 1810, 1880, 2015]


def norm(s):
    s = str(s or "").upper().replace("\u00ad", "")
    s = s.translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip(" |:-")


def normalize_contract(s):
    s = norm(s)
    return s.replace(" ", "").replace("/", "-").replace("_", "-")


def alias_match(contractor, aliases):
    n = norm(contractor)
    return any(norm(a) and norm(a) in n for a in aliases)


def identify_company(contractor, contract):
    """Identificación conservadora: para los dos consorcios exige contratista + No. CONTRATO."""
    c = normalize_contract(contract)
    contractor_hits = [
        company for company, cfg in COMPANIES.items()
        if alias_match(contractor, cfg["aliases"])
    ]

    for company in contractor_hits:
        cfg = COMPANIES[company]
        expected = normalize_contract(cfg["contract"])
        if cfg["strict_contract"]:
            if c and c == expected:
                return company, "CONTRATISTA + CONTRATO", "CONFIRMADA"
            # Si el nombre parece coincidir pero el contrato no, no se incluye como resultado.
            return None, "", "CONTRATO NO COINCIDE"
        return company, "CONTRATISTA", "CONFIRMADA"

    # No se asigna una empresa únicamente por el número de contrato.
    return None, "", "NO COINCIDE"


def get_printed_page(text, fallback):
    m = re.search(r"P[áa]gina\s+(\d+)\s+de\s+(\d+)", text, re.I)
    return int(m.group(1)) if m else fallback


def get_section(text):
    m = re.search(r"SECCI[ÓO]N\s+[^\n]+", text, re.I)
    return clean(m.group(0)) if m else ""


def classify(row):
    auth = norm(row.get("AUTORIZADO", ""))
    obs = norm(row.get("OBSERVACIONES", ""))
    if "FORMALIZACION" in auth or "FORMALIZA LA EMERGENCIA" in obs or "EMERGENCIA" in auth:
        return "FORMALIZACIÓN DE EMERGENCIA"
    if auth.startswith("NO") or auth == "NO":
        return "NO AUTORIZADO"
    if auth.startswith("SI") or auth == "SÍ":
        return "AUTORIZADO"
    if "NO AUTORIZA PMT" in obs:
        return "NO AUTORIZADO"
    if "AUTORIZA PMT" in obs:
        return "AUTORIZADO"
    return "OTRO"


def row_starts(page):
    starts = []
    for w in page.get_text("words"):
        x0, y0, x1, y1, t, *_ = w
        if 70 <= x0 <= 125 and x1 <= 135 and y0 > 300 and re.fullmatch(r"\d{4,7}", t.strip()):
            starts.append((float(y0), t.strip()))
    starts.sort()
    result = []
    for item in starts:
        if not result or abs(item[0] - result[-1][0]) > 2:
            result.append(item)
    return result


def words_in_interval(page, top, bottom):
    words = []
    for w in page.get_text("words"):
        x0, y0, x1, y1, t, *_ = w
        cy = (y0 + y1) / 2
        if top <= cy < bottom and x0 >= X[0] - 12 and x1 <= X[-1] + 12:
            words.append((x0, y0, x1, y1, t))
    return words


def field_text(words, left, right):
    vals = [w for w in words if left <= ((w[0] + w[2]) / 2) < right]
    vals.sort(key=lambda w: (w[1], w[0]))
    return clean(" ".join(w[4] for w in vals))


def parse_pdf(data):
    doc = fitz.open(stream=data, filetype="pdf")
    records = []
    page_hits = {c: [] for c in COMPANIES}
    pages_scanned = 0

    for pidx, page in enumerate(doc):
        text = page.get_text("text")
        ntext = norm(text)
        if "CODIGO DE IDENTIFICACION VIAL CIV" not in ntext:
            continue
        if "CONTRATISTA" not in ntext or "NO CONTRATO" not in ntext:
            continue
        starts = row_starts(page)
        if not starts:
            continue
        pages_scanned += 1

        for i, (y_start, row_no) in enumerate(starts):
            top = max(300, y_start - 3)
            bottom = starts[i + 1][0] - 3 if i + 1 < len(starts) else page.rect.height - 55
            if bottom <= top:
                continue
            words = words_in_interval(page, top, bottom)
            if not words:
                continue

            row = {col: field_text(words, X[j], X[j + 1]) for j, col in enumerate(COLS)}
            row["No."] = row_no
            contractor, contract = clean(row["CONTRATISTA"]), clean(row["No. CONTRATO"])
            company, method, validation = identify_company(contractor, contract)
            if not company:
                continue

            row["Empresa detectada"] = company
            row["MÉTODO IDENTIFICACIÓN"] = method
            row["IDENTIFICACIÓN"] = validation
            row["ESTADO INTERPRETADO"] = classify(row)
            row["PÁGINA PDF"] = pidx + 1
            row["PÁGINA COI"] = get_printed_page(text, pidx + 1)
            row["SECCIÓN"] = get_section(text)
            expected = COMPANIES[company]["contract"]
            row["CONTRATO ESPERADO"] = expected if expected else "No configurado"
            row["CONTRATO VALIDADO"] = "SÍ" if expected is None or normalize_contract(contract) == normalize_contract(expected) else "NO"
            records.append(row)
            if pidx + 1 not in page_hits[company]:
                page_hits[company].append(pidx + 1)

    extra = ["Empresa detectada", "MÉTODO IDENTIFICACIÓN", "IDENTIFICACIÓN", "ESTADO INTERPRETADO", "PÁGINA PDF", "PÁGINA COI", "SECCIÓN", "CONTRATO ESPERADO", "CONTRATO VALIDADO"]
    return doc, pd.DataFrame(records, columns=COLS + extra), page_hits, pages_scanned


def filtered_pdf(doc, pages):
    out = fitz.open()
    for p in sorted(set(int(x) for x in pages)):
        out.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
    return out.tobytes()


def make_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="COI - Registros")
        df.groupby(["Empresa detectada", "ESTADO INTERPRETADO"]).size().unstack(fill_value=0).to_excel(writer, sheet_name="Resumen")
        df.groupby(["Empresa detectada", "No. CONTRATO"]).size().reset_index(name="Registros").to_excel(writer, index=False, sheet_name="Contratos")
        df.groupby("Empresa detectada")["PÁGINA PDF"].nunique().reset_index(name="Páginas PDF").to_excel(writer, index=False, sheet_name="Páginas")
    return out.getvalue()


def single_page_pdf(doc, page_number):
    return filtered_pdf(doc, [page_number])

# ---------- ESTILO ----------
st.markdown("""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 2.5rem; max-width: 1500px;}
.hero {padding: 1.35rem 1.55rem; border-radius: 18px; background: linear-gradient(135deg,#eef5ff 0%,#f8fbff 55%,#eefaf7 100%); border:1px solid #d7e4f4; box-shadow:0 5px 18px rgba(16,24,40,.06);}
.hero h1 {margin:0; font-size:2.25rem; letter-spacing:-.02em;}
.hero p {margin:.35rem 0 0; color:#526071; font-size:1rem;}
.kpi {padding:.75rem 1rem; border:1px solid #e4e7ec; border-radius:14px; background:#fff; box-shadow:0 2px 10px rgba(16,24,40,.04);}
.section-title {font-size:1.15rem; font-weight:700; margin:.5rem 0 .7rem;}
.badge {display:inline-block; padding:.25rem .55rem; border-radius:999px; background:#eef4ff; color:#2457a6; font-size:.78rem; font-weight:600;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>📊 Monitor COI – Secretaría Distrital de Movilidad</h1>
<p>Control de PMT por <b>contratista + No. de contrato</b>, conservando las columnas oficiales del COI y dejando trazabilidad a la página original.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("🔎 Filtros")
    selected_companies = st.multiselect("Empresa", list(COMPANIES.keys()), default=list(COMPANIES.keys()))
    st.caption("Validación configurada")
    st.markdown("- **Segurvial:** `2024-3651`\n- **Señalizar Bogotá 2025:** `2024-3652`\n- **SOCINTER:** por contratista")
    st.divider()
    st.markdown("**Puedes filtrar después de cargar el PDF:**")
    st.caption("Nombre/contratista, No. CONTRATO, estado, localidad y sección.")
    st.divider()
    st.markdown("[🌐 Abrir portal oficial PMT de SDM](https://www.movilidadbogota.gov.co/pmt)")

pdf_file = st.file_uploader("📄 Cargar COI en PDF", type=["pdf"], help="Carga el COI oficial de la SDM.")

if not pdf_file:
    a,b,c,d = st.columns(4)
    for col, title, text in [
        (a,"🎯 Identificación","Cruza contratista y contrato para evitar falsos positivos."),
        (b,"📋 Formato COI","Mantiene las 16 columnas oficiales y agrega análisis al final."),
        (c,"📊 Dashboard","Indicadores, gráficos, filtros y distribución por empresa/estado."),
        (d,"📕 Trazabilidad","Página COI, página PDF y descarga de páginas por empresa."),
    ]:
        with col:
            st.markdown(f"<div class='kpi'><b>{title}</b><br><span class='small-note'>{text}</span></div>", unsafe_allow_html=True)
    st.info("Carga el COI para comenzar. La identificación de Segurvial y Señalizar exige coincidencia del nombre del contratista y del No. CONTRATO configurado.")
    st.stop()

data = pdf_file.getvalue()
with st.spinner("Analizando el COI fila por fila y validando contratista + contrato..."):
    doc, all_df, hits, pages_scanned = parse_pdf(data)

if all_df.empty:
    st.error("No se encontraron registros confirmados para las empresas configuradas. La versión actual exige coincidencia de contratista y contrato para Segurvial/Señalizar.")
    st.stop()

df = all_df[all_df["Empresa detectada"].isin(selected_companies)].copy()

# Filtros principales en la interfaz, con opciones reales del PDF.
st.markdown("<div class='section-title'>🎛️ Filtros de consulta</div>", unsafe_allow_html=True)
f1,f2,f3,f4 = st.columns([1.35,1.25,1,1])
with f1:
    contractor_options = sorted([x for x in df["CONTRATISTA"].dropna().unique() if str(x).strip()])
    contractor_filter = st.multiselect("Contratista / nombre", contractor_options)
with f2:
    contract_options = sorted([x for x in df["No. CONTRATO"].dropna().unique() if str(x).strip()])
    contract_filter = st.multiselect("No. CONTRATO", contract_options)
with f3:
    status_options = sorted(df["ESTADO INTERPRETADO"].unique().tolist())
    status_filter = st.multiselect("Estado", status_options, default=status_options)
with f4:
    locality_options = sorted([x for x in df["LOCALIDAD"].dropna().unique() if str(x).strip()])
    locality_filter = st.multiselect("Localidad", locality_options)

q = st.text_input("🔍 Búsqueda libre", placeholder="Escribe parte del nombre, CIV, dirección, contrato, radicado o localidad…")

if contractor_filter:
    df = df[df["CONTRATISTA"].isin(contractor_filter)]
if contract_filter:
    df = df[df["No. CONTRATO"].isin(contract_filter)]
if status_filter:
    df = df[df["ESTADO INTERPRETADO"].isin(status_filter)]
if locality_filter:
    df = df[df["LOCALIDAD"].isin(locality_filter)]
if q.strip():
    nq = norm(q)
    mask = df.apply(lambda r: nq in norm(" ".join(str(r.get(c,"")) for c in COLS)), axis=1)
    df = df[mask]

# KPIs
st.markdown(f"<span class='badge'>COI cargado: {pdf_file.name}</span> &nbsp; <span class='badge'>{pages_scanned} páginas COI revisadas</span>", unsafe_allow_html=True)

total=len(df); authorized=int((df["ESTADO INTERPRETADO"]=="AUTORIZADO").sum()); not_auth=int((df["ESTADO INTERPRETADO"]=="NO AUTORIZADO").sum()); emergency=int((df["ESTADO INTERPRETADO"]=="FORMALIZACIÓN DE EMERGENCIA").sum()); companies=df["Empresa detectada"].nunique()

k=st.columns(5)
for col,label,value in [(k[0],"📄 Registros",total),(k[1],"🟢 Autorizados",authorized),(k[2],"🔴 No autorizados",not_auth),(k[3],"🟠 Emergencias",emergency),(k[4],"🏢 Empresas",companies)]:
    col.metric(label,value)

if df.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

tab1,tab2,tab3 = st.tabs(["📊 Dashboard","📋 Registros COI","📥 Descargas"])

with tab1:
    c1,c2 = st.columns(2)
    status_counts=df["ESTADO INTERPRETADO"].value_counts().rename_axis("Estado").reset_index(name="Cantidad")
    with c1:
        st.markdown("#### Estado de los registros")
        fig=px.pie(status_counts,names="Estado",values="Cantidad",hole=.58)
        fig.update_layout(height=370,margin=dict(l=10,r=10,t=10,b=10),legend_title_text="")
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        st.markdown("#### Registros por empresa")
        cc=df["Empresa detectada"].value_counts().rename_axis("Empresa").reset_index(name="Cantidad")
        fig=px.bar(cc,x="Empresa",y="Cantidad",text="Cantidad")
        fig.update_traces(textposition="outside")
        fig.update_layout(height=370,margin=dict(l=10,r=10,t=10,b=80),xaxis_title="",yaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)

    c3,c4=st.columns(2)
    with c3:
        st.markdown("#### Empresa vs. estado")
        cross=df.groupby(["Empresa detectada","ESTADO INTERPRETADO"]).size().reset_index(name="Cantidad")
        fig=px.bar(cross,x="Empresa detectada",y="Cantidad",color="ESTADO INTERPRETADO",text="Cantidad",barmode="stack")
        fig.update_layout(height=430,margin=dict(l=10,r=10,t=10,b=90),xaxis_title="",yaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)
    with c4:
        st.markdown("#### Registros por localidad")
        lc=df["LOCALIDAD"].replace("","SIN DATO").value_counts().head(15).rename_axis("Localidad").reset_index(name="Cantidad")
        fig=px.bar(lc,y="Localidad",x="Cantidad",orientation="h",text="Cantidad")
        fig.update_layout(height=430,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)

    st.markdown("#### Contratos encontrados y validados")
    contracts=df.groupby(["Empresa detectada","No. CONTRATO","CONTRATO VALIDADO"]).size().reset_index(name="Registros")
    st.dataframe(contracts,use_container_width=True,hide_index=True)

with tab2:
    st.markdown("### 📋 Formato oficial + trazabilidad")
    st.caption("Las primeras 16 columnas son las columnas del COI. Los campos de análisis se muestran después.")
    show_extra=st.checkbox("Mostrar campos de análisis",value=False)
    extra=["Empresa detectada","MÉTODO IDENTIFICACIÓN","IDENTIFICACIÓN","ESTADO INTERPRETADO","PÁGINA PDF","PÁGINA COI","SECCIÓN","CONTRATO ESPERADO","CONTRATO VALIDADO"]
    cols=COLS+extra if show_extra else COLS
    st.dataframe(df[cols],use_container_width=True,hide_index=True,height=650,column_config={
        "OBSERVACIONES":st.column_config.TextColumn("OBSERVACIONES",width="large"),
        "CONTRATISTA":st.column_config.TextColumn("CONTRATISTA",width="medium"),
        "DIRECCIÓN DE LA OBRA INICIO":st.column_config.TextColumn("DIRECCIÓN DE LA OBRA INICIO",width="medium"),
        "DIRECCIÓN DE LA OBRA FIN":st.column_config.TextColumn("DIRECCIÓN DE LA OBRA FIN",width="medium"),
    })

    st.divider(); st.markdown("### 🔎 Trazabilidad individual")
    selected_no=st.selectbox("Seleccione un No. del COI",df["No."].astype(str).tolist())
    r=df[df["No."].astype(str)==str(selected_no)].iloc[0]
    c1,c2,c3=st.columns(3)
    c1.write(f"**Empresa:** {r['Empresa detectada']}")
    c2.write(f"**Contrato:** {r['No. CONTRATO']}")
    c3.write(f"**Página COI:** {r['PÁGINA COI']} | **Página PDF:** {r['PÁGINA PDF']}")
    st.write(f"**Dirección:** {r['DIRECCIÓN DE LA OBRA INICIO']} → {r['DIRECCIÓN DE LA OBRA FIN']}")
    st.write(f"**Observaciones:** {r['OBSERVACIONES']}")
    page_bytes=single_page_pdf(doc,int(r['PÁGINA PDF']))
    st.download_button("📄 Descargar página original de este registro",page_bytes,file_name=f"COI_pagina_{int(r['PÁGINA PDF'])}.pdf",mime="application/pdf")

with tab3:
    st.markdown("### 📊 Excel")
    st.download_button("⬇️ Descargar Excel con columnas COI + análisis",make_excel(df),file_name="Monitor_COI_resultados.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.divider(); st.markdown("### 📕 PDF por empresa")
    for company in selected_companies:
        cdf=df[df["Empresa detectada"]==company]
        if cdf.empty: continue
        pages=sorted(cdf["PÁGINA PDF"].unique().tolist())
        safe=re.sub(r"[^A-Za-z0-9]+","_",company).strip("_")
        st.download_button(f"📕 {company} — {len(pages)} páginas",filtered_pdf(doc,pages),file_name=f"{safe}_COI.pdf",mime="application/pdf",key=f"pdf_{safe}",use_container_width=True)

st.divider()
st.caption("Monitor COI – SDM Bogotá | v4.0 — identificación conservadora por contratista + contrato, columnas oficiales, filtros y dashboard.")
