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
    },
    "CONSORCIO SEGURVIAL BOGOTÁ": {
        "aliases": ["CONSORCIO SEGURVIAL BOGOTA", "CONSORCIO SEGURVIAL BOGOTÁ", "CONSORCIO SEG VIAL BOGOTA", "CONSORCIO SEG VIAL BOGOTÁ", "CONSORCIO SEGURIDAD VIAL BOGOTA 2025", "CONSORCIO SEGURIDAD VIAL BOGOTA", "CONSORCIO SEGURIDAD VIAL", "SEGURVIAL BOGOTA"],
        "contract": "2024-3651",
    },
    "SOCINTER S.A.S.": {
        "aliases": ["SOCINTER SAS", "SOCINTER S.A.S.", "SOCINTER S A S", "SOCINTER"],
        "contract": None,
    },
}

COLS = [
    "No.", "CIV INICIO", "CIV FIN", "DIRECCIÓN DE LA OBRA INICIO", "DIRECCIÓN DE LA OBRA FIN",
    "CONTRATISTA", "FECHA INICIO", "FECHA FIN", "HORARIO DE TRABAJO", "HORARIO DE CIERRE",
    "No. CONTRATO", "OBSERVACIONES", "AUTORIZADO", "LOCALIDAD", "ING. RESPONSABLE", "No RADICADO SDM"
]
X = [75, 125, 195, 265, 355, 470, 590, 655, 720, 800, 875, 950, 1600, 1695, 1810, 1880, 2015]


def norm(s):
    s = str(s or "").upper().replace("\u00ad", "")
    s = s.translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip(" |:-")


def canonical_contract(s):
    """Convierte formatos como SDM-3651-2024 y 2024-3651 al mismo contrato canónico."""
    n = norm(s)
    # El COI suele extraer "SDM-3651- 2024" como "SDM 3651 2024".
    nums = re.findall(r"\d{3,6}|20\d{2}", n)
    year = next((x for x in nums if re.fullmatch(r"20\d{2}", x)), None)
    if year:
        nums2 = [x for x in nums if x != year]
        if nums2:
            return f"{year}-{nums2[-1]}"
    return re.sub(r"[^A-Z0-9]", "", n)


def alias_match(contractor, aliases):
    n = norm(contractor)
    return any(norm(a) and norm(a) in n for a in aliases)


def identify_company(contractor, contract):
    c = canonical_contract(contract)
    hits = [name for name, cfg in COMPANIES.items() if alias_match(contractor, cfg["aliases"])]
    for company in hits:
        expected = COMPANIES[company]["contract"]
        if expected is None:
            return company, "CONTRATISTA", "CONFIRMADA"
        if c == canonical_contract(expected):
            return company, "CONTRATISTA + CONTRATO", "CONFIRMADA"
        return None, "", "CONTRATO NO COINCIDE"
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
    if auth.startswith("SI") or auth.startswith("SÍ") or auth == "VIGENTE":
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
        printed_page = get_printed_page(text, pidx + 1)
        section = get_section(text)
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
            row["EMPRESA CONFIGURADA"] = company or "No configurada"
            row["IDENTIFICACIÓN CONFIGURADA"] = validation
            row["MÉTODO IDENTIFICACIÓN"] = method
            row["ESTADO INTERPRETADO"] = classify(row)
            row["PÁGINA PDF"] = pidx + 1
            row["PÁGINA COI"] = printed_page
            row["SECCIÓN"] = section
            row["CONTRATO CANÓNICO"] = canonical_contract(contract)
            row["CONTRATO ESPERADO"] = COMPANIES[company]["contract"] if company else ""
            row["CONTRATO VALIDADO"] = "SÍ" if company and (COMPANIES[company]["contract"] is None or canonical_contract(contract) == canonical_contract(COMPANIES[company]["contract"])) else "NO"
            records.append(row)
    extra = ["EMPRESA CONFIGURADA", "IDENTIFICACIÓN CONFIGURADA", "MÉTODO IDENTIFICACIÓN", "ESTADO INTERPRETADO", "PÁGINA PDF", "PÁGINA COI", "SECCIÓN", "CONTRATO CANÓNICO", "CONTRATO ESPERADO", "CONTRATO VALIDADO"]
    return doc, pd.DataFrame(records, columns=COLS + extra), pages_scanned


def filtered_pdf(doc, pages):
    out = fitz.open()
    for p in sorted(set(int(x) for x in pages)):
        out.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
    return out.tobytes()


def make_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="COI completo")
        df.groupby(["ESTADO INTERPRETADO"]).size().reset_index(name="Registros").to_excel(writer, index=False, sheet_name="Estados")
        df.groupby(["CONTRATISTA", "No. CONTRATO"]).size().reset_index(name="Registros").sort_values("Registros", ascending=False).to_excel(writer, index=False, sheet_name="Empresas y contratos")
        df.groupby("LOCALIDAD").size().reset_index(name="Registros").sort_values("Registros", ascending=False).to_excel(writer, index=False, sheet_name="Localidades")
    return out.getvalue()


def single_page_pdf(doc, page_number):
    return filtered_pdf(doc, [page_number])

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2.5rem; max-width: 1550px;}
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
<p><b>Nuevo:</b> extrae todos los registros del COI y permite filtrar cualquier contratista, cualquier contrato, estado, localidad, sección, CIV, dirección o radicado.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("🔎 Filtros")
    st.markdown("### Empresas configuradas")
    selected_config = st.multiselect("Mis empresas", list(COMPANIES.keys()), default=list(COMPANIES.keys()))
    only_config = st.checkbox("Solo mostrar mis empresas configuradas", value=False)
    st.caption("Validación de tus contratos")
    st.markdown("- **Segurvial:** `SDM-3651-2024` = `2024-3651`\n- **Señalizar Bogotá 2025:** `SDM-3652-2024` = `2024-3652`\n- **SOCINTER:** por contratista")
    st.divider()
    st.markdown("**El PDF completo también queda disponible para consulta:**")
    st.caption("Puedes seleccionar cualquier contratista y cualquier No. CONTRATO que aparezca en el COI.")
    st.divider()
    st.markdown("[🌐 Abrir portal oficial PMT de SDM](https://www.movilidadbogota.gov.co/pmt)")

pdf_file = st.file_uploader("📄 Cargar COI en PDF", type=["pdf"], help="Carga el COI oficial de la SDM.")

if not pdf_file:
    a,b,c,d = st.columns(4)
    for col, title, text in [
        (a,"📚 COI completo","Extrae todos los contratistas y contratos del archivo, no solo los configurados."),
        (b,"🔎 Filtros","Filtra por cualquier empresa/contratista y cualquier No. CONTRATO encontrado."),
        (c,"📊 Dashboard","Resumen gráfico por estado, contratista, contrato, localidad y sección."),
        (d,"📕 Trazabilidad","Mantiene página COI, página PDF y permite descargar páginas originales."),
    ]:
        with col:
            st.markdown(f"<div class='kpi'><b>{title}</b><br>{text}</div>", unsafe_allow_html=True)
    st.info("Carga el COI para comenzar.")
    st.stop()

data = pdf_file.getvalue()
with st.spinner("Extrayendo todas las filas del COI y normalizando los números de contrato..."):
    doc, all_df, pages_scanned = parse_pdf(data)

if all_df.empty:
    st.error("No fue posible extraer registros del COI. Verifica que el PDF corresponda al formato oficial.")
    st.stop()

# Base de consulta: por defecto TODO el COI.
df = all_df.copy()
if only_config:
    df = df[df["EMPRESA CONFIGURADA"].isin(selected_config)]

st.markdown(f"<span class='badge'>COI cargado: {pdf_file.name}</span> &nbsp; <span class='badge'>{pages_scanned} páginas con tablas COI</span> &nbsp; <span class='badge'>{len(all_df):,} registros extraídos</span>", unsafe_allow_html=True)

st.markdown("<div class='section-title'>🎛️ Filtros de consulta del COI completo</div>", unsafe_allow_html=True)
f1,f2,f3,f4 = st.columns([1.4,1.25,1,1])
with f1:
    contractor_options = sorted([x for x in df["CONTRATISTA"].dropna().unique() if str(x).strip()])
    contractor_filter = st.multiselect("Contratista / empresa", contractor_options)
with f2:
    contract_options = sorted([x for x in df["No. CONTRATO"].dropna().unique() if str(x).strip()])
    contract_filter = st.multiselect("No. CONTRATO", contract_options)
with f3:
    status_options = sorted(df["ESTADO INTERPRETADO"].dropna().unique().tolist())
    status_filter = st.multiselect("Estado", status_options, default=status_options)
with f4:
    locality_options = sorted([x for x in df["LOCALIDAD"].dropna().unique() if str(x).strip()])
    locality_filter = st.multiselect("Localidad", locality_options)

f5,f6 = st.columns(2)
with f5:
    section_options = sorted([x for x in df["SECCIÓN"].dropna().unique() if str(x).strip()])
    section_filter = st.multiselect("Sección", section_options)
with f6:
    q = st.text_input("🔍 Búsqueda libre", placeholder="CIV, dirección, contrato, radicado, localidad, ingeniero…")

if contractor_filter:
    df = df[df["CONTRATISTA"].isin(contractor_filter)]
if contract_filter:
    df = df[df["No. CONTRATO"].isin(contract_filter)]
if status_filter:
    df = df[df["ESTADO INTERPRETADO"].isin(status_filter)]
if locality_filter:
    df = df[df["LOCALIDAD"].isin(locality_filter)]
if section_filter:
    df = df[df["SECCIÓN"].isin(section_filter)]
if q.strip():
    nq = norm(q)
    mask = df.apply(lambda r: nq in norm(" ".join(str(r.get(c,"")) for c in COLS + ["EMPRESA CONFIGURADA","SECCIÓN","CONTRATO CANÓNICO"])), axis=1)
    df = df[mask]

if df.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

# KPIs
k=st.columns(6)
total=len(df); authorized=int((df["ESTADO INTERPRETADO"]=="AUTORIZADO").sum()); not_auth=int((df["ESTADO INTERPRETADO"]=="NO AUTORIZADO").sum()); emergency=int((df["ESTADO INTERPRETADO"]=="FORMALIZACIÓN DE EMERGENCIA").sum()); contractors=df["CONTRATISTA"].nunique(); contracts=df["No. CONTRATO"].nunique()
for col,label,value in [(k[0],"📄 Registros",total),(k[1],"🟢 Autorizados",authorized),(k[2],"🔴 No autorizados",not_auth),(k[3],"🟠 Emergencias",emergency),(k[4],"🏢 Contratistas",contractors),(k[5],"📑 Contratos",contracts)]:
    col.metric(label,value)

tab1,tab2,tab3,tab4 = st.tabs(["📊 Dashboard","📋 Registros COI","🔎 Trazabilidad","📥 Descargas"])

with tab1:
    c1,c2 = st.columns(2)
    status_counts=df["ESTADO INTERPRETADO"].value_counts().rename_axis("Estado").reset_index(name="Cantidad")
    with c1:
        st.markdown("#### Estado de los registros")
        fig=px.pie(status_counts,names="Estado",values="Cantidad",hole=.58)
        fig.update_layout(height=360,margin=dict(l=10,r=10,t=10,b=10),legend_title_text="")
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        st.markdown("#### Top contratistas")
        cc=df["CONTRATISTA"].value_counts().head(15).rename_axis("Contratista").reset_index(name="Cantidad")
        fig=px.bar(cc,y="Contratista",x="Cantidad",orientation="h",text="Cantidad")
        fig.update_layout(height=360,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)
    c3,c4 = st.columns(2)
    with c3:
        st.markdown("#### Contratos con más registros")
        ct=df["No. CONTRATO"].value_counts().head(15).rename_axis("Contrato").reset_index(name="Cantidad")
        fig=px.bar(ct,x="Contrato",y="Cantidad",text="Cantidad")
        fig.update_traces(textposition="outside")
        fig.update_layout(height=390,margin=dict(l=10,r=10,t=10,b=90),xaxis_title="",yaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)
    with c4:
        st.markdown("#### Registros por localidad")
        lc=df["LOCALIDAD"].replace("","SIN DATO").value_counts().head(15).rename_axis("Localidad").reset_index(name="Cantidad")
        fig=px.bar(lc,y="Localidad",x="Cantidad",orientation="h",text="Cantidad")
        fig.update_layout(height=390,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros")
        st.plotly_chart(fig,use_container_width=True)
    st.markdown("#### Empresas configuradas encontradas")
    configured=df[df["EMPRESA CONFIGURADA"]!="No configurada"]
    if configured.empty:
        st.info("En la selección actual no aparecen las empresas configuradas.")
    else:
        cc2=configured.groupby(["EMPRESA CONFIGURADA","No. CONTRATO"]).size().reset_index(name="Registros")
        st.dataframe(cc2,use_container_width=True,hide_index=True)

with tab2:
    st.markdown("### 📋 Todos los registros extraídos")
    st.caption("Las primeras 16 columnas son exactamente las columnas oficiales del COI. Los campos de análisis quedan después.")
    show_extra=st.checkbox("Mostrar campos de análisis",value=False)
    extra=["EMPRESA CONFIGURADA","IDENTIFICACIÓN CONFIGURADA","MÉTODO IDENTIFICACIÓN","ESTADO INTERPRETADO","PÁGINA PDF","PÁGINA COI","SECCIÓN","CONTRATO CANÓNICO","CONTRATO ESPERADO","CONTRATO VALIDADO"]
    cols=COLS+extra if show_extra else COLS
    st.dataframe(df[cols],use_container_width=True,hide_index=True,height=650,column_config={
        "OBSERVACIONES":st.column_config.TextColumn("OBSERVACIONES",width="large"),
        "CONTRATISTA":st.column_config.TextColumn("CONTRATISTA",width="medium"),
        "DIRECCIÓN DE LA OBRA INICIO":st.column_config.TextColumn("DIRECCIÓN DE LA OBRA INICIO",width="medium"),
        "DIRECCIÓN DE LA OBRA FIN":st.column_config.TextColumn("DIRECCIÓN DE LA OBRA FIN",width="medium"),
    })

with tab3:
    st.markdown("### 🔎 Trazabilidad individual")
    selected_no=st.selectbox("Seleccione un No. del COI",df["No."].astype(str).tolist())
    r=df[df["No."].astype(str)==str(selected_no)].iloc[0]
    c1,c2,c3,c4=st.columns(4)
    c1.write(f"**Contratista:** {r['CONTRATISTA']}")
    c2.write(f"**Contrato:** {r['No. CONTRATO']}")
    c3.write(f"**Página COI:** {r['PÁGINA COI']}")
    c4.write(f"**Página PDF:** {r['PÁGINA PDF']}")
    st.write(f"**Dirección:** {r['DIRECCIÓN DE LA OBRA INICIO']} → {r['DIRECCIÓN DE LA OBRA FIN']}")
    st.write(f"**CIV:** {r['CIV INICIO']} → {r['CIV FIN']}")
    st.write(f"**Estado:** {r['ESTADO INTERPRETADO']} | **Localidad:** {r['LOCALIDAD']}")
    st.write(f"**Observaciones:** {r['OBSERVACIONES']}")
    st.write(f"**Empresa configurada:** {r['EMPRESA CONFIGURADA']} | **Contrato canónico:** {r['CONTRATO CANÓNICO']}")
    page_bytes=single_page_pdf(doc,int(r['PÁGINA PDF']))
    st.download_button("📄 Descargar página original de este registro",page_bytes,file_name=f"COI_pagina_{int(r['PÁGINA PDF'])}.pdf",mime="application/pdf")

with tab4:
    st.markdown("### 📊 Excel del COI filtrado")
    st.download_button("⬇️ Descargar Excel con registros + análisis",make_excel(df),file_name="Monitor_COI_resultados.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.divider(); st.markdown("### 📕 PDF por contratista seleccionado")
    if contractor_filter:
        for contractor in contractor_filter:
            cdf=df[df["CONTRATISTA"]==contractor]
            if cdf.empty: continue
            pages=sorted(cdf["PÁGINA PDF"].unique().tolist())
            safe=re.sub(r"[^A-Za-z0-9]+","_",contractor).strip("_")
            st.download_button(f"📕 {contractor} — {len(pages)} páginas",filtered_pdf(doc,pages),file_name=f"{safe}_COI.pdf",mime="application/pdf",key=f"pdf_{safe}",use_container_width=True)
    else:
        st.info("Selecciona uno o varios contratistas en el filtro para habilitar su PDF.")

st.divider()
st.caption("Monitor COI – SDM Bogotá | v5.0 — extracción del COI completo, filtros por cualquier contratista/contrato, validación de empresas configuradas y trazabilidad por página.")
