import streamlit as st
import pandas as pd
import re, io
import fitz
import plotly.express as px

st.set_page_config(
    page_title="Monitor COI – SDM Bogotá",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CONFIGURACIÓN
# ============================================================

COMPANIES = {
    "CONSORCIO SEÑALIZAR BOGOTÁ 2025": {
        "aliases": [
            "CONSORCIO SEÑALIZAR BOGOTA 2025",
            "CONSORCIO SEÑALIZAR BOGOTÁ 2025",
            "CONSORCIO SEÑALIZAR BOGOTA",
            "CONSORCIO SEÑALIZAR BOGOTÁ",
            "CONSORCIO SEÑALIZAR",
        ],
        "contract": "2024-3652",
    },
    "CONSORCIO SEGURVIAL BOGOTÁ": {
        "aliases": [
            "CONSORCIO SEGURVIAL BOGOTA",
            "CONSORCIO SEGURVIAL BOGOTÁ",
            "CONSORCIO SEG VIAL BOGOTA",
            "CONSORCIO SEG VIAL BOGOTÁ",
            "CONSORCIO SEGURIDAD VIAL BOGOTA 2025",
            "CONSORCIO SEGURIDAD VIAL BOGOTA",
            "CONSORCIO SEGURIDAD VIAL",
            "SEGURVIAL BOGOTA",
        ],
        "contract": "2024-3651",
    },
    "SOCINTER S.A.S.": {
        "aliases": [
            "SOCINTER SAS",
            "SOCINTER S.A.S.",
            "SOCINTER S A S",
            "SOCINTER",
        ],
        "contract": None,
    },
}

# Columnas del formato oficial COI PM02-PR01-F04
COLS = [
    "No.",
    "CIV INICIO",
    "CIV FIN",
    "DIRECCIÓN DE LA OBRA INICIO",
    "DIRECCIÓN DE LA OBRA FIN",
    "CONTRATISTA",
    "FECHA INICIO",
    "FECHA FIN",
    "HORARIO DE TRABAJO",
    "HORARIO DE CIERRE",
    "No. CONTRATO",
    "OBSERVACIONES",
    "AUTORIZADO",
    "LOCALIDAD",
    "ING. RESPONSABLE",
    "No RADICADO SDM",
]

# Límites horizontales reales de la tabla del COI.
X = [75, 125, 195, 265, 355, 470, 590, 655, 720, 800, 875, 950, 1600, 1695, 1810, 1880, 2015]


# ============================================================
# FUNCIONES DE TEXTO / IDENTIFICACIÓN
# ============================================================

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
    """
    Regla de identificación:
    1. Para Segurvial y Señalizar, el No. CONTRATO es obligatorio.
    2. El contrato también puede identificar directamente la empresa.
    3. SOCINTER se identifica por CONTRATISTA porque no se ha definido un contrato único.
    4. Nunca se busca el nombre de la empresa dentro de OBSERVACIONES.
    """
    c = normalize_contract(contract)
    contractor_matches = []

    for company, cfg in COMPANIES.items():
        if alias_match(contractor, cfg["aliases"]):
            contractor_matches.append(company)

    # Contratos únicos y verificables
    for company, cfg in COMPANIES.items():
        expected = normalize_contract(cfg["contract"])
        if expected and c == expected:
            # Para evitar falsos positivos, el contrato prevalece como identificador.
            return company, "CONTRATO + CONTRATISTA" if company in contractor_matches else "CONTRATO"

    # SOCINTER u otros casos donde el contrato no está configurado
    for company in contractor_matches:
        if COMPANIES[company]["contract"] is None:
            return company, "CONTRATISTA"

    return None, ""


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


# ============================================================
# EXTRACCIÓN ROBUSTA POR FILA
# ============================================================

def row_starts(page):
    """
    Detecta cada fila usando el número de registro de la primera columna.
    Esto es más estable que tomar una página completa como un único registro.
    """
    starts = []
    for w in page.get_text("words"):
        x0, y0, x1, y1, t, *_ = w
        if 70 <= x0 <= 125 and x1 <= 135 and y0 > 300 and re.fullmatch(r"\d{4,7}", t.strip()):
            starts.append((float(y0), t.strip()))

    # Deduplicar por coordenada vertical
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
    vals = [
        w for w in words
        if ((w[0] + w[2]) / 2) >= left and ((w[0] + w[2]) / 2) < right
    ]
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

        # Solo páginas que realmente tienen el encabezado/estructura COI
        if "CODIGO DE IDENTIFICACION VIAL CIV" not in ntext:
            continue
        if "CONTRATISTA" not in ntext or "No CONTRATO" not in ntext:
            continue

        starts = row_starts(page)
        if not starts:
            continue

        pages_scanned += 1

        for i, (y_start, row_no) in enumerate(starts):
            top = max(300, y_start - 3)
            if i + 1 < len(starts):
                bottom = starts[i + 1][0] - 3
            else:
                # Evitar incluir el pie de página como parte de observaciones
                bottom = page.rect.height - 55

            if bottom <= top:
                continue

            words = words_in_interval(page, top, bottom)
            if not words:
                continue

            row = {}
            for j, col in enumerate(COLS):
                row[col] = field_text(words, X[j], X[j + 1])

            # El número de la primera columna se toma directamente del detector.
            row["No."] = row_no

            contractor = clean(row["CONTRATISTA"])
            contract = clean(row["No. CONTRATO"])

            company, method = identify_company(contractor, contract)

            if not company:
                continue

            row["Empresa detectada"] = company
            row["MÉTODO IDENTIFICACIÓN"] = method
            row["ESTADO INTERPRETADO"] = classify(row)
            row["PÁGINA PDF"] = pidx + 1
            row["PÁGINA COI"] = get_printed_page(text, pidx + 1)
            row["SECCIÓN"] = get_section(text)

            expected = COMPANIES[company]["contract"]
            row["CONTRATO ESPERADO"] = expected if expected else "No configurado"
            row["CONTRATO VALIDADO"] = (
                "SÍ" if expected is None or normalize_contract(contract) == normalize_contract(expected)
                else "NO"
            )

            records.append(row)

            if pidx + 1 not in page_hits[company]:
                page_hits[company].append(pidx + 1)

    extra = [
        "Empresa detectada",
        "MÉTODO IDENTIFICACIÓN",
        "ESTADO INTERPRETADO",
        "PÁGINA PDF",
        "PÁGINA COI",
        "SECCIÓN",
        "CONTRATO ESPERADO",
        "CONTRATO VALIDADO",
    ]
    columns = COLS + extra
    return doc, pd.DataFrame(records, columns=columns), page_hits, pages_scanned


# ============================================================
# EXPORTACIONES
# ============================================================

def filtered_pdf(doc, pages):
    out = fitz.open()
    for p in sorted(set(pages)):
        out.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
    return out.tobytes()


def make_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="COI - Registros")

        resumen = (
            df.groupby(["Empresa detectada", "ESTADO INTERPRETADO"])
            .size()
            .unstack(fill_value=0)
        )
        resumen.to_excel(writer, sheet_name="Resumen")

        contratos = (
            df.groupby(["Empresa detectada", "No. CONTRATO"])
            .size()
            .reset_index(name="Registros")
        )
        contratos.to_excel(writer, index=False, sheet_name="Contratos")

        paginas = (
            df.groupby("Empresa detectada")["PÁGINA PDF"]
            .nunique()
            .reset_index(name="Páginas PDF")
        )
        paginas.to_excel(writer, index=False, sheet_name="Páginas")

    return out.getvalue()


# ============================================================
# ESTILO
# ============================================================

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    .hero {
        padding: 1.1rem 1.3rem;
        border-radius: 16px;
        background: linear-gradient(135deg, #f7f9fc 0%, #eef4ff 100%);
        border: 1px solid #dce6f5;
        margin-bottom: 1rem;
    }
    .hero h1 {margin: 0; font-size: 2.05rem;}
    .hero p {margin: .35rem 0 0 0; color: #5d6878;}
    .small-note {color:#667085; font-size:.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>📊 Monitor COI – Secretaría Distrital de Movilidad</h1>
        <p>Analizador del COI oficial con identificación por <b>contratista + No. de contrato</b>, filtros y trazabilidad hasta la página original.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("🔎 Filtros")

    selected_companies = st.multiselect(
        "Empresa",
        list(COMPANIES.keys()),
        default=list(COMPANIES.keys()),
    )

    st.caption("Contratos configurados:")
    st.markdown(
        "- **Segurvial:** `2024-3651`\n"
        "- **Señalizar Bogotá 2025:** `2024-3652`\n"
        "- **SOCINTER:** identificación por contratista"
    )

    st.divider()
    search_name = st.text_input(
        "Buscar por nombre / contratista",
        placeholder="Ej. SEGURVIAL, SOCINTER..."
    )
    search_contract = st.text_input(
        "Buscar por No. CONTRATO",
        placeholder="Ej. 2024-3651"
    )

    st.divider()
    st.info(
        "La aplicación NO usa las menciones de empresas dentro de OBSERVACIONES para identificar al contratista. "
        "Para Segurvial y Señalizar, el contrato sirve como validación adicional."
    )

pdf_file = st.file_uploader(
    "📄 Cargar COI en PDF",
    type=["pdf"],
    help="Carga el PDF oficial del COI. La aplicación reconstruye cada fila del formato original.",
)

# ============================================================
# ESTADO INICIAL
# ============================================================

if not pdf_file:
    st.subheader("¿Qué hace esta versión?")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### 🎯 Identificación precisa")
        st.write("Usa la columna CONTRATISTA y el No. CONTRATO para reducir falsos positivos.")
    with c2:
        st.markdown("### 📋 Mismas columnas")
        st.write("Conserva las 16 columnas del formato COI y agrega campos de análisis al final.")
    with c3:
        st.markdown("### 📊 Dashboard")
        st.write("Filtra por empresa, contratista, contrato y estado; consulta y descarga los resultados.")
    st.divider()
    st.markdown("**Empresas configuradas:**")
    st.write("• CONSORCIO SEÑALIZAR BOGOTÁ 2025 — contrato 2024-3652")
    st.write("• CONSORCIO SEGURVIAL BOGOTÁ — contrato 2024-3651")
    st.write("• SOCINTER S.A.S.")
    st.caption("Próxima mejora prevista: botón para buscar automáticamente el COI más reciente publicado por la SDM.")
    st.stop()

# ============================================================
# PROCESAMIENTO
# ============================================================

data = pdf_file.getvalue()

with st.spinner("Leyendo el COI y separando cada fila individual..."):
    doc, all_df, hits, pages_scanned = parse_pdf(data)

if all_df.empty:
    st.error(
        "No se encontraron registros de las empresas configuradas. "
        "Verifica que el PDF corresponda al formato oficial del COI."
    )
    st.stop()

df = all_df[all_df["Empresa detectada"].isin(selected_companies)].copy()

# Filtros
if search_name.strip():
    q = norm(search_name)
    df = df[df["CONTRATISTA"].map(norm).str.contains(re.escape(q), na=False)]

if search_contract.strip():
    q = normalize_contract(search_contract)
    df = df[df["No. CONTRATO"].map(normalize_contract).str.contains(re.escape(q), na=False)]

# Selectores secundarios
statuses = sorted(df["ESTADO INTERPRETADO"].dropna().unique().tolist())
localities = sorted(df["LOCALIDAD"].dropna().unique().tolist())

f1, f2 = st.columns(2)
with f1:
    status_filter = st.multiselect("Estado", statuses, default=statuses)
with f2:
    locality_filter = st.multiselect("Localidad", localities, default=localities)

if status_filter:
    df = df[df["ESTADO INTERPRETADO"].isin(status_filter)]
if locality_filter:
    df = df[df["LOCALIDAD"].isin(locality_filter)]

# ============================================================
# DASHBOARD
# ============================================================

st.success(
    f"Análisis terminado: **{len(df)} registro(s)** visibles. "
    f"Se revisaron **{pages_scanned} páginas** con estructura COI."
)

total = len(df)
authorized = int((df["ESTADO INTERPRETADO"] == "AUTORIZADO").sum())
not_auth = int((df["ESTADO INTERPRETADO"] == "NO AUTORIZADO").sum())
emergency = int((df["ESTADO INTERPRETADO"] == "FORMALIZACIÓN DE EMERGENCIA").sum())
other = int((df["ESTADO INTERPRETADO"] == "OTRO").sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("📄 Registros", total)
k2.metric("🟢 Autorizados", authorized)
k3.metric("🔴 No autorizados", not_auth)
k4.metric("🟠 Emergencias", emergency)
k5.metric("🏢 Empresas", df["Empresa detectada"].nunique())

tab_dash, tab_rows, tab_download = st.tabs(
    ["📊 Dashboard", "📋 Registros del COI", "📥 Descargas"]
)

with tab_dash:
    left, right = st.columns(2)

    with left:
        st.markdown("#### Estado de los registros")
        status_counts = (
            df["ESTADO INTERPRETADO"]
            .value_counts()
            .rename_axis("Estado")
            .reset_index(name="Cantidad")
        )
        fig = px.pie(
            status_counts,
            names="Estado",
            values="Cantidad",
            hole=0.52,
        )
        fig.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            legend_title_text="",
            height=360,
        )
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("#### Registros por empresa")
        company_counts = (
            df["Empresa detectada"]
            .value_counts()
            .rename_axis("Empresa")
            .reset_index(name="Cantidad")
        )
        fig2 = px.bar(
            company_counts,
            x="Empresa",
            y="Cantidad",
            text="Cantidad",
        )
        fig2.update_traces(textposition="outside")
        fig2.update_layout(
            margin=dict(l=10, r=10, t=10, b=80),
            xaxis_title="",
            yaxis_title="Registros",
            height=360,
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Empresa vs. estado")
    cross = (
        df.groupby(["Empresa detectada", "ESTADO INTERPRETADO"])
        .size()
        .reset_index(name="Cantidad")
    )
    fig3 = px.bar(
        cross,
        x="Empresa detectada",
        y="Cantidad",
        color="ESTADO INTERPRETADO",
        text="Cantidad",
        barmode="stack",
    )
    fig3.update_layout(
        margin=dict(l=10, r=10, t=10, b=80),
        xaxis_title="",
        yaxis_title="Registros",
        height=420,
    )
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("#### Contratos encontrados")
    contract_counts = (
        df.groupby(["Empresa detectada", "No. CONTRATO"])
        .size()
        .reset_index(name="Registros")
    )
    st.dataframe(contract_counts, use_container_width=True, hide_index=True)

with tab_rows:
    st.markdown("### 📋 Registros — columnas originales del COI")

    # Las columnas oficiales quedan primero y las de análisis después.
    official_view = COLS
    extra_view = [
        "Empresa detectada",
        "MÉTODO IDENTIFICACIÓN",
        "ESTADO INTERPRETADO",
        "PÁGINA PDF",
        "PÁGINA COI",
        "CONTRATO VALIDADO",
        "SECCIÓN",
    ]

    show_extra = st.checkbox(
        "Mostrar columnas de análisis adicionales",
        value=True,
    )

    display_cols = official_view + extra_view if show_extra else official_view

    st.dataframe(
        df[display_cols],
        use_container_width=True,
        hide_index=True,
        height=620,
        column_config={
            "OBSERVACIONES": st.column_config.TextColumn(
                "OBSERVACIONES",
                width="large",
            ),
            "CONTRATISTA": st.column_config.TextColumn(
                "CONTRATISTA",
                width="medium",
            ),
            "PÁGINA PDF": st.column_config.NumberColumn(
                "PÁGINA PDF",
                format="%d",
            ),
            "PÁGINA COI": st.column_config.NumberColumn(
                "PÁGINA COI",
                format="%d",
            ),
        },
    )

    st.caption(
        "La columna PÁGINA PDF corresponde a la página física del archivo cargado. "
        "PÁGINA COI corresponde a la numeración impresa por la SDM."
    )

with tab_download:
    st.markdown("### 📥 Exportar resultados")

    excel_bytes = make_excel(df)
    st.download_button(
        "📊 Descargar Excel",
        data=excel_bytes,
        file_name="Monitor_COI_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()
    st.markdown("### 📕 PDF por empresa")

    for company in selected_companies:
        company_df = df[df["Empresa detectada"] == company]
        if company_df.empty:
            continue

        pages = sorted(company_df["PÁGINA PDF"].unique().tolist())
        pdf_bytes = filtered_pdf(doc, pages)
        safe = re.sub(r"[^A-Za-z0-9]+", "_", company).strip("_")

        st.download_button(
            f"📕 {company} — {len(pages)} páginas",
            data=pdf_bytes,
            file_name=f"{safe}_COI.pdf",
            mime="application/pdf",
            key=f"pdf_{safe}",
            use_container_width=True,
        )

    st.divider()
    st.markdown("### 🔐 Validación de contratos")

    validation = (
        df.groupby(["Empresa detectada", "No. CONTRATO", "CONTRATO VALIDADO"])
        .size()
        .reset_index(name="Registros")
    )
    st.dataframe(validation, use_container_width=True, hide_index=True)

st.divider()
st.caption(
    "Monitor COI – SDM Bogotá | Versión 3: identificación por contratista/contrato, "
    "extracción fila por fila, filtros, dashboard y exportaciones."
)
