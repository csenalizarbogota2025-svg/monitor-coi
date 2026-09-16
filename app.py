import io, re
from pathlib import Path
import fitz
import pandas as pd
import streamlit as st
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_VERSION = "8.0"
APP_NAME = "Monitor COI"
COLS = [
    "No.", "CIV INICIO", "CIV FIN", "DIRECCIÓN DE LA OBRA INICIO", "DIRECCIÓN DE LA OBRA FIN",
    "CONTRATISTA", "FECHA INICIO", "FECHA FIN", "HORARIO DE TRABAJO", "HORARIO DE CIERRE",
    "No. CONTRATO", "OBSERVACIONES", "AUTORIZADO", "LOCALIDAD", "ING. RESPONSABLE", "No RADICADO SDM"
]
ANALYSIS_COLS = [
    "ESTADO INTERPRETADO", "PÁGINA PDF", "PÁGINA COI", "SECCIÓN", "CONTRATO CANÓNICO"
]

st.set_page_config(page_title=f"{APP_NAME} – SDM Bogotá", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# ---------- Helpers ----------
def norm(s):
    s = str(s or "").upper().replace("\u00ad", "")
    s = s.translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip(" |:-")

def canonical_contract(s):
    n = norm(s)
    nums = re.findall(r"\d{3,6}|20\d{2}", n)
    year = next((x for x in nums if re.fullmatch(r"20\d{2}", x)), None)
    if year:
        others = [x for x in nums if x != year]
        if others:
            return f"{year}-{others[-1]}"
    return re.sub(r"[^A-Z0-9]", "", n)

def get_printed_page(text, fallback):
    m = re.search(r"P[áa]gina\s+(\d+)\s+de\s+(\d+)", text, re.I)
    return int(m.group(1)) if m else fallback

def get_section(text):
    m = re.search(r"SECCI[ÓO]N\s+[^\n]+", text, re.I)
    return clean(m.group(0)) if m else ""

def classify(row):
    auth = norm(row.get("AUTORIZADO", ""))
    obs = norm(row.get("OBSERVACIONES", ""))
    if "FORMALIZACION" in auth or "EMERGENCIA" in auth or "FORMALIZA LA EMERGENCIA" in obs:
        return "FORMALIZACIÓN DE EMERGENCIA"
    if auth in {"NO", "NO AUTORIZADO", "NO AUTORIZADA"} or auth.startswith("NO "):
        return "NO AUTORIZADO"
    if auth in {"SI", "VIGENTE", "AUTORIZADO", "AUTORIZADA"} or auth.startswith("SI "):
        return "AUTORIZADO"
    if "NO AUTORIZA PMT" in obs:
        return "NO AUTORIZADO"
    if "AUTORIZA PMT" in obs:
        return "AUTORIZADO"
    return "OTRO"

def looks_like_data_row(row):
    if not row or len(row) < 16:
        return False
    first = clean(row[0])
    if not re.fullmatch(r"\d{4,7}", first):
        return False
    # CIVs in the official COI are numeric. This validation catches the
    # exact failure where text from the next cell leaked into CIV.
    civ1, civ2 = clean(row[1]), clean(row[2])
    if civ1 and not re.fullmatch(r"\d{5,10}", civ1):
        return False
    if civ2 and not re.fullmatch(r"\d{5,10}", civ2):
        return False
    return True

def table_boundaries(page):
    """Obtiene los límites reales de las 16 columnas desde las líneas del PDF.
    Evita find_tables(), que es mucho más lento y puede fusionar texto de celdas.
    """
    interior=[]; outer=[]
    for dr in page.get_drawings():
        for item in dr.get("items",[]):
            if item[0] == "l":
                p1,p2=item[1],item[2]
                if abs(p1.x-p2.x)<1.5 and abs(p2.y-p1.y)>500:
                    interior.append((p1.x,p1.y,p2.y))
            elif item[0] == "re":
                r=item[1]
                if r.height>500 and r.width<5:
                    outer.append((r.x0,r.x1))
    xs=[]
    for x,_,_ in interior:
        if not any(abs(x-y)<2 for y in xs): xs.append(x)
    xs=sorted(xs)
    left=min((x0 for x0,x1 in outer), default=74.5)
    right=max((x1 for x0,x1 in outer), default=2008.5)
    # Remove any accidental outer candidates from interior and build 17 boundaries.
    xs=[x for x in xs if left+10 < x < right-10]
    if len(xs)>=15:
        # The official format has 15 interior verticals.
        xs=xs[:15]
    if len(xs)==15:
        return [left]+xs+[right]
    return []

def extract_rows_from_words(page, xs):
    words=page.get_text("words")
    # First-column numeric words define row starts.
    starts=[]
    for w in words:
        x0,y0,x1,y1,t,*_=w
        if xs[0]-1 <= x0 <= xs[1]+1 and x1 <= xs[1]+2 and y0>200 and re.fullmatch(r"\d{4,7}",t.strip()):
            starts.append((float(y0),t.strip()))
    starts.sort()
    uniq=[]
    for z in starts:
        if not uniq or abs(z[0]-uniq[-1][0])>2: uniq.append(z)
    if not uniq: return []
    tops=[z[0] for z in uniq]
    rows=[[] for _ in uniq]
    import bisect
    for w in words:
        x0,y0,x1,y1,t,*_=w
        cy=(y0+y1)/2
        ri=bisect.bisect_right(tops,cy)-1
        if ri<0: continue
        bottom=tops[ri+1]-1 if ri+1<len(tops) else page.rect.height-5
        if not (tops[ri]-1 <= cy < bottom): continue
        cx=(x0+x1)/2
        # find column by center x
        ci=bisect.bisect_right(xs,cx)-1
        if 0<=ci<16:
            rows[ri].append((y0,x0,t))
    result=[]
    for ri,rowwords in enumerate(rows):
        cells=[[] for _ in range(16)]
        for y,x,t in rowwords:
            # Recompute column for the grouped word; keeps sorting simple.
            ci=bisect.bisect_right(xs,x)-1
            if 0<=ci<16: cells[ci].append((y,x,t))
        vals=[]
        for cw in cells:
            cw.sort(key=lambda z:(z[0],z[1]))
            vals.append(clean(" ".join(z[2] for z in cw)))
        if vals and re.fullmatch(r"\d{4,7}",vals[0]) and (not vals[1] or re.fullmatch(r"\d{5,10}",vals[1])) and (not vals[2] or re.fullmatch(r"\d{5,10}",vals[2])):
            result.append(vals)
    return result

def parse_pdf(data):
    doc=fitz.open(stream=data,filetype="pdf")
    records=[]; seen=set(); pages_scanned=0
    total=len(doc)
    progress=st.progress(0,text="Preparando análisis…")
    for pidx,page in enumerate(doc):
        text=page.get_text("text")
        nt=norm(text)
        if "CODIGO DE IDENTIFICACION VIAL CIV" not in nt or "CONTRATISTA" not in nt or "NO CONTRATO" not in nt:
            progress.progress((pidx+1)/total,text=f"Revisando página {pidx+1} de {total}…")
            continue
        xs=table_boundaries(page)
        if len(xs)!=17:
            progress.progress((pidx+1)/total,text=f"Revisando página {pidx+1} de {total}…")
            continue
        rows=extract_rows_from_words(page,xs)
        if rows: pages_scanned+=1
        printed=get_printed_page(text,pidx+1); section=get_section(text)
        for vals in rows:
            key=(pidx+1,vals[0],vals[1],vals[2],vals[5],vals[10])
            if key in seen: continue
            seen.add(key)
            row=dict(zip(COLS,vals))
            row["ESTADO INTERPRETADO"]=classify(row)
            row["PÁGINA PDF"]=pidx+1
            row["PÁGINA COI"]=printed
            row["SECCIÓN"]=section
            row["CONTRATO CANÓNICO"]=canonical_contract(row["No. CONTRATO"])
            records.append(row)
        progress.progress((pidx+1)/total,text=f"Analizando página {pidx+1} de {total}…")
    progress.empty()
    return doc,pd.DataFrame(records,columns=COLS+ANALYSIS_COLS),pages_scanned,0

def schedule_analysis(df):
    if df.empty:
        return pd.DataFrame(columns=["HORARIO DE TRABAJO","REGISTROS","AUTORIZADOS","NO AUTORIZADOS","EMERGENCIAS"])
    t=df.copy(); t["HORARIO NORMALIZADO"]=t["HORARIO DE TRABAJO"].fillna("").map(clean).map(norm).replace("","SIN DATO")
    out=t.groupby("HORARIO NORMALIZADO").agg(
        REGISTROS=("No.","size"),
        AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="AUTORIZADO").sum())),
        NO_AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="NO AUTORIZADO").sum())),
        EMERGENCIAS=("ESTADO INTERPRETADO",lambda s:int((s=="FORMALIZACIÓN DE EMERGENCIA").sum())),
    ).reset_index().rename(columns={"HORARIO NORMALIZADO":"HORARIO DE TRABAJO"})
    return out.sort_values(["REGISTROS","HORARIO DE TRABAJO"],ascending=[False,True])

def company_summary(df):
    return df.groupby("CONTRATISTA",dropna=False).agg(
        REGISTROS=("No.","size"),
        AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="AUTORIZADO").sum())),
        NO_AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="NO AUTORIZADO").sum())),
        EMERGENCIAS=("ESTADO INTERPRETADO",lambda s:int((s=="FORMALIZACIÓN DE EMERGENCIA").sum())),
        CONTRATOS=("No. CONTRATO",lambda s:s.nunique()),
    ).reset_index().sort_values("REGISTROS",ascending=False)

def contract_summary(df):
    return df.groupby(["No. CONTRATO","CONTRATISTA"],dropna=False).agg(
        REGISTROS=("No.","size"),
        AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="AUTORIZADO").sum())),
        NO_AUTORIZADOS=("ESTADO INTERPRETADO",lambda s:int((s=="NO AUTORIZADO").sum())),
        EMERGENCIAS=("ESTADO INTERPRETADO",lambda s:int((s=="FORMALIZACIÓN DE EMERGENCIA").sum())),
    ).reset_index().sort_values("REGISTROS",ascending=False)

def make_excel(df, pdf_name):
    out=io.BytesIO(); horarios=schedule_analysis(df); empresas=company_summary(df); contratos=contract_summary(df)
    h24=df["HORARIO DE TRABAJO"].map(norm)=="24 HORAS"
    resumen=pd.DataFrame({"INDICADOR":["Archivo COI","Registros","Contratistas","Contratos","Autorizados","No autorizados","Formalización de emergencia","Otros estados","Registros 24 HORAS","24 HORAS autorizados"],"VALOR":[pdf_name,len(df),df["CONTRATISTA"].nunique(),df["No. CONTRATO"].nunique(),int((df["ESTADO INTERPRETADO"]=="AUTORIZADO").sum()),int((df["ESTADO INTERPRETADO"]=="NO AUTORIZADO").sum()),int((df["ESTADO INTERPRETADO"]=="FORMALIZACIÓN DE EMERGENCIA").sum()),int((df["ESTADO INTERPRETADO"]=="OTRO").sum()),int(h24.sum()),int((h24&(df["ESTADO INTERPRETADO"]=="AUTORIZADO")).sum())]})
    with pd.ExcelWriter(out,engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="COI completo")
        resumen.to_excel(w,index=False,sheet_name="Resumen")
        empresas.to_excel(w,index=False,sheet_name="Empresas")
        contratos.to_excel(w,index=False,sheet_name="Contratos")
        horarios.to_excel(w,index=False,sheet_name="Horarios")
        df.groupby("LOCALIDAD",dropna=False).size().reset_index(name="REGISTROS").sort_values("REGISTROS",ascending=False).to_excel(w,index=False,sheet_name="Localidades")
    out.seek(0); wb=load_workbook(out); wb.calculation.fullCalcOnLoad=True
    for ws in wb.worksheets:
        ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:
            cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="17365D"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.row_dimensions[1].height=30
        for col in range(1,ws.max_column+1):
            letter=get_column_letter(col); maxlen=0
            for row in ws.iter_rows(min_row=2,max_row=min(ws.max_row,250),min_col=col,max_col=col):
                v=row[0].value
                maxlen=max(maxlen,len(str(v)) if v is not None else 0)
            ws.column_dimensions[letter].width=min(max(maxlen+2,12),55)
        if ws.title=="COI completo":
            ws.column_dimensions["L"].width=70
    out2=io.BytesIO(); wb.save(out2); return out2.getvalue()

def filtered_pdf(doc,pages):
    out=fitz.open()
    for p in sorted(set(int(x) for x in pages)): out.insert_pdf(doc,from_page=p-1,to_page=p-1)
    return out.tobytes()

# ---------- UI ----------
st.markdown("""
<style>
.block-container{padding-top:1rem;padding-bottom:2rem;max-width:1600px}
.hero{padding:1.25rem 1.5rem;border-radius:22px;background:linear-gradient(135deg,#eef5ff,#ffffff 55%,#eefaf5);border:1px solid #dce6f2;box-shadow:0 7px 25px rgba(16,24,40,.07)}
.hero-row{display:flex;align-items:center;gap:16px}.logo{width:64px;height:64px;border-radius:16px;box-shadow:0 4px 12px rgba(0,0,0,.08)}
.hero h1{margin:0;font-size:2.25rem;letter-spacing:-.03em}.hero p{margin:.35rem 0 0;color:#536273}.version{font-size:.72rem;color:#718096;margin-top:.55rem}
.card{padding:1rem 1.1rem;border:1px solid #e5e9ef;border-radius:16px;background:#fff;box-shadow:0 3px 14px rgba(16,24,40,.04);height:100%}
.step{font-weight:800;color:#17365D}.small{font-size:.78rem;color:#718096}
</style>
""",unsafe_allow_html=True)

logo_svg='''<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><defs><linearGradient id="g" x1="0" x2="1"><stop stop-color="#17365D"/><stop offset="1" stop-color="#2F80ED"/></linearGradient></defs><rect width="128" height="128" rx="28" fill="url(#g)"/><rect x="22" y="70" width="16" height="32" rx="5" fill="#fff"/><rect x="47" y="52" width="16" height="50" rx="5" fill="#fff"/><rect x="72" y="35" width="16" height="67" rx="5" fill="#fff"/><path d="M22 106h76" stroke="#9FE3C4" stroke-width="7" stroke-linecap="round"/><circle cx="99" cy="29" r="10" fill="#9FE3C4"/></svg>'''
import base64
logo_b64=base64.b64encode(logo_svg.encode()).decode()

st.markdown(f'''<div class="hero"><div class="hero-row"><img class="logo" src="data:image/svg+xml;base64,{logo_b64}"><div><h1>📊 Monitor COI – SDM Bogotá</h1><p>Analiza el COI completo, descubre todas las empresas y contratos, filtra múltiples opciones y conserva la trazabilidad al PDF original.</p><div class="version">Versión {APP_VERSION} · {APP_NAME} · Secretaría Distrital de Movilidad</div></div></div></div>''',unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📄 Flujo de trabajo")
    st.markdown("**1.** Carga el COI  →  **2.** Analiza  →  **3.** Filtra y genera la lista")
    st.divider()
    st.markdown("### ℹ️ Sobre esta versión")
    st.caption(f"v{APP_VERSION} · Extracción del COI completo. Las empresas y contratos se obtienen directamente del archivo cargado; no hay empresas preconfiguradas obligatorias.")
    st.divider()
    st.markdown("[🌐 Portal oficial PMT de SDM](https://www.movilidadbogota.gov.co/pmt)")

pdf_file=st.file_uploader("📥 CARGA 1 · Selecciona el COI oficial en PDF",type=["pdf"])
if not pdf_file:
    c1,c2,c3=st.columns(3)
    for c,title,txt in [(c1,"🔍 Analiza","Extrae todos los registros, empresas, contratos, horarios, estados y localidades."),(c2,"🎛️ Filtra","Selecciona una o varias empresas, contratos, estados, localidades o usa búsqueda libre."),(c3,"📊 Resume","Dashboard + Excel con resumen por empresa, contrato, estado, horario y 24 HORAS.")]:
        with c: st.markdown(f"<div class='card'><div class='step'>{title}</div><p>{txt}</p></div>",unsafe_allow_html=True)
    st.info("Carga un COI para habilitar el análisis.")
    st.stop()

file_key=f"{pdf_file.name}:{len(pdf_file.getvalue())}"
if st.session_state.get("file_key")!=file_key:
    st.session_state.file_key=file_key
    st.session_state.analysis=None
    st.session_state.generated=False

data=pdf_file.getvalue()
colA,colB=st.columns([1,4])
with colA:
    analyze=st.button("🔎 ANALIZAR COI",type="primary",use_container_width=True)
with colB:
    st.caption("Primero analiza el archivo. Después aparecerán los filtros y el botón para generar la lista seleccionada.")

if analyze:
    with st.spinner("Analizando el COI completo y validando las celdas..."):
        st.session_state.analysis=parse_pdf(data)
    st.session_state.generated=False

if st.session_state.get("analysis") is None:
    st.warning("El archivo está cargado, pero todavía no ha sido analizado. Pulsa **🔎 ANALIZAR COI**.")
    st.stop()

doc,all_df,pages_scanned,fallback_pages=st.session_state.analysis
if all_df.empty:
    st.error("No se encontraron filas válidas del formato COI en el PDF.")
    st.stop()

st.markdown(f"<div class='small'>📄 <b>{pdf_file.name}</b> · {len(all_df):,} registros extraídos · {pages_scanned} páginas con datos · {fallback_pages} páginas resueltas con método de respaldo</div>",unsafe_allow_html=True)

st.markdown("### 🎛️ CARGA 2 · Selecciona lo que quieres incluir en la lista")
base=all_df.copy()

c1,c2=st.columns(2)
with c1:
    companies=sorted([x for x in base["CONTRATISTA"].dropna().unique() if str(x).strip()])
    company_filter=st.multiselect("🏢 Empresas / contratistas",companies,help="Puedes seleccionar una, varias o dejar vacío para todas.")
with c2:
    contracts=sorted([x for x in base["No. CONTRATO"].dropna().unique() if str(x).strip()])
    contract_filter=st.multiselect("📑 No. CONTRATO",contracts,help="Puedes seleccionar uno, varios o dejar vacío para todos.")
c3,c4,c5=st.columns([1,1,1])
with c3:
    statuses=sorted(base["ESTADO INTERPRETADO"].dropna().unique())
    status_filter=st.multiselect("Estado",statuses,default=statuses)
with c4:
    localities=sorted([x for x in base["LOCALIDAD"].dropna().unique() if str(x).strip()])
    locality_filter=st.multiselect("Localidad",localities)
with c5:
    sections=sorted([x for x in base["SECCIÓN"].dropna().unique() if str(x).strip()])
    section_filter=st.multiselect("Sección",sections)
q=st.text_input("🔍 Búsqueda libre",placeholder="CIV, dirección, contrato, radicado, ingeniero, observación…")

if st.button("⚡ GENERAR LISTA SELECCIONADA",type="primary",use_container_width=True):
    df=base.copy()
    if company_filter: df=df[df["CONTRATISTA"].isin(company_filter)]
    if contract_filter: df=df[df["No. CONTRATO"].isin(contract_filter)]
    if status_filter: df=df[df["ESTADO INTERPRETADO"].isin(status_filter)]
    if locality_filter: df=df[df["LOCALIDAD"].isin(locality_filter)]
    if section_filter: df=df[df["SECCIÓN"].isin(section_filter)]
    if q.strip():
        nq=norm(q); mask=df.apply(lambda r:nq in norm(" ".join(str(r.get(c,"")) for c in COLS+ANALYSIS_COLS)),axis=1); df=df[mask]
    st.session_state.generated_df=df
    st.session_state.generated=True

if not st.session_state.get("generated",False):
    st.info("Selecciona los filtros que quieras y pulsa **⚡ GENERAR LISTA SELECCIONADA**. Si dejas empresas y contratos vacíos, podrás generar la lista de todo el COI.")
    st.stop()

df=st.session_state.generated_df
if df.empty:
    st.warning("No hay registros con la combinación de filtros seleccionada.")
    st.stop()

# KPIs
h24=df["HORARIO DE TRABAJO"].map(norm)=="24 HORAS"
k=st.columns(7)
vals=[("📄 Registros",len(df)),("🟢 Autorizados",int((df["ESTADO INTERPRETADO"]=="AUTORIZADO").sum())),("🔴 No autorizados",int((df["ESTADO INTERPRETADO"]=="NO AUTORIZADO").sum())),("🟠 Emergencias",int((df["ESTADO INTERPRETADO"]=="FORMALIZACIÓN DE EMERGENCIA").sum())),("🕐 24 HORAS",int(h24.sum())),("🏢 Empresas",df["CONTRATISTA"].nunique()),("📑 Contratos",df["No. CONTRATO"].nunique())]
for c,(lab,val) in zip(k,vals): c.metric(lab,val)

st.success(f"Lista generada: **{len(df):,} registros** de **{df['CONTRATISTA'].nunique():,} empresas** y **{df['No. CONTRATO'].nunique():,} contratos**.")
t1,t2,t3,t4=st.tabs(["📊 Dashboard","📋 Lista COI","🔎 Trazabilidad","📥 Excel y PDF"])
with t1:
    c1,c2=st.columns(2)
    sc=df["ESTADO INTERPRETADO"].value_counts().rename_axis("Estado").reset_index(name="Cantidad")
    with c1:
        st.markdown("#### Estado")
        fig=px.pie(sc,names="Estado",values="Cantidad",hole=.58); fig.update_layout(height=360,margin=dict(l=10,r=10,t=10,b=10),legend_title_text=""); st.plotly_chart(fig,use_container_width=True)
    with c2:
        st.markdown("#### Registros por empresa")
        cc=df["CONTRATISTA"].value_counts().head(20).rename_axis("Contratista").reset_index(name="Cantidad"); fig=px.bar(cc,y="Contratista",x="Cantidad",orientation="h",text="Cantidad"); fig.update_layout(height=360,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros"); st.plotly_chart(fig,use_container_width=True)
    c3,c4=st.columns(2)
    with c3:
        st.markdown("#### Contratos")
        ct=df["No. CONTRATO"].value_counts().head(20).rename_axis("Contrato").reset_index(name="Cantidad"); fig=px.bar(ct,x="Contrato",y="Cantidad",text="Cantidad"); fig.update_layout(height=390,margin=dict(l=10,r=10,t=10,b=100),xaxis_title="",yaxis_title="Registros",xaxis_tickangle=-45); st.plotly_chart(fig,use_container_width=True)
    with c4:
        st.markdown("#### Localidades")
        lc=df["LOCALIDAD"].replace("","SIN DATO").value_counts().head(20).rename_axis("Localidad").reset_index(name="Cantidad"); fig=px.bar(lc,y="Localidad",x="Cantidad",orientation="h",text="Cantidad"); fig.update_layout(height=390,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros"); st.plotly_chart(fig,use_container_width=True)
    horarios=schedule_analysis(df)
    c5,c6=st.columns(2)
    with c5:
        st.markdown("#### ⏰ Todos los horarios")
        hp=horarios.head(25); fig=px.bar(hp,y="HORARIO DE TRABAJO",x="REGISTROS",orientation="h",text="REGISTROS"); fig.update_layout(height=470,margin=dict(l=10,r=10,t=10,b=20),yaxis_title="",xaxis_title="Registros"); st.plotly_chart(fig,use_container_width=True)
    with c6:
        st.markdown("#### ⏰ Horario vs. estado")
        long=horarios.head(20).melt(id_vars="HORARIO DE TRABAJO",value_vars=["AUTORIZADOS","NO_AUTORIZADOS","EMERGENCIAS"],var_name="Estado",value_name="Cantidad"); fig=px.bar(long,x="HORARIO DE TRABAJO",y="Cantidad",color="Estado",barmode="group",text="Cantidad"); fig.update_layout(height=470,margin=dict(l=10,r=10,t=10,b=110),xaxis_title="",yaxis_title="Registros",xaxis_tickangle=-45); st.plotly_chart(fig,use_container_width=True)
    st.markdown("#### 🕐 Resumen 24 HORAS")
    h24a=int((h24&(df["ESTADO INTERPRETADO"]=="AUTORIZADO")).sum())
    a,b,c=st.columns(3); a.metric("Registros 24 HORAS",int(h24.sum())); b.metric("24 HORAS autorizados",h24a); c.metric("24 HORAS no autorizados",int(h24.sum())-h24a)
    st.dataframe(horarios,use_container_width=True,hide_index=True)
    st.markdown("#### 🏢 Resumen por todas las empresas")
    st.dataframe(company_summary(df),use_container_width=True,hide_index=True)

with t2:
    show_extra=st.checkbox("Mostrar campos de análisis",False)
    cols=COLS+ANALYSIS_COLS if show_extra else COLS
    st.dataframe(df[cols],use_container_width=True,hide_index=True,height=680,column_config={"OBSERVACIONES":st.column_config.TextColumn(width="large"),"CONTRATISTA":st.column_config.TextColumn(width="medium"),"DIRECCIÓN DE LA OBRA INICIO":st.column_config.TextColumn(width="medium"),"DIRECCIÓN DE LA OBRA FIN":st.column_config.TextColumn(width="medium")})

with t3:
    options=df["No."].astype(str).tolist(); selected=st.selectbox("Seleccione un No. del COI",options)
    r=df[df["No."].astype(str)==selected].iloc[0]
    a,b,c,d=st.columns(4); a.write(f"**Contratista:** {r['CONTRATISTA']}"); b.write(f"**Contrato:** {r['No. CONTRATO']}"); c.write(f"**Página COI:** {r['PÁGINA COI']}"); d.write(f"**Página PDF:** {r['PÁGINA PDF']}")
    st.write(f"**CIV:** `{r['CIV INICIO']}` → `{r['CIV FIN']}`")
    st.write(f"**Dirección:** {r['DIRECCIÓN DE LA OBRA INICIO']} → {r['DIRECCIÓN DE LA OBRA FIN']}")
    st.write(f"**Horario:** {r['HORARIO DE TRABAJO']} · **Estado:** {r['ESTADO INTERPRETADO']} · **Localidad:** {r['LOCALIDAD']}")
    st.write(f"**Observaciones:** {r['OBSERVACIONES']}")
    st.download_button("📄 Descargar página original",filtered_pdf(doc,[int(r['PÁGINA PDF'])]),file_name=f"COI_pagina_{int(r['PÁGINA PDF'])}.pdf",mime="application/pdf")

with t4:
    excel=make_excel(df,pdf_file.name)
    st.download_button("📊 DESCARGAR EXCEL COMPLETO + RESUMEN",excel,file_name="Monitor_COI_resultados.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.markdown("#### 📕 PDF por empresa seleccionada")
    if company_filter:
        for company in company_filter:
            cdf=df[df["CONTRATISTA"]==company]
            if cdf.empty: continue
            pages=sorted(cdf["PÁGINA PDF"].unique().tolist()); safe=re.sub(r"[^A-Za-z0-9]+","_",company).strip("_")
            st.download_button(f"📕 {company} · {len(pages)} páginas",filtered_pdf(doc,pages),file_name=f"{safe}_COI.pdf",mime="application/pdf",key=f"pdf_{safe}",use_container_width=True)
    else: st.caption("Para generar un PDF por empresa, selecciona una o varias empresas arriba y vuelve a generar la lista.")

st.divider(); st.caption(f"Monitor COI · SDM Bogotá · versión {APP_VERSION} · {pdf_file.name} · extracción del COI completo y trazabilidad al documento original.")
