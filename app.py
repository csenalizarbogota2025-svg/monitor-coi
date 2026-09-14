import streamlit as st
import pandas as pd
import re, io
import fitz

st.set_page_config(page_title='Monitor COI – SDM Bogotá', page_icon='📊', layout='wide')

COMPANIES = {
    'CONSORCIO SEÑALIZAR BOGOTÁ 2025': [
        'CONSORCIO SEÑALIZAR BOGOTA 2025', 'CONSORCIO SEÑALIZAR BOGOTÁ 2025',
        'CONSORCIO SEÑALIZAR BOGOTA', 'CONSORCIO SEÑALIZAR BOGOTÁ',
    ],
    'CONSORCIO SEGURVIAL BOGOTÁ': [
        'CONSORCIO SEGURVIAL BOGOTA', 'CONSORCIO SEGURVIAL BOGOTÁ',
        'CONSORCIO SEG VIAL BOGOTA', 'CONSORCIO SEG VIAL BOGOTÁ',
        'CONSORCIO SEGURIDAD VIAL BOGOTA 2025', 'CONSORCIO SEGURIDAD VIAL BOGOTA',
        'CONSORCIO SEGURIDAD VIAL',
    ],
    'SOCINTER S.A.S.': ['SOCINTER SAS', 'SOCINTER S.A.S.', 'SOCINTER S A S', 'SOCINTER'],
}

# X limits of the official PM02-PR01-F04 COI table in the supplied PDF.
# They are deliberately based on the actual table geometry, not on the order of extracted text.
X = [75, 125, 195, 265, 355, 470, 590, 655, 720, 800, 875, 950, 1600, 1695, 1810, 1880, 2015]
COLS = [
    'No.', 'CIV INICIO', 'CIV FIN', 'DIRECCIÓN DE LA OBRA INICIO',
    'DIRECCIÓN DE LA OBRA FIN', 'CONTRATISTA', 'FECHA INICIO', 'FECHA FIN',
    'HORARIO DE TRABAJO', 'HORARIO DE CIERRE', 'No. CONTRATO', 'OBSERVACIONES',
    'AUTORIZADO', 'LOCALIDAD', 'ING. RESPONSABLE', 'No RADICADO SDM'
]


def norm(s):
    s = str(s or '').upper().replace('\u00ad', '')
    s = s.translate(str.maketrans('ÁÉÍÓÚÜÑ', 'AEIOUUN'))
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def clean(s):
    s = re.sub(r'\s+', ' ', str(s or '')).strip(' |:-')
    return s


def company_from_contractor(contractor):
    n = norm(contractor)
    for company, aliases in COMPANIES.items():
        for alias in aliases:
            a = norm(alias)
            if a and a in n:
                return company
    return None


def get_printed_page(text, fallback):
    m = re.search(r'P[áa]gina\s+(\d+)\s+de\s+(\d+)', text, re.I)
    return int(m.group(1)) if m else fallback


def get_section(text):
    m = re.search(r'SECCI[ÓO]N\s+[^\n]+', text, re.I)
    if m:
        return clean(m.group(0))
    return ''


def horizontal_table_lines(page):
    vals=[]
    for dr in page.get_drawings():
        r=dr['rect']
        if r.width > page.rect.width*0.70 and r.height < 2.5 and r.y0 > 290 and r.y0 < page.rect.height-250:
            vals.append(float(r.y0))
    vals=sorted(vals)
    out=[]
    for y in vals:
        if not out or abs(y-out[-1])>1.5:
            out.append(y)
    return out


def words_in_interval(page, top, bottom):
    words=[]
    for w in page.get_text('words'):
        x0,y0,x1,y1,t,*_=w
        cy=(y0+y1)/2
        if top <= cy < bottom and x0 >= X[0]-10 and x1 <= X[-1]+10:
            words.append((x0,y0,x1,y1,t))
    return words


def field_text(words, left, right):
    vals=[w for w in words if ((w[0]+w[2])/2) >= left and ((w[0]+w[2])/2) < right]
    vals.sort(key=lambda w:(w[1],w[0]))
    # Keep line structure internally, then normalize spaces. This preserves multiline names and notes.
    return clean(' '.join(w[4] for w in vals))


def is_data_row(words):
    # First column must contain a row number (usually 28xxx) or, defensively, a numeric token.
    first=[w for w in words if ((w[0]+w[2])/2) < X[1]]
    txt=clean(' '.join(w[4] for w in sorted(first,key=lambda w:(w[1],w[0]))))
    return bool(re.match(r'^\d{4,7}\b', txt))


def classify(row):
    auth=norm(row['AUTORIZADO'])
    obs=norm(row['OBSERVACIONES'])
    if 'FORMALIZACION' in auth or 'FORMALIZA LA EMERGENCIA' in obs or 'EMERGENCIA' in auth:
        return 'FORMALIZACIÓN DE EMERGENCIA'
    if auth.startswith('NO') or 'NO' == auth:
        return 'NO AUTORIZADO'
    if auth.startswith('SI') or auth == 'SÍ':
        return 'AUTORIZADO'
    if 'NO AUTORIZA PMT' in obs:
        return 'NO AUTORIZADO'
    if 'AUTORIZA PMT' in obs:
        return 'AUTORIZADO'
    return 'OTRO'


def parse_pdf(data):
    doc=fitz.open(stream=data,filetype='pdf')
    records=[]
    page_hits={c:[] for c in COMPANIES}
    pages_scanned=0
    for pidx,page in enumerate(doc):
        text=page.get_text('text')
        ntext=norm(text)
        if 'CODIGO DE IDENTIFICACION VIAL CIV' not in ntext or 'DIRECCION DE LA OBRA' not in ntext or 'CONTRATISTA' not in ntext:
            continue
        pages_scanned += 1
        lines=horizontal_table_lines(page)
        # A COI table starts at the first long line around y=304 and each following long line closes a row.
        start_candidates=[y for y in lines if 295 <= y <= 340]
        if not start_candidates:
            continue
        start=start_candidates[0]
        row_lines=[y for y in lines if y > start+5]
        for top,bottom in zip([start]+row_lines[:-1], row_lines):
            words=words_in_interval(page, top+1, bottom-1)
            if not words or not is_data_row(words):
                continue
            row={}
            for i,col in enumerate(COLS):
                row[col]=field_text(words,X[i],X[i+1])
            # Remove accidental header/footer bleed.
            if not re.match(r'^\d{4,7}\b', row['No.']):
                continue
            contractor=clean(row['CONTRATISTA'])
            company=company_from_contractor(contractor)
            if not company:
                continue
            row['Empresa detectada']=company
            row['ESTADO INTERPRETADO']=classify(row)
            row['PÁGINA PDF']=pidx+1
            row['PÁGINA COI']=get_printed_page(text,pidx+1)
            row['SECCIÓN']=get_section(text)
            records.append(row)
            if pidx+1 not in page_hits[company]: page_hits[company].append(pidx+1)
    return doc,pd.DataFrame(records,columns=COLS+['Empresa detectada','ESTADO INTERPRETADO','PÁGINA PDF','PÁGINA COI','SECCIÓN']),page_hits,pages_scanned


def filtered_pdf(doc,pages):
    out=fitz.open()
    for p in sorted(set(pages)):
        out.insert_pdf(doc,from_page=p-1,to_page=p-1)
    return out.tobytes()


def make_excel(df):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,sheet_name='COI - Registros')
        resumen=df.groupby(['Empresa detectada','ESTADO INTERPRETADO']).size().unstack(fill_value=0)
        resumen.to_excel(writer,sheet_name='Resumen')
        pages=df.groupby('Empresa detectada')['PÁGINA PDF'].nunique().reset_index(name='Páginas PDF')
        pages.to_excel(writer,index=False,sheet_name='Páginas')
    return out.getvalue()

st.title('📊 Monitor COI – Secretaría Distrital de Movilidad')
st.caption('Extracción por geometría de la tabla oficial PM02-PR01-F04. Cada fila se procesa de forma independiente para evitar cruces entre CIV, contratista, fechas, observaciones y radicados.')

with st.sidebar:
    st.header('Empresas')
    selected=st.multiselect('Mostrar',list(COMPANIES),default=list(COMPANIES))
    st.markdown('---')
    st.info('El sistema busca la empresa únicamente en la columna CONTRATISTA. Esto evita que una mención de otra empresa dentro de OBSERVACIONES sea tomada como el contratista del registro.')

pdf_file=st.file_uploader('📄 Cargar COI en PDF',type=['pdf'])

if pdf_file:
    data=pdf_file.getvalue()
    with st.spinner('Leyendo estructura de tablas y separando registros...'):
        doc,df,hits,pages_scanned=parse_pdf(data)
    if df.empty:
        st.error('No se encontraron registros de las empresas seleccionadas. Revisa que el nombre esté en la columna CONTRATISTA del COI.')
    else:
        df=df[df['Empresa detectada'].isin(selected)].copy()
        st.success(f'Análisis terminado: {len(df)} registro(s) individual(es) encontrados en {pages_scanned} páginas con formato COI.')
        a,b,c,d,e=st.columns(5)
        a.metric('Registros',len(df))
        b.metric('Autorizados',int((df['ESTADO INTERPRETADO']=='AUTORIZADO').sum()))
        c.metric('No autorizados',int((df['ESTADO INTERPRETADO']=='NO AUTORIZADO').sum()))
        d.metric('Emergencias',int((df['ESTADO INTERPRETADO']=='FORMALIZACIÓN DE EMERGENCIA').sum()))
        e.metric('Empresas',df['Empresa detectada'].nunique())

        st.subheader('Resultados — mismas columnas del COI')
        st.dataframe(df,use_container_width=True,hide_index=True,height=560)

        st.subheader('Resumen por empresa')
        st.dataframe(df.groupby(['Empresa detectada','ESTADO INTERPRETADO']).size().reset_index(name='Cantidad'),use_container_width=True,hide_index=True)

        st.subheader('Descargas')
        st.download_button('📊 Descargar Excel con columnas del COI',make_excel(df),'Monitor_COI_resultados.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        for company in selected:
            pages=hits.get(company,[])
            if not pages: continue
            pdf_bytes=filtered_pdf(doc,pages)
            safe=re.sub(r'[^A-Za-z0-9]+','_',company).strip('_')
            st.download_button(f'📕 PDF – {company} ({len(pages)} páginas)',pdf_bytes,f'{safe}_COI.pdf','application/pdf',key='pdf_'+safe)
else:
    st.markdown('''
    ### Versión 2 — extracción fila por fila
    Esta versión no toma una página completa como un registro. Lee las líneas de la tabla oficial y reconstruye cada fila usando la posición real de cada columna.

    **Incluye:**
    - Las columnas originales del COI.
    - Separación correcta de varios registros que aparecen en una misma página.
    - Empresa identificada exclusivamente desde CONTRATISTA.
    - CIV inicio y CIV fin por separado.
    - Dirección inicio y fin por separado.
    - Fecha, horario de trabajo y horario de cierre.
    - Observaciones completas del registro.
    - Autorizado, localidad, ingeniero y radicado.
    - Página PDF y página COI.
    - Estado interpretado sin alterar el valor original de AUTORIZADO.
    ''')
