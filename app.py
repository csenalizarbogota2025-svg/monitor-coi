import io,re,html
from pathlib import Path
import fitz
import pandas as pd
import streamlit as st
import plotly.express as px
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_VERSION='12.0'
APP_NAME='Monitor COI'
COLS=['No.','CIV INICIO','CIV FIN','DIRECCIÓN DE LA OBRA INICIO','DIRECCIÓN DE LA OBRA FIN','CONTRATISTA','FECHA INICIO','FECHA FIN','HORARIO DE TRABAJO','HORARIO DE CIERRE','No. CONTRATO','OBSERVACIONES','AUTORIZADO','LOCALIDAD','ING. RESPONSABLE','No RADICADO SDM']

st.set_page_config(page_title=f'{APP_NAME} · SDM Bogotá',page_icon='📊',layout='wide',initial_sidebar_state='expanded')

st.markdown('''<style>
.block-container{padding-top:1.2rem;padding-bottom:2rem;max-width:1500px}
.hero{padding:22px 26px;border-radius:18px;background:linear-gradient(120deg,#eef6ff,#f8fbff);border:1px solid #dbe8f5;margin-bottom:18px}
.hero h1{margin:0;color:#17365d;font-size:32px}.hero p{margin:8px 0 0;color:#5d6b7a}
.small{font-size:12px;color:#718096}.step{font-weight:700;color:#17365d;font-size:18px}
.stButton>button{border-radius:10px;font-weight:700}
.metricbox{border:1px solid #e2e8f0;border-radius:14px;padding:12px 14px;background:white}
</style>''',unsafe_allow_html=True)

# ---------- Helpers ----------
def clean(s):
    return re.sub(r'\s+',' ',str(s or '').replace('\u00a0',' ')).strip()

def norm(s):
    import unicodedata
    s=clean(s).upper()
    s=''.join(c for c in unicodedata.normalize('NFD',s) if unicodedata.category(c)!='Mn')
    return re.sub(r'[^A-Z0-9]+','',s)

def contract_canonical(s):
    s=clean(s).upper().replace(' ','')
    m=re.search(r'(?:SDM-)?(\d{3,5})-(20\d{2})$',s)
    if m:return f'{m.group(2)}-{m.group(1)}'
    m=re.search(r'(20\d{2})-(\d{3,5})$',s)
    if m:return f'{m.group(1)}-{m.group(2)}'
    return s

def is_civ(s): return bool(re.fullmatch(r'\d{5,10}',clean(s)))

def find_verticals(page, y0=250):
    xs=[]
    outer=[]
    for dr in page.get_drawings():
        for item in dr.get('items',[]):
            if item[0]=='l':
                a,b=item[1],item[2]
                if abs(a.x-b.x)<1.5 and max(a.y,b.y)-min(a.y,b.y)>100 and max(a.y,b.y)>y0:
                    xs.append(round(a.x,1))
    xs=sorted(xs)
    out=[]
    for x in xs:
        if not out or abs(x-out[-1])>2: out.append(x)
    return out

def find_horizontal_rows(page):
    lines=[]
    for dr in page.get_drawings():
        for item in dr.get('items',[]):
            if item[0]=='l':
                a,b=item[1],item[2]
                if abs(a.y-b.y)<1.5 and (b.x-a.x)>1700:
                    lines.append(round((a.y+b.y)/2,1))
    lines=sorted(lines)
    out=[]
    for y in lines:
        if not out or abs(y-out[-1])>2: out.append(y)
    return out

def extract_page(page):
    words=page.get_text('words')
    # Dynamic geometry: use the actual vertical borders drawn by the PDF.
    xs=find_verticals(page)
    # Outer borders are often shorter than the internal vertical lines; recover them from full-width table lines.
    hfull=[]
    for dr in page.get_drawings():
        for item in dr.get('items',[]):
            if item[0]=='l':
                a,b=item[1],item[2]
                if abs(a.y-b.y)<1.5 and (b.x-a.x)>1700:
                    hfull.append((min(a.x,b.x),max(a.x,b.x),round((a.y+b.y)/2,1)))
    if hfull:
        xmin=min(x[0] for x in hfull); xmax=max(x[1] for x in hfull)
        xs=[xmin]+xs+[xmax]
    xs=sorted(set(round(x,1) for x in xs))
    # The official form has 17 column borders. Remove duplicates/extra lines by matching the stable layout.
    if len(xs)>17:
        target=[74,123,191,259,358,456,585,652,718,796,874,952,1593,1688,1785,1886,2009]
        chosen=[]
        for t in target:
            nearest=min(xs,key=lambda x:abs(x-t))
            if nearest not in chosen: chosen.append(nearest)
        if len(chosen)==17: xs=chosen
    if len(xs)!=17:return []
    ys=find_horizontal_rows(page)
    # row intervals are horizontal lines after header. Find intervals containing numeric No. in first cell.
    if len(ys)<2:return []
    out=[]
    for ya,yb in zip(ys,ys[1:]):
        if yb-ya<15: continue
        rowwords=[w for w in words if w[1]>=ya-1 and w[3]<=yb+1]
        nums=[w for w in rowwords if xs[0]-3<=w[0]<=xs[1]+3 and re.fullmatch(r'\d{3,8}',w[4].strip())]
        if not nums: continue
        cells=['' for _ in range(16)]
        for w in rowwords:
            x0,y0,x1,y1,text,*_=w
            cx=(x0+x1)/2
            # assign to x band; ignore text outside table
            j=None
            for k in range(16):
                if xs[k]-1<=cx<=xs[k+1]+1:
                    j=k;break
            if j is not None:
                cells[j]=(cells[j]+' '+text).strip()
        cells=[clean(c) for c in cells]
        if is_civ(cells[1]) or (cells[0].isdigit() and cells[0]):
            # CIV cells must not contain names; salvage first numeric token if necessary
            for j in [1,2]:
                m=re.match(r'(\d{5,10})',cells[j])
                cells[j]=m.group(1) if m else (cells[j] if cells[j] in ['N/A','-'] else cells[j])
            out.append(cells)
    return out

def _source_rows_from_text(data):
    """Build a source-row catalog using robust 5-digit No. anchors in raw page text.

    COI record numbers are 5 digits in these reports, while CIVs are usually 7-8+
    digits and radicados are much longer. We therefore locate every 5-digit token
    and use it as a candidate row anchor, then attach the following text until the
    next candidate anchor. This is deliberately independent of table x-coordinates.
    """
    doc=fitz.open(stream=data,filetype='pdf')
    catalog=[]
    no_re=re.compile(r'(?<!\d)(\d{5})(?!\d)')
    for pi,page in enumerate(doc):
        text=page.get_text('text') or ''
        if not text.strip():
            continue
        matches=list(no_re.finditer(text))
        # Candidate No. values in the COI are typically 20xxx-29xxx; keep a broad
        # range to tolerate future reports, while rejecting obvious years/dates.
        anchors=[]; seen=set()
        for m in matches:
            no=m.group(1)
            n=int(no)
            if 10000 <= n <= 99999 and no not in seen:
                seen.add(no); anchors.append((m.start(),m.end(),no))
        # A page can contain duplicate 5-digit references elsewhere (e.g. in obs).
        # Keep only anchors that look like the start of a record: after the number,
        # the next ~120 chars should contain a plausible CIV/N/A/address token or the
        # row's company/contract data. This reduces false anchors in observations.
        for j,(st,en,no) in enumerate(anchors):
            nxt=anchors[j+1][0] if j+1<len(anchors) else len(text)
            block=text[st:nxt]
            block_norm=norm(block)
            # Avoid anchors that are clearly buried in a long observation/radicado.
            head=norm(text[st:min(en+140,nxt)])
            if len(block_norm)<15:
                continue
            # Record rows start with the number followed by CIV/address or N/A.
            plausible=bool(re.search(r'\b(?:N/?A|\d{5,10}|AK|KR|CL|AC|AV|DG|TV|CARRERA|CALLE|AUTOPISTA|DIAGONAL|TRANSVERSAL)\b', head))
            if not plausible:
                continue
            catalog.append({'PÁGINA PDF':pi+1,'No.':no,'BLOQUE_NORM':block_norm})
    doc.close()
    return catalog

def _validation_row_catalog(data):
    return _source_rows_from_text(data)

def _pdf_row_catalog(data):
    return _source_rows_from_text(data)

def _contract_variants(contract):
    k=contract_canonical(contract)
    raw=norm(contract)
    variants={k,raw}
    # Canonical formats: YYYY-NNNN and NNNN-YYYY.
    for value in [k, raw]:
        m=re.match(r'(20\d{2})-(\d{3,5})$', value or '')
        if m:
            y,n=m.group(1),m.group(2)
            variants.update({norm(f'SDM-{n}-{y}'),norm(f'SDM-{n}- {y}'),norm(f'{n}-{y}'),norm(f'{y}-{n}'),norm(f'{n}{y}'),norm(f'{y}{n}')})
        m=re.match(r'(\d{3,5})-(20\d{2})$', value or '')
        if m:
            n,y=m.group(1),m.group(2)
            variants.update({norm(f'SDM-{n}-{y}'),norm(f'SDM-{n}- {y}'),norm(f'{n}-{y}'),norm(f'{y}-{n}'),norm(f'{n}{y}'),norm(f'{y}{n}')})
    return {v for v in variants if v}

def _row_matches_company_contract(block_norm, company_norm, contract_canon):
    if not block_norm or not company_norm or not contract_canon:
        return False
    if company_norm not in block_norm:
        return False
    positions=[]
    for v in _contract_variants(contract_canon):
        pos=block_norm.find(v)
        if pos>=0:
            positions.append(pos)
    if not positions:
        return False
    # Require company and contract to be in the same row block. Do not require
    # textual order because PDF reading order can put contract before contractor.
    return True

def source_validation(data, df):
    """Control referencial simple y robusto: cuenta menciones textuales en el PDF.

    No intenta reconstruir filas. El objetivo es comparar, de forma orientativa,
    cuántas veces aparece el contratista/contrato en el PDF frente a los registros
    extraídos. Las menciones pueden aparecer en OBSERVACIONES, encabezados o
    referencias a terceros.
    """
    doc=fitz.open(stream=data,filetype='pdf')
    page_texts=[p.get_text('text') or '' for p in doc]
    doc.close()
    grouped=(df.copy())
    # Asegurar columnas y tipos para que la validación no derribe la app.
    if 'CONTRATISTA' not in grouped.columns:
        grouped['CONTRATISTA']=''
    if 'CONTRATO CANÓNICO' not in grouped.columns:
        if 'No. CONTRATO' in grouped.columns:
            grouped['CONTRATO CANÓNICO']=grouped['No. CONTRATO'].map(contract_canonical)
        else:
            grouped['CONTRATO CANÓNICO']=''
    grouped['CONTRATISTA']=grouped['CONTRATISTA'].fillna('').map(clean)
    grouped['CONTRATO CANÓNICO']=grouped['CONTRATO CANÓNICO'].fillna('').map(clean)
    g=grouped.groupby(['CONTRATISTA','CONTRATO CANÓNICO'],dropna=False).size().reset_index(name='REGISTROS_EXTRAÍDOS')
    rows=[]
    full_text='\n'.join(page_texts)
    full_norm=norm(full_text)
    for _,r in g.iterrows():
        company=clean(r['CONTRATISTA'])
        contract=clean(r['CONTRATO CANÓNICO'])
        c=norm(company)
        mentions_company=full_norm.count(c) if c else 0
        # Para el contrato usamos una búsqueda tolerante de sus variantes.
        mentions_contract=0
        try:
            for v in _contract_variants(contract):
                mentions_contract=max(mentions_contract,full_norm.count(norm(v)))
        except Exception:
            mentions_contract=0
        extracted=int(r['REGISTROS_EXTRAÍDOS'])
        diff=mentions_company-extracted
        rows.append({
            'CONTRATISTA':company,
            'CONTRATO CANÓNICO':contract,
            'MENCIONES CONTRATISTA EN PDF':mentions_company,
            'MENCIONES CONTRATO EN PDF':mentions_contract,
            'REGISTROS EXTRAÍDOS':extracted,
            'DIFERENCIA REFERENCIAL':diff,
            'VALIDACIÓN':'⚠️ REVISAR' if diff!=0 else '✅ OK',
        })
    if not rows:
        return pd.DataFrame(columns=['CONTRATISTA','CONTRATO CANÓNICO','MENCIONES CONTRATISTA EN PDF','MENCIONES CONTRATO EN PDF','REGISTROS EXTRAÍDOS','DIFERENCIA REFERENCIAL','VALIDACIÓN'])
    return pd.DataFrame(rows).sort_values(['VALIDACIÓN','REGISTROS EXTRAÍDOS'],ascending=[True,False]).reset_index(drop=True)

def extract_pdf(data, progress=None):
    doc=fitz.open(stream=data,filetype='pdf')
    records=[]; page_diag=[]; total=len(doc)
    for pi,page in enumerate(doc):
        rows=extract_page(page)
        txt=page.get_text('text').upper()
        section='SECCIÓN 1' if 'SECCIÓN 1.' in txt else ('SECCIÓN 2' if 'SECCIÓN 2.' in txt else '')
        page_diag.append((pi+1,len(rows)))
        for r in rows:
            r=r[:16]+['']*max(0,16-len(r))
            rec=dict(zip(COLS,r[:16]))
            rec['PÁGINA PDF']=pi+1
            rec['SECCIÓN']=section
            rec['CONTRATO CANÓNICO']=contract_canonical(rec['No. CONTRATO'])
            auth=norm(rec['AUTORIZADO']); obs=norm(rec['OBSERVACIONES'])
            if 'EMERGENCIA' in auth or 'FORMALIZACIONDEEMERGENCIA' in auth:
                state='FORMALIZACIÓN DE EMERGENCIA'
            elif auth in ('NO','NOAUTORIZADO','NOAUTORIZADA'):
                state='NO AUTORIZADO'
            elif auth in ('SI','SIAUTORIZADO','AUTORIZADO','AUTORIZADOCONOBSERVACIONES') or 'AUTORIZA' in obs:
                state='AUTORIZADO'
            else:
                state=clean(rec['AUTORIZADO']) or 'OTRO'
            rec['ESTADO INTERPRETADO']=state
            records.append(rec)
        if progress is not None:
            pct=int(((pi+1)/max(total,1))*100)
            progress.progress(pct, text=f'Analizando página {pi+1:,} de {total:,} · {len(records):,} registros encontrados')
    doc.close()
    return pd.DataFrame(records),page_diag

def make_excel(df, validation=None):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,sheet_name='COI completo')
        summary=pd.DataFrame([{'Indicador':'Registros','Valor':len(df)},{'Indicador':'Autorizados','Valor':int((df['ESTADO INTERPRETADO']=='AUTORIZADO').sum())},{'Indicador':'No autorizados','Valor':int((df['ESTADO INTERPRETADO']=='NO AUTORIZADO').sum())},{'Indicador':'Emergencias','Valor':int((df['ESTADO INTERPRETADO']=='FORMALIZACIÓN DE EMERGENCIA').sum())},{'Indicador':'24 HORAS','Valor':int(df['HORARIO DE TRABAJO'].map(norm).eq('24HORAS').sum())},{'Indicador':'Empresas','Valor':df['CONTRATISTA'].nunique()},{'Indicador':'Contratos','Valor':df['CONTRATO CANÓNICO'].nunique()}])
        summary.to_excel(writer,index=False,sheet_name='Resumen')
        emp=df.groupby('CONTRATISTA',dropna=False).agg(REGISTROS=('No.','count'),AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='AUTORIZADO').sum()),NO_AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='NO AUTORIZADO').sum()),EMERGENCIAS=('ESTADO INTERPRETADO',lambda s:(s=='FORMALIZACIÓN DE EMERGENCIA').sum()),CONTRATOS=('CONTRATO CANÓNICO','nunique')).reset_index().sort_values('REGISTROS',ascending=False)
        emp.to_excel(writer,index=False,sheet_name='Empresas')
        con=df.groupby(['CONTRATO CANÓNICO','No. CONTRATO','CONTRATISTA'],dropna=False).agg(REGISTROS=('No.','count'),AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='AUTORIZADO').sum()),NO_AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='NO AUTORIZADO').sum()),EMERGENCIAS=('ESTADO INTERPRETADO',lambda s:(s=='FORMALIZACIÓN DE EMERGENCIA').sum())).reset_index().sort_values('REGISTROS',ascending=False)
        con.to_excel(writer,index=False,sheet_name='Contratos')
        hor=df.groupby('HORARIO DE TRABAJO',dropna=False).agg(REGISTROS=('No.','count'),AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='AUTORIZADO').sum()),NO_AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='NO AUTORIZADO').sum()),EMERGENCIAS=('ESTADO INTERPRETADO',lambda s:(s=='FORMALIZACIÓN DE EMERGENCIA').sum())).reset_index().sort_values('REGISTROS',ascending=False)
        hor.to_excel(writer,index=False,sheet_name='Horarios')
        loc=df.groupby('LOCALIDAD',dropna=False).size().reset_index(name='REGISTROS').sort_values('REGISTROS',ascending=False); loc.to_excel(writer,index=False,sheet_name='Localidades')
        if validation is not None:
            validation.to_excel(writer,index=False,sheet_name='Validación extracción')
    bio.seek(0)
    wb=load_workbook(bio); 
    for ws in wb.worksheets:
        ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
        for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='17365D'); c.alignment=Alignment(horizontal='center',vertical='center')
        for col in ws.columns:
            letter=get_column_letter(col[0].column); ws.column_dimensions[letter].width=min(max(max(len(str(c.value or '')) for c in col)+2,10),55)
        if ws.title=='COI completo': ws.column_dimensions['L'].width=70
    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

def pdf_pages(data,pages):
    src=fitz.open(stream=data,filetype='pdf'); out=fitz.open();
    for p in sorted(set(pages)): out.insert_pdf(src,from_page=p-1,to_page=p-1)
    b=out.tobytes(); out.close(); src.close(); return b

# ---------- UI ----------
st.markdown(f'''<div class="hero"><h1>📊 {APP_NAME} · Secretaría Distrital de Movilidad</h1><p>Analiza el COI completo, valida la extracción y genera listas por empresas, contratos, estados y horarios.</p><div class="small">Versión {APP_VERSION} · extracción por celdas dinámicas del formato oficial COI</div></div>''',unsafe_allow_html=True)

with st.sidebar:
    st.markdown('### ⚙️ Flujo de trabajo')
    st.write('1️⃣ Carga el PDF')
    st.write('2️⃣ Analiza el COI')
    st.write('3️⃣ Filtra empresas y contratos')
    st.write('4️⃣ Genera la lista')
    st.divider(); st.markdown('<div class="small">Monitor COI · SDM Bogotá<br>Versión 12.0</div>',unsafe_allow_html=True)

uploaded=st.file_uploader('📥 CARGA 1 · Selecciona el COI oficial en PDF',type=['pdf'])
if uploaded:
    data=uploaded.getvalue(); st.caption(f'Archivo: **{uploaded.name}** · {len(data)/1024/1024:.1f} MB')
    if st.button('🔎 ANALIZAR COI',type='primary',use_container_width=True):
        st.session_state.pop('filtered',None)
        progress=st.progress(0, text='Preparando análisis del COI...')
        status=st.empty()
        status.info('⏳ Iniciando extracción página por página...')
        df,diag=extract_pdf(data, progress)
        progress.progress(100, text=f'✅ Análisis terminado · {len(df):,} registros extraídos')
        status.success('✅ Extracción terminada. Ahora se cuentan las menciones de empresas y contratos en el PDF...')
        try:
            validation=source_validation(data, df)
            status.success('✅ Control de menciones terminado.')
        except Exception as exc:
            validation=pd.DataFrame()
            status.warning(f'⚠️ La extracción terminó, pero la validación no pudo completarse en esta versión. Los registros extraídos se conservan. Detalle técnico: {type(exc).__name__}')
        st.session_state['coi_df']=df; st.session_state['coi_data']=data; st.session_state['coi_name']=uploaded.name; st.session_state['diag']=diag; st.session_state['validation']=validation
        st.session_state['generated']=True
        st.rerun()

if 'coi_df' not in st.session_state: st.info('Carga el PDF y pulsa **ANALIZAR COI**. La aplicación no procesará el documento antes de ese botón.')
else:
    df=st.session_state['coi_df'].copy(); data=st.session_state['coi_data']
    st.success(f"✅ {st.session_state['coi_name']} · **{len(df):,} registros extraídos** · {sum(n>0 for _,n in st.session_state['diag']):,} páginas con datos")
    validation=st.session_state.get('validation',pd.DataFrame())
    with st.expander('🔎 Control de menciones en PDF · referencia para revisión', expanded=True):
        if not validation.empty:
            ok=int((validation['VALIDACIÓN']=='✅ OK').sum())
            rev=int((validation['VALIDACIÓN']=='⚠️ REVISAR').sum())
            m1,m2,m3=st.columns(3)
            m1.metric('Empresas/contratos revisados',f'{len(validation):,}')
            m2.metric('Sin diferencia referencial',f'{ok:,}')
            m3.metric('Para revisión manual',f'{rev:,}')
            st.dataframe(validation,use_container_width=True,hide_index=True)
            if rev:
                st.warning('Este cuadro cuenta menciones textuales en todo el PDF, incluidas posibles menciones en OBSERVACIONES y encabezados. No reemplaza la revisión fila por fila; úsalo como indicador para orientar la revisión.')
            else:
                st.success('No se detectaron diferencias entre menciones textuales y registros extraídos en los grupos revisados.')
        else:
            st.info('No se pudo construir el control de menciones para este archivo.')
    if len(df):
        companies=sorted([x for x in df['CONTRATISTA'].dropna().unique() if clean(x)],key=lambda x:x.upper())
        contracts=sorted([x for x in df['CONTRATO CANÓNICO'].dropna().unique() if clean(x)])
        states=sorted(df['ESTADO INTERPRETADO'].dropna().unique())
        localities=sorted([x for x in df['LOCALIDAD'].dropna().unique() if clean(x)])
        sections=sorted([x for x in df['SECCIÓN'].dropna().unique() if clean(x)])
        st.markdown('### 🎛️ CARGA 2 · Selecciona lo que quieres incluir en la lista')
        c1,c2=st.columns(2)
        with c1: sel_comp=st.multiselect('🏢 Empresas / contratistas',companies,placeholder='Selecciona una o varias')
        with c2: sel_con=st.multiselect('📄 No. CONTRATO',contracts,placeholder='Selecciona uno o varios')
        c3,c4,c5=st.columns(3)
        with c3: sel_state=st.multiselect('Estado',states,default=states)
        with c4: sel_loc=st.multiselect('Localidad',localities)
        with c5: sel_sec=st.multiselect('Sección',sections)
        q=st.text_input('🔎 Búsqueda libre','',placeholder='CIV, dirección, contrato, radicado, ingeniero, observación...')
        mask=pd.Series(True,index=df.index)
        if sel_comp: mask &= df['CONTRATISTA'].isin(sel_comp)
        if sel_con: mask &= df['CONTRATO CANÓNICO'].isin(sel_con)
        if sel_state: mask &= df['ESTADO INTERPRETADO'].isin(sel_state)
        if sel_loc: mask &= df['LOCALIDAD'].isin(sel_loc)
        if sel_sec: mask &= df['SECCIÓN'].isin(sel_sec)
        if q:
            qq=norm(q); mask &= df[COLS].astype(str).apply(lambda col: col.map(norm).str.contains(qq,regex=False)).any(axis=1)
        candidate=df[mask].copy()
        if st.button('⚡ GENERAR LISTA SELECCIONADA',type='primary',use_container_width=True): st.session_state['filtered']=candidate
        result=st.session_state.get('filtered',candidate)
        if not result.empty:
            a,b,c,d,e,f,g=st.columns(7)
            vals=[len(result),int((result['ESTADO INTERPRETADO']=='AUTORIZADO').sum()),int((result['ESTADO INTERPRETADO']=='NO AUTORIZADO').sum()),int((result['ESTADO INTERPRETADO']=='FORMALIZACIÓN DE EMERGENCIA').sum()),int(result['HORARIO DE TRABAJO'].map(norm).eq('24HORAS').sum()),result['CONTRATISTA'].nunique(),result['CONTRATO CANÓNICO'].nunique()]
            labs=['📄 Registros','🟢 Autorizados','🔴 No autorizados','🟠 Emergencias','🕘 24 HORAS','🏢 Empresas','📑 Contratos']
            for col,lab,val in zip([a,b,c,d,e,f,g],labs,vals): col.metric(lab,f'{val:,}')
            st.success(f"Lista generada: **{len(result):,} registros** · {result['CONTRATISTA'].nunique()} empresas · {result['CONTRATO CANÓNICO'].nunique()} contratos")
            tab1,tab2,tab3,tab4,tab5=st.tabs(['📊 Dashboard','📋 Lista COI','🔍 Trazabilidad','📥 Excel y PDF','🤖 Automático'])
            with tab1:
                st.subheader('Resumen del resultado seleccionado')
                r1,r2=st.columns(2)
                state_counts=result['ESTADO INTERPRETADO'].value_counts().reset_index(); state_counts.columns=['Estado','Cantidad']
                with r1: st.plotly_chart(px.pie(state_counts,names='Estado',values='Cantidad',hole=.45,title='Estado de los registros'),use_container_width=True)
                comp_counts=result['CONTRATISTA'].value_counts().head(20).reset_index(); comp_counts.columns=['Contratista','Cantidad']
                with r2: st.plotly_chart(px.bar(comp_counts,y='Contratista',x='Cantidad',orientation='h',title='Registros por empresa'),use_container_width=True)
                hcounts=result['HORARIO DE TRABAJO'].value_counts().reset_index(); hcounts.columns=['Horario','Cantidad']
                st.plotly_chart(px.bar(hcounts,x='Horario',y='Cantidad',title='Registros por horario'),use_container_width=True)
                cross=pd.crosstab(result['HORARIO DE TRABAJO'],result['ESTADO INTERPRETADO']).reset_index()
                st.plotly_chart(px.bar(cross,x='HORARIO DE TRABAJO',y=[c for c in cross.columns if c!='HORARIO DE TRABAJO'],title='Horario vs. estado'),use_container_width=True)
            with tab2:
                st.dataframe(result[COLS+['PÁGINA PDF','SECCIÓN','CONTRATO CANÓNICO','ESTADO INTERPRETADO']],use_container_width=True,height=520,hide_index=True)
            with tab3:
                nums=result['No.'].astype(str).tolist(); chosen=st.selectbox('Selecciona un No. del COI',nums)
                rr=result[result['No.'].astype(str)==chosen].iloc[0]
                st.write({k:rr[k] for k in ['No.','CIV INICIO','CIV FIN','DIRECCIÓN DE LA OBRA INICIO','DIRECCIÓN DE LA OBRA FIN','CONTRATISTA','No. CONTRATO','HORARIO DE TRABAJO','HORARIO DE CIERRE','AUTORIZADO','LOCALIDAD','PÁGINA PDF']})
                st.download_button('📄 Descargar página original del PDF',pdf_pages(data,[int(rr['PÁGINA PDF'])]),file_name=f'COI_pagina_{int(rr["PÁGINA PDF"])}.pdf',mime='application/pdf')
            with tab4:
                st.download_button('📊 Descargar Excel completo filtrado',make_excel(result, source_validation(data, result)),file_name='Monitor_COI_resultado.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                pages=pdf_pages(data,result['PÁGINA PDF'].astype(int).tolist()); st.download_button('📄 Descargar PDF con páginas originales',pages,file_name='Monitor_COI_paginas_seleccionadas.pdf',mime='application/pdf')
                hor=result.groupby('HORARIO DE TRABAJO').agg(REGISTROS=('No.','count'),AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='AUTORIZADO').sum()),NO_AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='NO AUTORIZADO').sum()),EMERGENCIAS=('ESTADO INTERPRETADO',lambda s:(s=='FORMALIZACIÓN DE EMERGENCIA').sum())).reset_index().sort_values('REGISTROS',ascending=False)
                st.dataframe(hor,use_container_width=True,hide_index=True)
            with tab5:
                st.subheader('🤖 Resultado automático semanal')
                st.caption('El proceso automático busca el COI más reciente directamente en la página oficial de SDM, lo procesa cada viernes y envía por Gmail el PDF original, el Excel completo y un PDF resumen. No compara con el COI anterior.')
                status_url='https://raw.githubusercontent.com/csenalizarbogota2025-svg/monitor-coi/main/automatic/status.json'
                try:
                    rr=requests.get(status_url,timeout=8)
                    if rr.ok:
                        js=rr.json(); a,b,c,d=st.columns(4)
                        a.metric('Último COI procesado',str(js.get('coi_number','—')))
                        b.metric('Registros',f"{int(js.get('records_extracted',0)):,}")
                        c.metric('Empresas',f"{int(js.get('companies',0)):,}")
                        d.metric('Contratos',f"{int(js.get('contracts',0)):,}")
                        st.success(f"✅ COI No. {js.get('coi_number','—')} · fecha {js.get('coi_date','—')} · procesado {js.get('processed_at_utc','—')}")
                        st.write(f"Fuente oficial: {js.get('source_url','')}")
                        st.write(f"24 HORAS: {int(js.get('24_hours',0)):,} · Autorizados: {int(js.get('authorized',0)):,} · No autorizados: {int(js.get('not_authorized',0)):,} · Emergencias: {int(js.get('emergencies',0)):,}")
                    else:
                        st.info('Aún no existe un resultado automático publicado en el historial.')
                except Exception:
                    st.info('El historial automático aparecerá aquí después de la primera ejecución de GitHub Actions.')
                st.markdown('**Programación:** viernes · ejecución automática en GitHub Actions. También existe un botón de ejecución manual para hacer pruebas sin esperar al viernes.')
        else: st.warning('No hay registros que cumplan los filtros seleccionados.')
