import os, re, json, ssl, smtplib, html, sys
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage
from email.utils import formatdate
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import fitz
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.enums import TA_LEFT

from engine import extract_pdf, source_validation, make_excel

PMT_URL = 'https://www.movilidadbogota.gov.co/pmt'
USER_AGENT = 'MonitorCOI/12.0 (+https://github.com/csenalizarbogota2025-svg/monitor-coi)'
TIMEOUT = 60
ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'automatic' / 'latest'
OUT.mkdir(parents=True, exist_ok=True)


def parse_date_from_label(s):
    m = re.search(r'(20\d{2})[_-](\d{2})[_-](\d{2})', s)
    if not m: return None
    return datetime(int(m.group(1)),int(m.group(2)),int(m.group(3)))


def find_latest_coi():
    r=requests.get(PMT_URL,headers={'User-Agent':USER_AGENT},timeout=TIMEOUT)
    r.raise_for_status()
    soup=BeautifulSoup(r.text,'html.parser')
    candidates=[]
    for a in soup.find_all('a',href=True):
        text=' '.join(a.stripped_strings)
        href=html.unescape(urljoin(PMT_URL,a['href']))
        probe=f'{text} {href}'
        if 'COI' not in probe.upper():
            continue
        if any(x in probe.upper() for x in ['SIN REVISI','_AD_','_COR_']):
            continue
        m=re.search(r'COI[_-]?(\d+)[^0-9]*(20\d{2})[_-](\d{2})[_-](\d{2})',probe,re.I)
        if not m:
            continue
        num=int(m.group(1)); dt=datetime(int(m.group(2)),int(m.group(3)),int(m.group(4)))
        candidates.append((dt,num,href,text))
    if not candidates:
        raise RuntimeError('No se encontró un enlace COI válido en la página oficial de SDM.')
    # Prefer the newest document date, then highest COI number.
    return max(candidates,key=lambda x:(x[0],x[1]))


def download_and_verify(meta):
    dt,num,url,label=meta
    r=requests.get(url,headers={'User-Agent':USER_AGENT},timeout=120)
    r.raise_for_status()
    data=r.content
    doc=fitz.open(stream=data,filetype='pdf')
    sample='\n'.join((doc.load_page(i).get_text('text') or '') for i in range(min(3,len(doc))))
    pages=len(doc); doc.close()
    if not re.search(rf'COI\s*No\.?\s*{num}\b',sample,re.I):
        raise RuntimeError(f'El PDF descargado no valida como COI No. {num}.')
    expected=dt.strftime('%B %d, %Y')
    # More permissive date validation: look for numeric month/day/year or Spanish month.
    months={'01':'enero','02':'febrero','03':'marzo','04':'abril','05':'mayo','06':'junio','07':'julio','08':'agosto','09':'septiembre','10':'octubre','11':'noviembre','12':'diciembre'}
    month_name=months[dt.strftime('%m')]
    if month_name not in sample.lower() and dt.strftime('%d') not in sample:
        raise RuntimeError(f'La fecha interna del PDF no parece corresponder al COI esperado ({dt:%Y-%m-%d}).')
    return data,pages


def build_summary_pdf(df, meta, validation, outpath):
    dt,num,url,label=meta
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Small',parent=styles['BodyText'],fontSize=8,leading=10))
    doc=SimpleDocTemplate(str(outpath),pagesize=landscape(letter),rightMargin=28,leftMargin=28,topMargin=28,bottomMargin=28)
    story=[]
    story.append(Paragraph('Monitor COI · Informe automático semanal',styles['Title']))
    story.append(Paragraph(f'<b>COI No. {num}</b> · {dt:%d/%m/%Y}',styles['Heading2']))
    story.append(Paragraph(f'Fuente oficial: {html.escape(url)}',styles['Small']))
    story.append(Spacer(1,10))
    indicators=[
        ['Indicador','Valor'],
        ['Registros extraídos',f'{len(df):,}'],
        ['Empresas/contratistas',f'{df["CONTRATISTA"].nunique():,}'],
        ['Contratos',f'{df["CONTRATO CANÓNICO"].nunique():,}'],
        ['Autorizados',f'{int((df["ESTADO INTERPRETADO"]=="AUTORIZADO").sum()):,}'],
        ['No autorizados',f'{int((df["ESTADO INTERPRETADO"]=="NO AUTORIZADO").sum()):,}'],
        ['Formalizaciones de emergencia',f'{int((df["ESTADO INTERPRETADO"]=="FORMALIZACIÓN DE EMERGENCIA").sum()):,}'],
        ['24 HORAS',f'{int(df["HORARIO DE TRABAJO"].map(lambda x:str(x).upper().replace(" ","")).eq("24HORAS").sum()):,}'],
    ]
    t=Table(indicators,colWidths=[210,100])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.25,colors.grey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t); story.append(Spacer(1,12))

    emp=df.groupby('CONTRATISTA').agg(REGISTROS=('No.','count'),AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='AUTORIZADO').sum()),NO_AUTORIZADOS=('ESTADO INTERPRETADO',lambda s:(s=='NO AUTORIZADO').sum()),EMERGENCIAS=('ESTADO INTERPRETADO',lambda s:(s=='FORMALIZACIÓN DE EMERGENCIA').sum())).reset_index().sort_values('REGISTROS',ascending=False).head(20)
    data=[['Contratista','Registros','Autorizados','No autorizados','Emergencias']]+[[str(r['CONTRATISTA'])[:65],str(int(r['REGISTROS'])),str(int(r['AUTORIZADOS'])),str(int(r['NO_AUTORIZADOS'])),str(int(r['EMERGENCIAS']))] for _,r in emp.iterrows()]
    te=Table(data,colWidths=[300,70,80,85,85],repeatRows=1)
    te.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),7)]))
    story.append(Paragraph('Principales contratistas por número de registros',styles['Heading3']))
    story.append(te); story.append(Spacer(1,10))

    hor=df.groupby('HORARIO DE TRABAJO').size().reset_index(name='REGISTROS').sort_values('REGISTROS',ascending=False).head(25)
    dh=[['Horario','Registros']]+[[str(r['HORARIO DE TRABAJO'])[:70],str(int(r['REGISTROS']))] for _,r in hor.iterrows()]
    th=Table(dh,colWidths=[420,90],repeatRows=1)
    th.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),7)]))
    story.append(Paragraph('Horarios de trabajo',styles['Heading3'])); story.append(th)

    if validation is not None and not validation.empty:
        story.append(PageBreak()); story.append(Paragraph('Control de menciones en PDF · referencia para revisión',styles['Heading2']))
        story.append(Paragraph('Las menciones cuentan apariciones textuales y pueden incluir OBSERVACIONES, encabezados u otras referencias. No equivalen necesariamente al número de registros.',styles['Small']))
        val=validation.head(40)
        dv=[['Contratista','Contrato','Menciones','Extraídos','Dif.']]+[[str(r['CONTRATISTA'])[:55],str(r['CONTRATO CANÓNICO']),str(int(r['MENCIONES CONTRATISTA EN PDF'])),str(int(r['REGISTROS EXTRAÍDOS'])),str(int(r['DIFERENCIA REFERENCIAL']))] for _,r in val.iterrows()]
        tv=Table(dv,colWidths=[300,90,80,70,60],repeatRows=1)
        tv.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),7)]))
        story.append(tv)
    doc.build(story)


def send_email(pdf_path, xlsx_path, summary_path, meta, df):
    user=os.environ.get('GMAIL_USER','').strip(); app_pw=os.environ.get('GMAIL_APP_PASSWORD','').strip(); recipients=[x.strip() for x in os.environ.get('MAIL_TO','').split(',') if x.strip()]; cc=[x.strip() for x in os.environ.get('MAIL_CC','').split(',') if x.strip()]
    if not user or not app_pw or not recipients:
        raise RuntimeError('Faltan secretos GMAIL_USER, GMAIL_APP_PASSWORD o MAIL_TO.')
    dt,num,url,label=meta
    msg=EmailMessage()
    msg['From']=user; msg['To']=', '.join(recipients)
    if cc: msg['Cc']=', '.join(cc)
    msg['Date']=formatdate(localtime=True)
    msg['Subject']=f'Monitor COI · COI No. {num} · {dt:%d/%m/%Y}'
    body=(f'Se procesó automáticamente el COI No. {num} publicado por la Secretaría Distrital de Movilidad.\n\n'
          f'Fecha del COI: {dt:%d/%m/%Y}\nRegistros extraídos: {len(df):,}\nEmpresas/contratistas: {df["CONTRATISTA"].nunique():,}\nContratos: {df["CONTRATO CANÓNICO"].nunique():,}\n\n'
          f'Fuente oficial: {url}\n\nAdjuntos:\n- PDF original del COI\n- Excel completo procesado\n- PDF resumen del análisis\n')
    msg.set_content(body)
    for p,ctype,subtype in [(pdf_path,'application','pdf'),(xlsx_path,'application','vnd.openxmlformats-officedocument.spreadsheetml.sheet'),(summary_path,'application','pdf')]:
        payload=Path(p).read_bytes(); msg.add_attachment(payload,maintype=ctype.split('/')[0],subtype=ctype.split('/')[1],filename=Path(p).name)
    context=ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context,timeout=60) as smtp:
        smtp.login(user,app_pw); smtp.send_message(msg)


def main():
    meta=find_latest_coi()
    data,pages=download_and_verify(meta)
    dt,num,url,label=meta
    pdf_path=OUT/f'COI_{num}_{dt:%Y_%m_%d}.pdf'; pdf_path.write_bytes(data)
    df,diag=extract_pdf(data)
    validation=source_validation(data,df)
    xlsx_path=OUT/f'Monitor_COI_{num}_{dt:%Y_%m_%d}.xlsx'; xlsx_path.write_bytes(make_excel(df,validation))
    summary_path=OUT/f'Resumen_COI_{num}_{dt:%Y_%m_%d}.pdf'; build_summary_pdf(df,meta,validation,summary_path)
    send_email(pdf_path,xlsx_path,summary_path,meta,df)
    status={
      'ok':True,'processed_at_utc':datetime.utcnow().isoformat(timespec='seconds')+'Z','coi_number':num,'coi_date':dt.strftime('%Y-%m-%d'),
      'source_url':url,'source_label':label,'pdf_pages':pages,'records_extracted':int(len(df)),'companies':int(df['CONTRATISTA'].nunique()),'contracts':int(df['CONTRATO CANÓNICO'].nunique()),
      'authorized':int((df['ESTADO INTERPRETADO']=='AUTORIZADO').sum()),'not_authorized':int((df['ESTADO INTERPRETADO']=='NO AUTORIZADO').sum()),'emergencies':int((df['ESTADO INTERPRETADO']=='FORMALIZACIÓN DE EMERGENCIA').sum()),
      '24_hours':int(df['HORARIO DE TRABAJO'].map(lambda x:str(x).upper().replace(' ','')).eq('24HORAS').sum()),'email_to':os.environ.get('MAIL_TO','')
    }
    status_path=ROOT/'automatic'/'status.json'; status_path.write_text(json.dumps(status,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(status,ensure_ascii=False))

if __name__=='__main__':
    main()
