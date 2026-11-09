#!/usr/bin/env python3
"""Santander Reportos: python extractor_santander.py CARPETA [--salida CARPETA]"""
import argparse, csv, json, re, sys
from pathlib import Path
from decimal import Decimal
from datetime import datetime

DATE = r'\d{2}-[A-Za-z]{3}-\d{4}'
NUM = r'-?[\d,]+(?:\.\d+)?'
ROW = re.compile(rf'^\s*(\d{{6,10}})\s+({DATE})\s+(COMPRA|VENTA)\s+(REPORTO|\w+)\s+(?:(\S+)\s+)?([A-Z]{{1,4}})\s+({NUM})\s+\$\s*({NUM})\s+\$\s*({NUM})\s+\$\s*({NUM})\s+({NUM})\s+(\d+)\s+({NUM})\s+\$\s*({NUM})\s+({DATE})\s+([A-Z]{{3}})\s+({DATE})\s*$',re.I)
TOTAL = re.compile(rf'Totales:\s*(\d+)\s+Titulos\s+({NUM})\s+Monto\s+\$\s*({NUM})\s+Solicitado\s+({NUM})\s+Diferencia\s+({NUM})',re.I)
FIELDS = ['archivo','pagina','contrato','cliente','numero_operacion','fecha_operacion','compra_venta','tipo_operacion','emisora','serie','tipo_valor','titulos','precio_limpio','precio_sucio','monto_neto','tasa','plazo_dias','premio','importe_vencimiento','fecha_vencimiento','moneda','cuenta_indeval','fecha_liquidacion','tipo_ejecucion','total_operaciones_bloque','total_titulos_bloque','total_monto_bloque','solicitado','diferencia','validacion','observaciones']
MONEY = ['precio_limpio','precio_sucio','monto_neto','tasa','premio','importe_vencimiento','total_monto_bloque','solicitado','diferencia']

def n(s): return str(Decimal(s.replace(',','')))
def fecha(s): return datetime.strptime(s.title(),'%d-%b-%Y').strftime('%Y-%m-%d')
def parse_page(text,filename,page):
    lines = text.splitlines(); head='\n'.join(lines[:15]); cm=re.search(r'CONTRATO\s+(\d+)\s+(.+)',head)
    contrato=cm.group(1) if cm else ''; cliente=cm.group(2).strip() if cm else ''
    if 'CONFIRMACI' not in head.upper() or 'REPORTOS' not in head.upper(): return [],['Formato no identificado']
    results=[]; issues=[]
    for i,line in enumerate(lines):
        # Extract text often puts issuer and series on adjacent lines above/below the main row.
        m=ROW.match(line)
        if not m: continue
        g=m.groups(); previous=lines[i-1].strip() if i else ''
        following=lines[i+1].strip() if i+1<len(lines) else ''
        issuer=(g[4] or '').strip(); serie=''
        # Santander PDFs position the issuer on a separate line just above and the series below.
        if not issuer and re.match(r'^[A-Z][A-Z0-9]{2,20}$',previous): issuer=previous
        if re.fullmatch(r'\d{5,8}',following): serie=following
        if not issuer:
            # Sometimes the previous line also contains Indeval text; use the leftmost word.
            pre=re.match(r'^\s*([A-Z][A-Z0-9]{2,20})\b',previous)
            if pre: issuer=pre.group(1)
        if not issuer: issues.append(f'{filename} página {page}: emisora no encontrada para {g[0]}')
        rec=dict(zip(FIELDS,[None]*len(FIELDS)))
        rec.update(archivo=filename,pagina=page,contrato=contrato,cliente=cliente,numero_operacion=g[0],fecha_operacion=fecha(g[1]),compra_venta=g[2],tipo_operacion=g[3],emisora=issuer,serie=serie,tipo_valor=g[5],titulos=int(g[6].replace(',','')),precio_limpio=n(g[7]),precio_sucio=n(g[8]),monto_neto=n(g[9]),tasa=n(g[10]),plazo_dias=int(g[11]),premio=n(g[12]),importe_vencimiento=n(g[13]),fecha_vencimiento=fecha(g[14]),moneda=g[15],fecha_liquidacion=fecha(g[16]),cuenta_indeval='LIQUIDACION AL CONTRATO',tipo_ejecucion='Cliente Institucional')
        # Totales are one line after series; search until next operation or 6 lines.
        for look in lines[i+1:i+7]:
            mt=TOTAL.search(look)
            if mt:
                rec.update(total_operaciones_bloque=int(mt[1]),total_titulos_bloque=int(mt[2].replace(',','')),total_monto_bloque=n(mt[3]),solicitado=n(mt[4]),diferencia=n(mt[5]));break
        errors=[]
        if Decimal(rec['monto_neto'])+Decimal(rec['premio']) != Decimal(rec['importe_vencimiento']): errors.append('Monto neto + premio != vencimiento')
        if rec['total_monto_bloque'] and Decimal(rec['total_monto_bloque'])!=Decimal(rec['monto_neto']): errors.append('Monto bloque distinto')
        if rec['solicitado'] and rec['diferencia'] and Decimal(rec['solicitado'])-Decimal(rec['monto_neto'])!=Decimal(rec['diferencia']): errors.append('Diferencia no coincide')
        if not rec['solicitado']: errors.append('Totales no encontrados')
        if not rec['emisora'] or not rec['serie']: errors.append('Emisora o serie incompleta')
        rec['validacion']='REVISAR' if errors else 'OK';rec['observaciones']='; '.join(errors)
        results.append(rec)
    if not results: issues.append(f'{filename} página {page}: sin operaciones detectadas')
    return results,issues

def xlsx(rows, path):
    from artifact_tool import Workbook, SpreadsheetFile
    wb=Workbook.create(); sheet=wb.worksheets.add('Operaciones')
    data=[FIELDS]+[[float(r[k]) if r[k] in MONEY and r[k] is not None else r[k] for k in FIELDS] for r in rows]
    sheet.get_range_by_indexes(0,0,len(data),len(FIELDS)).values=data
    sheet.get_range_by_indexes(0,0,1,len(FIELDS)).format={'fill':'#B91C1C','font':{'bold':True,'color':'#FFFFFF'},'row_height':28}
    sheet.freeze_panes.freeze_rows(1)
    sheet.get_range('A:AE').format.column_width=17
    sheet.get_range('D:D').format.column_width=35
    sheet.get_range('AE:AE').format.column_width=45
    if rows:
        sheet.tables.add(f'A1:AE{len(rows)+1}',True,'TablaReportos')
    SpreadsheetFile.export_xlsx(wb).save(str(path))

def main():
    ap=argparse.ArgumentParser(description='Extrae reportos Santander de todos los PDF de una carpeta y subcarpetas')
    ap.add_argument('carpeta');ap.add_argument('--salida');args=ap.parse_args()
    root=Path(args.carpeta).expanduser().resolve()
    if not root.is_dir(): print(json.dumps({'estado':'ERROR','mensaje':'Carpeta no existe','ruta':str(root)}));return 2
    output=Path(args.salida).expanduser().resolve() if args.salida else root/'Resultados_Santander'
    output.mkdir(parents=True,exist_ok=True)
    try: import fitz
    except ImportError: print('Instalar PyMuPDF: py -m pip install pymupdf',file=sys.stderr);return 2
    pdfs=[p for p in root.rglob('*.pdf') if output not in p.parents]
    rows=[];issues=[]; docs_ok=0
    for pdf in sorted(pdfs):
        try:
            with fitz.open(pdf) as doc:
                found=0
                for ix,page in enumerate(doc):
                    rr,ee=parse_page(page.get_text(sort=True),str(pdf.relative_to(root)),ix+1)
                    rows.extend(rr);issues.extend(ee);found+=len(rr)
                if found:docs_ok+=1
                else:issues.append(f'{pdf.name}: revisar PDF (puede ser escaneado; OCR no habilitado)')
        except Exception as exc: issues.append(f'{pdf.name}: {type(exc).__name__}: {exc}')
    csvpath=output/'operaciones_santander.csv'
    with csvpath.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.DictWriter(f,fieldnames=FIELDS);w.writeheader();w.writerows(rows)
    (output/'operaciones_santander.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
    (output/'incidencias.json').write_text(json.dumps(issues,ensure_ascii=False,indent=2),encoding='utf-8')
    try: xlsx(rows,output/'operaciones_santander.xlsx')
    except Exception as exc: issues.append('No se pudo crear Excel: '+str(exc))
    summary={'estado':'OK' if rows and not issues else 'REVISAR','pdfs':len(pdfs),'pdfs_con_operaciones':docs_ok,'operaciones':len(rows),'operaciones_revisar':sum(r['validacion']!='OK' for r in rows),'incidencias':issues,'salida':str(output)}
    (output/'resultado_ejecucion.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False));return 0 if rows and not issues and not summary['operaciones_revisar'] else 1
if __name__=='__main__':sys.exit(main())