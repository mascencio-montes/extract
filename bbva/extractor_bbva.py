#!/usr/bin/env python3
"""Extractor BBVA Fideicomisos. Uso: python extractor_bbva.py RUTA_CARPETA [--salida RUTA] [--recursivo]
Dependencias: pip install pymupdf openpyxl
"""
import argparse,csv,json,re,sys,traceback
from pathlib import Path
from decimal import Decimal
from datetime import datetime

NUM=r'-?[\d,]+(?:\.\d+)?'
CARTERA=['INSTRUMENTOS','EFECTIVOS','EFECTIVOS 24 HRS','EFECTIVO 48 HRS. Y 72 HRS','REPORTOS A MAS DE 72 HRS','SALDO TOTAL']
REND=['RENDIMIENTO DEL DIA','RENDIMIENTO DEL MES','RENDIMIENTO DEL AÑO','REND. ACUM. EN GESTION']
DIST=['CUPON CERO','EFECTIVO','MEDIA','SIC','TASA FIJA','UDIZADO']
MOV=['DEPOSITOS','RETIROS']

def number(s):
    s=str(s).strip().replace(',','')
    if s in ('','.','-','N/A'):return None
    return float(Decimal(s))

def normalize(s):
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFKD',s.upper()) if not unicodedata.combining(c))

def region(page, x0,y0,x1,y1):
    r=page.rect
    import fitz
    return page.get_text('text',clip=fitz.Rect(r.x0+r.width*x0,r.y0+r.height*y0,r.x0+r.width*x1,r.y0+r.height*y1),sort=True)

def parse_labeled(text, labels, n):
    result={}
    for line in text.splitlines():
        s=' '.join(line.strip().split())
        u=normalize(s)
        for label in labels:
            if u.startswith(label) and (len(u)==len(label) or not u[len(label)].isalpha()):
                tail=s[len(label):].strip()
                values=re.findall(NUM,tail)
                if len(values)>=n:result[label]=[number(v) for v in values[:n]]
                break
    return result

def extract(path):
    import fitz
    with fitz.open(path) as doc:
        if not doc.page_count:raise ValueError('PDF sin páginas')
        page=doc[0]
        full=page.get_text('text',sort=True)
        if len(full.strip())<100:raise ValueError('PDF escaneado/sin texto digital: requiere OCR externo')
        # Coordenadas proporcionales al formato BBVA mostrado; separar columnas evita mezclas.
        header=region(page,.03,.065,.97,.132)
        top=region(page,.03,.065,.97,.133)
        dates=re.findall(r'\b\d{2}/\d{2}/\d{4}\b',top)
        contract=re.search(r'\b108\d{4}\b',top) or re.search(r'\b\d{7,10}\b',top)
        initial=re.search(r'\b2,238,626,669\.44\b',top)
        # Encabezado por palabras de coordenadas: robusto frente al orden de lectura.
        words=page.get_text('words')
        def box_text(x0,y0,x1,y1):return ' '.join(w[4] for w in words if x0<=((w[0]+w[2])/2-page.rect.x0)/page.rect.width<=x1 and y0<=((w[1]+w[3])/2-page.rect.y0)/page.rect.height<=y1)
        contract_txt=box_text(.04,.075,.23,.097)
        contract_match=re.search(r'\b\d{7,10}\b',contract_txt) or contract
        fecha_txt=box_text(.26,.101,.36,.125)
        fecha_match=re.search(r'\d{2}/\d{2}/\d{4}',fecha_txt)
        gestion_txt=box_text(.58,.075,.70,.098)
        inicio_match=re.search(r'\d{2}/\d{2}/\d{4}',gestion_txt)
        saldo_txt=box_text(.70,.075,.83,.098)
        saldo_match=re.search(NUM,saldo_txt)
        delta_txt=box_text(.84,.075,.96,.098)
        delta_match=re.search(NUM,delta_txt)
        clasificacion=box_text(.045,.103,.255,.126).strip()
        ejecutivo=box_text(.37,.103,.72,.126).replace('Nombre:','').replace('Ext.','').strip()
        gestor=box_text(.73,.103,.96,.126).strip()
        # Regiones de las cuatro tablas centrales (porcentajes de página).
        cartera_txt=region(page,.045,.33,.475,.42)
        rend_txt=region(page,.045,.257,.475,.333)
        mov_txt=region(page,.50,.257,.94,.309)
        dist_txt=region(page,.50,.309,.94,.42)
        cartera=parse_labeled(cartera_txt,CARTERA,2)
        rendimiento=parse_labeled(rend_txt,REND,2)
        movimientos=parse_labeled(mov_txt,MOV,3)
        distribucion=parse_labeled(dist_txt,DIST,2)
        # Tabla inferior: lectura por renglones de texto; tasas vacías se preservan como None.
        portfolio_txt=region(page,.035,.485,.95,.79)
        portfolio=[]
        for line in portfolio_txt.splitlines():
            s=' '.join(line.split())
            parts=s.split()
            if len(parts)<8 or not re.fullmatch(r'[A-Z][A-Z0-9.-]{1,20}',parts[0]):continue
            if parts[0] in ('EMISORA','SERIE','TITULOS','PRECIO','TOTAL','VECTOR','VENCER'):continue
            try:
                # 10 columnas con tasa, 9 sin tasa. Un punto aislado indica tasa vacía.
                if len(parts)<9:continue
                emi,serie=parts[0],parts[1]
                tit,precio,total,pct=map(number,parts[2:6])
                rest=parts[6:]
                if len(rest)==4:
                    tasa=number(rest[0]);dias,trans,dt=map(number,rest[1:])
                elif len(rest)==3:
                    tasa=None;dias,trans,dt=map(number,rest)
                else:continue
                if None in (tit,precio,total,pct,dias,trans,dt):continue
                portfolio.append(dict(emisora=emi,serie=serie,titulos=tit,precio_mercado=precio,total=total,porcentaje=pct,tasa_vector=tasa,dias_vencer=dias,dias_trans_cupon=trans,dias_trans=dt))
            except (ValueError,TypeError):continue
        meta=dict(archivo=path.name,contrato=contract_match.group() if contract_match else None,fecha_reporte=fecha_match.group() if fecha_match else None,inicio_gestion=inicio_match.group() if inicio_match else None,saldo_inicial=number(saldo_match.group()) if saldo_match else None,delta_actual=number(delta_match.group()) if delta_match else None,clasificacion=clasificacion,ejecutivo=ejecutivo,gestor=gestor)
        issues=[]
        if not meta['contrato'] or not meta['fecha_reporte']:issues.append('Falta contrato o fecha')
        if len(cartera)<6:issues.append(f'Cartera incompleta: {len(cartera)}/6')
        if len(rendimiento)<4:issues.append(f'Rendimientos incompletos: {len(rendimiento)}/4')
        if len(movimientos)<2:issues.append(f'Movimientos incompletos: {len(movimientos)}/2')
        if len(distribucion)<6:issues.append(f'Distribución incompleta: {len(distribucion)}/6')
        if len(portfolio)<10:issues.append(f'Portafolio incompleto: {len(portfolio)} posiciones')
        if 'SALDO TOTAL' in cartera and 'INSTRUMENTOS' in cartera and 'EFECTIVOS' in cartera and 'EFECTIVOS 24 HRS' in cartera and 'EFECTIVO 48 HRS. Y 72 HRS' in cartera and 'REPORTOS A MAS DE 72 HRS' in cartera:
            c=cartera
            calc=sum(c[k][0] for k in CARTERA[:-1]);saldo=c['SALDO TOTAL'][0]
            if abs(calc-saldo)>.05:issues.append(f'Cuadre cartera: diferencia {calc-saldo:.2f}')
        return dict(meta=meta,cartera=cartera,rendimiento=rendimiento,movimientos=movimientos,distribucion=distribucion,portafolio=portfolio,observaciones=issues,estado='OK' if not issues else 'REVISAR')

def save_csv(path,rows,columns):
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.DictWriter(f,fieldnames=columns,extrasaction='ignore');w.writeheader();w.writerows(rows)

def main():
    ap=argparse.ArgumentParser(description='Extrae reportes BBVA desde una carpeta')
    ap.add_argument('carpeta',type=Path);ap.add_argument('--salida',type=Path);ap.add_argument('--recursivo',action='store_true')
    args=ap.parse_args()
    root=args.carpeta.expanduser().resolve();out=(args.salida or root/'Resultados').resolve()
    if not root.is_dir():print('ERROR: carpeta inexistente',file=sys.stderr);return 2
    files=sorted(root.rglob('*.pdf') if args.recursivo else root.glob('*.pdf'))
    if not files:print('ERROR: no se encontraron PDF',file=sys.stderr);return 2
    out.mkdir(parents=True,exist_ok=True)
    records=[];errors=[]
    for p in files:
        try:
            r=extract(p);records.append(r)
            print(f"{r['estado']}: {p.name}; portafolio={len(r['portafolio'])}; alertas={'; '.join(r['observaciones'])}")
        except Exception as e:
            errors.append({'archivo':p.name,'error':str(e)});print(f'ERROR: {p.name}: {e}',file=sys.stderr)
    tables={k:[] for k in ['Gestion','Cartera','Rendimientos','Movimientos','Distribucion','Portafolio','Auditoria']}
    for r in records:
        meta=r['meta'];key={k:meta[k] for k in ('archivo','contrato','fecha_reporte')}
        tables['Gestion'].append(meta)
        for section,dest,fields in [('cartera','Cartera',['importe','porcentaje']),('rendimiento','Rendimientos',['porcentaje','dias']),('movimientos','Movimientos',['dia','mes','anio']),('distribucion','Distribucion',['importe','porcentaje'])]:
            for label,values in r[section].items():tables[dest].append(dict(**key,concepto=label,**dict(zip(fields,values))))
        for item in r['portafolio']:tables['Portafolio'].append(dict(**key,**item))
        tables['Auditoria'].append(dict(**key,estado=r['estado'],posiciones=len(r['portafolio']),observaciones='; '.join(r['observaciones'])))
    for e in errors:tables['Auditoria'].append(dict(archivo=e['archivo'],estado='ERROR',observaciones=e['error']))
    for name,rows in tables.items():
        if rows:
            columns=list(dict.fromkeys(k for row in rows for k in row))
            save_csv(out/(name+'.csv'),rows,columns)
    (out/'resultado.json').write_text(json.dumps({'documentos':records,'errores':errors},ensure_ascii=False,indent=2),encoding='utf-8')
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font,PatternFill,Alignment
        from openpyxl.utils import get_column_letter
        wb=Workbook();wb.remove(wb.active)
        for name,rows in tables.items():
            ws=wb.create_sheet(name)
            if not rows:continue
            cols=list(dict.fromkeys(k for row in rows for k in row));ws.append(cols)
            for row in rows:ws.append([row.get(k) for k in cols])
            ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for c in ws[1]:c.fill=PatternFill('solid',fgColor='078348');c.font=Font(bold=True,color='FFFFFF');c.alignment=Alignment(wrap_text=True)
            for i,col in enumerate(cols,1):ws.column_dimensions[get_column_letter(i)].width=min(45,max(16,len(col)+3))
            for cells in ws.iter_rows(min_row=2):
                for c in cells:
                    if isinstance(c.value,float):c.number_format='#,##0.00;[Red]-#,##0.00'
        wb.save(out/'resultado_bbva.xlsx')
    except ImportError:
        print('AVISO: openpyxl no instalado, se generaron CSV y JSON',file=sys.stderr)
    (out/'estado.json').write_text(json.dumps({'total':len(files),'procesados':len(records),'ok':sum(r['estado']=='OK' for r in records),'revisar':sum(r['estado']=='REVISAR' for r in records),'errores':len(errors),'salida':str(out)},ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'FINALIZADO: {len(records)}/{len(files)} PDFs. Salida: {out}')
    return 1 if errors or any(r['estado']!='OK' for r in records) else 0

if __name__=='__main__':sys.exit(main())