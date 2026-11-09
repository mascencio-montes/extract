#!/usr/bin/env python3
"""Santander Posicion Diaria - extractor local para Automation Anywhere A360.
Uso: python extractor_posicion_diaria.py "C:\\RPA\\SANTANDER\\POSICION DIARIA" --salida "C:\\RPA\\SALIDAS\\POSICION"
Dependencia: pip install pdfplumber
"""
import argparse, csv, json, re, sys, hashlib
from pathlib import Path
from decimal import Decimal
from datetime import datetime

N=r'-?\d[\d,]*(?:\.\d+)?'
MESES={'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12}
CATEGORIAS=['Renta Variable','Sociedades de Inversion Renta Variable','Sociedades de Inversion Deuda','Sociedades de Inversion de Capitales','Sociedades de Inversion de Objeto','Renta Fija','Mercado de Dinero','Otras Custodias','Reportos Vigentes','Constancias de Deposito','Efectivo','Intereses Devengados No Cobrados']

def number(s):
    try: return str(Decimal(str(s).replace(',','').replace('$','').strip()))
    except: return ''
def date_of(s,year=2026):
    m=re.search(r'(\d{1,2})[/-]([A-Za-z]{3}|\d{1,2})[/-]?(\d{2,4})?',s,re.I)
    if not m:return ''
    d,mo,y=m.groups(); month=MESES.get(mo.upper(),int(mo) if mo.isdigit() else 0)
    if not month:return ''
    yr=int(y) if y else year
    if yr<100:yr+=2000
    try:return datetime(yr,month,int(d)).strftime('%Y-%m-%d')
    except:return ''
def clean(s):return ' '.join(s.split())
def extract(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        pages=[p.extract_text(layout=False) or '' for p in pdf.pages]
        # layout=True conserva separacion de columnas para extraer importes por fila
        layouts=[p.extract_text(layout=True,x_tolerance=1,y_tolerance=3) or '' for p in pdf.pages]
    whole='\n'.join(pages); first=pages[0]; flat=clean(first)
    acct=re.search(r'CUENTA\s*NO\s*:\s*(\d+)',flat,re.I)
    per=re.search(r'PERIODO\s+(\d{2}/[A-Z]{3}/\d{4})\s+AL\s+(\d{2}/[A-Z]{3}/\d{4})',flat,re.I)
    fecha=date_of(per.group(2)) if per else (re.search(r'\d{4}-\d{2}-\d{2}',path.name).group() if re.search(r'\d{4}-\d{2}-\d{2}',path.name) else '')
    base={'archivo':str(path),'fecha_corte':fecha,'cuenta':acct.group(1) if acct else '', 'paginas':len(pages)}
    alerts=[]
    if not acct:alerts.append('CUENTA_NO_ENCONTRADA')
    if not per:alerts.append('PERIODO_NO_ENCONTRADO')
    # Primera pagina: columnas izquierda/derecha en una misma linea. Cortar por coordenada de pagina.
    with pdfplumber.open(path) as pdf:
        p=pdf.pages[0]; w=p.width
        left=p.crop((0,0,w*.56,p.height)).extract_text() or ''
        right=p.crop((w*.54,0,w,p.height)).extract_text() or ''
    summary=[]
    roman=r'\b(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)\.-'
    for line in left.splitlines():
        if re.search(roman,line):
            m=re.search(r'('+N+r')\s+(\d+(?:\.\d+)?)\s*%',line)
            if m:
                desc=clean(re.split(roman,line,maxsplit=1)[-1].split('$')[0]).rstrip(':')
                summary.append({**base,'categoria':desc,'importe_mxn':number(m.group(1)),'porcentaje':number(m.group(2))})
    mt=re.search(r'\bTOTAL\s*\$?\s*('+N+r')\s+100\.00',left,re.I)
    total=number(mt.group(1)) if mt else ''
    results={**base,'total_valores_mxn':total}
    # Resultados: anclar conceptos, tomar ultimo numero de la linea en columna derecha
    labels={'Importe de los Valores':'importe_valores','Operaciones pendientes por':'operaciones_pendientes','Valor de la cuenta el mes actual':'valor_cuenta_actual','Depósitos en efectivo/valores':'depositos','Retiros de efectivo/valores':'retiros','Plusvalía o Minusvalía':'plusvalia_minusvalia'}
    for line in right.splitlines():
        for label,key in labels.items():
            if label.casefold() in line.casefold():
                nums=re.findall(N,line)
                if nums:results[key]=number(nums[-1])
    # Posicion valores: parseo por linea usando ancla 3 bloques numericos al inicio
    positions=[]; section=''
    for pi,layout in enumerate(layouts):
        for line in layout.splitlines():
            sec=re.search(r'\b(?:I|VI|VII)\.-\s*(RENTA VARIABLE|RENTA FIJA|MERCADO DE DINERO)',line,re.I)
            if sec:section=sec.group(1).upper()
            if 'IX.- REPORTOS VIGENTES' in line:section='REPORTOS';continue
            if 'MOVIMIENTOS DEL PERIODO' in line:section='MOVIMIENTOS';continue
            if section not in ('RENTA VARIABLE','RENTA FIJA','MERCADO DE DINERO'):continue
            m=re.match(r'^\s*('+N+r')\s+('+N+r')\s+('+N+r')\s+(.+?)\s+('+N+r')\s+('+N+r')\s+('+N+r')\s+('+N+r')\s+('+N+r')\s*$',line)
            if not m:continue
            a,b,c,middle,cost,price,value,pl,pct=m.groups()
            # 'middle' contiene emisora y cupon, cuando este existe
            middle=clean(middle)
            cm=re.match(r'(.+?)\s+(\d{1,2})$',middle)
            emisora=cm.group(1) if cm and section=='RENTA FIJA' else middle
            cupon=cm.group(2) if cm and section=='RENTA FIJA' else ''
            if not re.search(r'[A-Za-z]',emisora):continue
            positions.append({**base,'pagina':pi+1,'categoria':section,'emisora':emisora,'cupon':cupon,'titulos_mes_anterior':number(a),'valor_mercado_mes_anterior':number(b),'titulos_mes_actual':number(c),'costo_unitario':number(cost),'precio_cierre':number(price),'valor_mercado':number(value),'plusvalia_minusvalia':number(pl),'porcentaje_cartera':number(pct)})
    # Reportos: el renglon REPORTO NO. y su siguiente renglon emisora forman una sola operacion
    reportos=[]
    for pi,txt in enumerate(pages):
        lines=txt.splitlines()
        for j,line in enumerate(lines):
            m=re.search(r'REPORTO\s+NO\.\s*(\d+)\s+(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(\d+)\s+('+N+r')\s+(\d+)\s+('+N+r')\s+('+N+r')\s+('+N+r')',line,re.I)
            if not m:continue
            op,ini,ven,dias,monto,plazo,tasa,actual,pct=m.groups()
            nextline=lines[j+1] if j+1<len(lines) else ''
            em=re.match(r'\s*([A-Z][A-Z0-9]+\s+\d{5,8})',nextline)
            reportos.append({**base,'pagina':pi+1,'numero_reporto':op,'fecha_inicio':date_of(ini, int(fecha[:4]) if fecha else 2026),'fecha_vencimiento':date_of(ven,int(fecha[:4]) if fecha else 2026),'dias_transcurridos':int(dias),'emisora':em.group(1) if em else '','monto_inversion':number(monto),'plazo_dias':int(plazo),'tasa_rendimiento':number(tasa),'valor_actual':number(actual),'porcentaje_cartera':number(pct)})
    # Los movimientos son acumulados. Conservar evidencia completa y deduplicar globalmente por clave.
    movements=[]; in_mov=False
    for pi,txt in enumerate(pages):
        for line in txt.splitlines():
            if 'MOVIMIENTOS DEL PERIODO' in line:in_mov=True;continue
            if 'RESUMEN FISCAL' in line or 'SALDO DISPONIBLE' in line:in_mov=False
            if not in_mov:continue
            m=re.match(r'^\s*(\d{2}/\d{2})\s+(.+?)\s+([A-Z]{2}\d+/\d{3})\s*(.*)$',line)
            if not m:continue
            f,desc,folio,tail=m.groups()
            key=f'{base["cuenta"]}|{fecha[:4]}|{f}|{folio}|{clean(desc)}'
            movements.append({**base,'pagina':pi+1,'fecha_movimiento':date_of(f,int(fecha[:4]) if fecha else 2026),'descripcion':clean(desc),'folio':folio,'detalle_original':clean(tail),'clave_unica':key})
    # Fiscal, saldo disponible, pendientes: captura controlada sin inferir valores ausentes
    fiscal=[]
    for pi,txt in enumerate(pages):
        for line in txt.splitlines():
            if re.search(r'\b(?:I\.V\.A\.|I\.S\.R\.)',line,re.I):
                vals=re.findall(N,line.split(':')[-1]);
                if len(vals)>=2:fiscal.append({**base,'pagina':pi+1,'concepto':clean(line.split(':')[0]),'del_periodo':number(vals[-2]),'acumulado':number(vals[-1])})
    if not summary:alerts.append('RESUMEN_VACIO')
    if not positions:alerts.append('POSICIONES_VACIAS_REVISAR_LAYOUT')
    if total and summary:
        try:
            diff=sum(Decimal(x['importe_mxn']) for x in summary)-Decimal(total)
            if abs(diff)>Decimal('0.02'):alerts.append(f'DESCUADRE_RESUMEN_{diff}')
        except:alerts.append('ERROR_VALIDACION_RESUMEN')
    return {'base':base,'resumen':summary,'resultados':[results],'posiciones':positions,'reportos':reportos,'movimientos':movements,'fiscal':fiscal,'alertas':alerts}

def save_csv(path,rows):
    if not rows:return
    keys=list(dict.fromkeys(k for row in rows for k in row))
    with path.open('w',encoding='utf-8-sig',newline='') as f:
        wr=csv.DictWriter(f,fieldnames=keys,extrasaction='ignore');wr.writeheader();wr.writerows(rows)

def main():
    ap=argparse.ArgumentParser();ap.add_argument('carpeta');ap.add_argument('--salida',default='salida_posicion_diaria');args=ap.parse_args()
    root=Path(args.carpeta);out=Path(args.salida);out.mkdir(parents=True,exist_ok=True)
    files=sorted(root.rglob('*.pdf')) if root.is_dir() else ([root] if root.is_file() else [])
    data={k:[] for k in ('resumen','resultados','posiciones','reportos','movimientos','fiscal')};audit=[];seen=set()
    for path in files:
        try:
            r=extract(path)
            for k in data:
                for row in r[k]:
                    if k=='movimientos':
                        if row['clave_unica'] in seen:continue
                        seen.add(row['clave_unica'])
                    data[k].append(row)
            audit.append({**r['base'],'estado':'REVISAR' if r['alertas'] else 'OK','alertas':'; '.join(r['alertas']), 'posiciones':len(r['posiciones']),'reportos':len(r['reportos']),'movimientos_nuevos':len(r['movimientos'])})
        except Exception as e:audit.append({'archivo':str(path),'estado':'ERROR','alertas':str(e)})
    for name,rows in data.items():save_csv(out/(name+'.csv'),rows)
    save_csv(out/'auditoria.csv',audit)
    status={'pdf_encontrados':len(files),'pdf_ok':sum(x['estado']=='OK' for x in audit),'pdf_revisar':sum(x['estado']=='REVISAR' for x in audit),'pdf_error':sum(x['estado']=='ERROR' for x in audit),'filas':{k:len(v) for k,v in data.items()},'salida':str(out.resolve())}
    (out/'resultado_ejecucion.json').write_text(json.dumps(status,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(status,ensure_ascii=False));return 2 if status['pdf_error'] else (1 if status['pdf_revisar'] or not files else 0)
if __name__=='__main__':sys.exit(main())