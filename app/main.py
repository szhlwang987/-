from pathlib import Path
from decimal import Decimal, InvalidOperation
from datetime import datetime, timezone
from io import BytesIO
import os, json, sqlite3, hashlib, re, uuid, unicodedata
from zoneinfo import ZoneInfo
from threading import Lock
import pandas as pd
import xlsxwriter
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parents[1]
BOM_DIR = Path(os.getenv('BOM_DIRECTORY', str(ROOT))).resolve()
DATA = Path(os.getenv('PCBA_DATA_DIRECTORY', str(ROOT / '.local')))
DATA.mkdir(parents=True, exist_ok=True)
app = FastAPI(title='PCBA 物料计划')
INVENTORY_OUTPUT_DIR = BOM_DIR
INVENTORY_WRITE_LOCK = Lock()

def connection():
    c = sqlite3.connect(DATA / 'pcba.sqlite3')
    c.execute('CREATE TABLE IF NOT EXISTS batches(id TEXT PRIMARY KEY, created TEXT, payload TEXT)')
    return c

def dec(v):
    try:
        d = Decimal(str(v).strip())
        if not d.is_finite(): raise ValueError()
        return d
    except (InvalidOperation, ValueError):
        raise ValueError(f'无效数量：{v}')

def txt(v):
    return '' if pd.isna(v) else str(v).strip()

def sheets(raw):
    engine = 'openpyxl' if raw[:2] == b'PK' else 'xlrd'
    return pd.read_excel(BytesIO(raw), sheet_name=None, header=None, engine=engine)

def table(raw, aliases, required):
    for sheet, df in sheets(raw).items():
        for index, row in df.iterrows():
            cols = {}
            for i, v in enumerate(row):
                s = txt(v)
                for key, names in aliases.items():
                    if s in names: cols[key] = i
            if all(k in cols for k in required):
                result=[]
                for j in range(index+1,len(df)):
                    r=df.iloc[j]
                    record={key:txt(r.iloc[i]) for key,i in cols.items()}
                    if not record.get('code'): continue
                    record['_row']=j+1; record['_sheet']=sheet
                    result.append(record)
                return result
    raise ValueError('未识别必需表头：'+','.join(required))

ALIASES={'code':['物料编号','物料编码'], 'spec':['规格/型号','规格型号','型号'], 'package':['封装'], 'qty':['采购数量'], 'stock':['库存','结存','结存数量','库存数量'], 'usage':['用量','单机用量'], 'order':['订单号','订单编号'], 'supplier':['供应商'], 'min_unit':['最小包装','最小采购单位','最小采购量','采购最小单位','MOQ'], 'type':['类型','物料类型']}

def display_width(value):
    text=str(value if value is not None else '').strip()
    return sum(2 if unicodedata.east_asian_width(ch) in 'WFA' else 1 for ch in text)

def digest(raw): return hashlib.sha256(raw).hexdigest()

def ceil_unit(value, unit):
    if unit <= 0: return value
    return (value / unit).to_integral_value(rounding='ROUND_CEILING') * unit

def scan():
    if not BOM_DIR.is_dir(): raise ValueError(f'目录不存在：{BOM_DIR}')
    found=[]; errors=[]
    for p in sorted(BOM_DIR.iterdir()):
        if p.suffix.lower() not in ['.xls','.xlsx'] or p.name.startswith('~$') or p.is_symlink(): continue
        if any(w in p.stem for w in ['库存','订单','采购单','订购单']): continue
        try:
            raw=p.read_bytes(); rows=table(raw,ALIASES,['code','usage'])
            text=' '.join(txt(v) for df in sheets(raw).values() for row in df.values for v in row)
            version=re.search(r'REV[\w.-]+',text,re.I)
            found.append({'file':p.name,'code':p.stem.split('__')[0], 'version':version.group(0) if version else '未标识', 'sha':digest(raw),'rows':rows})
        except Exception as e: errors.append({'file':p.name,'message':str(e)})
    return {'directory':str(BOM_DIR),'boms':found,'errors':errors}

@app.get('/')
def home(): return FileResponse(ROOT/'app/static/index.html')

@app.get('/api/boms')
def boms():
    try: return scan()
    except Exception as e: raise HTTPException(400,str(e))

@app.post('/api/import')
async def import_file(kind:str,file:UploadFile=File(...)):
    raw=await file.read(20*1024*1024+1)
    if len(raw)>20*1024*1024: raise HTTPException(400,'文件不能超过20MB')
    name=Path(file.filename or '').name
    try:
        required={'order':['code','qty'],'stock':['code','stock']}.get(kind)
        if required is None: raise ValueError('不支持的文件类别')
        if kind=='stock' and '库存' not in name: raise ValueError('库存文件名须包含“库存”')
        rows=table(raw,ALIASES,required)
        if not rows: raise ValueError('没有数据行')
        sha=digest(raw)
        (DATA/(sha+'.source')).write_bytes(raw)
        return {'name':name,'sha':sha,'rows':rows}
    except Exception as e: raise HTTPException(400,str(e))

class Calculate(BaseModel):
    orders:dict
    inventory:dict
    versions:dict={}

@app.post('/api/calculate')
def calculate(req:Calculate):
    try:
        orders=req.orders.get('rows',[]); stocks=req.inventory.get('rows',[])
        if not orders or not stocks: raise ValueError('请先导入订单和库存')
        candidates=scan()['boms']; selected={}; issues=[]; materials={}; inventory={}
        for r in stocks:
            if r['code'] in inventory: raise ValueError('库存编码重复：'+r['code'])
            dec(r['stock']); inventory[r['code']]=r
        for r in orders:
            code=r['code']; q=dec(r.get('qty',''))
            if q<=0: issues.append(f'{code}：采购数量必须大于零'); continue
            options=[b for b in candidates if b['code']==code]
            choice=req.versions.get(code)
            if choice: options=[b for b in options if b['sha']==choice]
            if len(options)!=1: issues.append(f'{code}：缺少 BOM 或存在多个版本，请选择'); continue
            b=options[0]; selected[code]={'file':b['file'],'version':b['version'],'sha':b['sha']}
            for item in b['rows']:
                m=item['code']; usage=dec(item['usage'])
                if usage<=0: raise ValueError(f'{m}：BOM 用量必须大于零')
                spec=item.get('spec',''); package=item.get('package','')
                old=materials.get(m)
                if old and (old['spec']!=spec or old['package']!=package): issues.append(f'{m}：BOM 规格或封装冲突'); continue
                s=inventory.get(m)
                if s and any(s.get(k) and item.get(k) and s[k]!=item[k] for k in ['spec','package']): issues.append(f'{m}：库存与 BOM 属性冲突'); continue
                if old is None: materials[m]={'code':m,'spec':spec,'package':package,'gross':Decimal(0),'sources':[]}
                materials[m]['gross']+=q*usage
                materials[m]['sources'].append({'pcba':code,'qty':str(q),'usage':str(usage),'version':b['version'],'order':r.get('order','')})
        # V1 requires all errors resolved, preventing incomplete shared-material purchases.
        if issues: raise ValueError('；'.join(dict.fromkeys(issues)))
        lines=[]; supplemented=[]
        for m,item in sorted(materials.items()):
            s=inventory.get(m); stock=dec(s['stock']) if s else Decimal(0)
            demand=item['gross']*Decimal('1.003'); auto=max(Decimal(0),demand-stock)
            min_unit=dec((s or {}).get('min_unit','0') or '0')
            if min_unit < 0: raise ValueError(f'{m}：最小采购单位不能为负数')
            automatic_final=ceil_unit(auto,min_unit)
            item.update(gross=str(item['gross']),demand=str(demand),stock=str(stock),supplier=(s or {}).get('supplier',''),min_unit=str(min_unit),auto=str(auto),final=str(automatic_final),balance=str(stock+automatic_final-demand),reason='')
            lines.append(item)
            if not s:
                supplemented.append({'code':m,'spec':item['spec'],'package':item['package'],'stock':'0','source':'BOM 自动补充'})
        batch={'id':uuid.uuid4().hex,'created':datetime.now(timezone.utc).isoformat(),'orders':req.orders,'inventory':req.inventory,'boms':selected,'lines':lines,'supplemented':supplemented,'factor':'1.003','algorithm':'1.0','purchase':None,'simulation':[]}
        # Retain exact source BOM bytes for history independent of scanning directory.
        for b in selected.values():
            raw=(BOM_DIR/b['file']).read_bytes()
            if digest(raw)!=b['sha']: raise ValueError('BOM 正在修改，请重新扫描')
            (DATA/(b['sha']+'.source')).write_bytes(raw)
        return save(batch)
    except Exception as e: raise HTTPException(400,str(e))

def save(batch):
    with connection() as c: c.execute('INSERT INTO batches VALUES(?,?,?)',(batch['id'],batch['created'],json.dumps(batch,ensure_ascii=False)))
    return batch

def load(id):
    with connection() as c: r=c.execute('SELECT payload FROM batches WHERE id=?',(id,)).fetchone()
    if not r: raise HTTPException(404,'批次不存在')
    return json.loads(r[0])

@app.get('/api/history')
def history():
    with connection() as c: rows=c.execute('SELECT id,created,payload FROM batches ORDER BY created DESC').fetchall()
    return [{'id':r[0],'created':r[1],'count':len(json.loads(r[2])['lines'])} for r in rows]

@app.get('/api/batches/{id}')
def batch(id:str): return load(id)

class Revision(BaseModel):
    adjustments:dict={}
    purchase_basis:str|None=None
    simulation:list=[]

def closing_inventory(batch):
    rows={r['code']:dict(r) for r in batch['inventory']['rows']}
    for r in batch['supplemented']:
        rows.setdefault(r['code'],dict(r))
    for line in batch['lines']:
        rows[line['code']]['stock']=line['balance']
    return [rows[k] for k in sorted(rows)]

def closing_workbook(batch):
    output=BytesIO()
    workbook=xlsxwriter.Workbook(output,{'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    sheet=workbook.add_worksheet('库存明细')
    columns=[('code','物料编号'),('spec','规格/型号'),('package','封装'),('stock','库存'),('min_unit','最小包装'),('supplier','供应商')]
    for j,(key,label) in enumerate(columns):
        left=key in ['code','spec','package','supplier']
        widths={'code':19,'spec':36,'package':30,'supplier':30,'stock':12,'min_unit':15}
        fmt=workbook.add_format({'align':'left' if left else 'center','valign':'vcenter','num_format':'0.#########'})
        values=[label]+[str(row.get(key,'')) for row in closing_inventory(batch)]; sheet.set_column(j,j,max(1,max((display_width(v) for v in values),default=0)+5),fmt)
        sheet.write(0,j,label,workbook.add_format({'bold':True,'align':'left' if left else 'center','bg_color':'#DBEAFE','border':1}))
        for i,row in enumerate(closing_inventory(batch),1):
            value=float(dec(row['stock'])) if key=='stock' else row.get(key,'')
            sheet.write(i,j,value,fmt)
    sheet.freeze_panes(1,0)
    workbook.close()
    return output.getvalue()

@app.get('/api/batches/{id}/closing-inventory')
def download_closing_inventory(id:str):
    b=load(id)
    if not b.get('closing_inventory'): raise HTTPException(404,'该批次尚未生成日期库存表')
    from urllib.parse import quote
    return Response(closing_workbook(b),media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':"attachment; filename*=UTF-8''"+quote(b['closing_inventory']['name'])})

def next_purchase_number(date_code, count=1):
    used=[]
    with connection() as c:
        for (payload,) in c.execute('SELECT payload FROM batches'):
            try:
                for order in (json.loads(payload).get('purchase') or {}).get('orders',[]):
                    m=re.match(r'^PO-'+date_code+r'-(\d{3})$',order.get('number',''))
                    if m: used.append(int(m.group(1)))
            except Exception: pass
    start=max(used,default=0)+1
    return [f'PO-{date_code}-{start+i:03d}' for i in range(count)]

@app.post('/api/batches/{id}/revise')
def revise(id:str,req:Revision):
    b=load(id)
    try:
        for line in b['lines']:
            a=req.adjustments.get(line['code'])
            if a:
                final=dec(a['final'])
                if final<0: raise ValueError('最终发料不能为负数')
                if final!=dec(line['final']) and not a.get('reason','').strip(): raise ValueError('调整发料须填写原因')
                line['final']=str(final); line['reason']=a.get('reason','')
            line['balance']=str(dec(line['stock'])+dec(line['final'])-dec(line['demand']))
        if req.purchase_basis:
            if req.purchase_basis not in ['auto','final']: raise ValueError('请选择采购口径')
            all_lines=[dict(code=l['code'],qty=l[req.purchase_basis],supplier=l.get('supplier','')) for l in b['lines'] if dec(l[req.purchase_basis])>0]
            grouped={}
            for line in all_lines: grouped.setdefault(line['supplier'] or '未指定供应商',[]).append(line)
            orders=[]
            date_code=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d')
            numbers=next_purchase_number(date_code,len(grouped))
            for number,(supplier,lines) in zip(numbers,sorted(grouped.items())):
                orders.append({'number':number,'supplier':supplier,'lines':lines})
            b['purchase']={'basis':req.purchase_basis,'orders':orders,'lines':all_lines}
        else: b['purchase']=None
        by_code={l['code']:l for l in b['lines']}
        sims=[]
        for s in req.simulation:
            if s['code'] not in by_code: raise ValueError('模拟物料不存在')
            incoming,outgoing,production=[dec(s.get(k,'0')) for k in ['incoming','outgoing','production']]
            if min(incoming,outgoing,production)<0: raise ValueError('模拟输入不能为负数')
            sims.append(dict(s,balance=str(dec(by_code[s['code']]['stock'])+incoming-outgoing-production*Decimal('1.003'))))
        b['simulation']=sims; b['parent']=id; b['id']=uuid.uuid4().hex; b['created']=datetime.now(timezone.utc).isoformat()
        b.pop('closing_inventory',None)
        if req.purchase_basis:
            date=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d')
            name=f'库存表-{date}.xlsx'
            raw=closing_workbook(b)
            target=INVENTORY_OUTPUT_DIR/name
            temporary=INVENTORY_OUTPUT_DIR/('.'+uuid.uuid4().hex+'.tmp')
            with INVENTORY_WRITE_LOCK:
                try:
                    previous=target.read_bytes() if target.exists() else None
                    if previous is not None:
                        (DATA/('inventory-backup-'+uuid.uuid4().hex+'.xlsx')).write_bytes(previous)
                    temporary.write_bytes(raw)
                    os.replace(temporary,target)
                    b['closing_inventory']={'name':name,'path':str(target),'sha':digest(raw)}
                    try:
                        return save(b)
                    except Exception:
                        if previous is None: target.unlink(missing_ok=True)
                        else: target.write_bytes(previous)
                        raise
                finally:
                    temporary.unlink(missing_ok=True)
        return save(b)
    except Exception as e: raise HTTPException(400,str(e))

@app.get('/api/batches/{id}/export')
def export(id:str,inventory_only:bool=False):
    b=load(id); output=BytesIO(); w=xlsxwriter.Workbook(output,{'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    field_widths={'客户物料编码':16,'物料编号':19,'规格型号':36,'规格/型号':36,'BOM版本':12,'采购数量':15,'单机用量':15,'损耗后需求':16,'封装':30,'供应商':30,'来源':30,'最小包装':15}
    long_fields={'客户物料编码','物料编号','规格型号','规格/型号','封装','供应商','来源'}
    formats={align:w.add_format({'align':align,'valign':'vcenter','num_format':'0.#########'}) for align in ['left','center']}
    headers={align:w.add_format({'bold':True,'bg_color':'#DBEAFE','border':1,'align':align,'valign':'vcenter'}) for align in ['left','center']}
    def sheet(name,cols,rows):
        s=w.add_worksheet(name); s.freeze_panes(1,0)
        for j,label in enumerate(cols):
            align='left' if label in long_fields else 'center'
            base=field_widths.get(label,12)
            values=[str(label)]+[str(row[j]) for row in rows if j<len(row)]
            content_width=max(1,max((display_width(v) for v in values),default=0)+5)
            s.set_column(j,j,content_width,formats[align])
            s.write(0,j,label,headers[align])
        for i,row in enumerate(rows,1):
            for j,v in enumerate(row): s.write(i,j,v,formats['left' if cols[j] in long_fields else 'center'])
        if rows: s.autofilter(0,0,len(rows),len(cols)-1)
        return s
    inventory=[dict(r,source='原库存') for r in b['inventory']['rows']]+b['supplemented']
    if not inventory_only:
        sheet('订单汇总',['订单号','客户物料编码','采购数量'],[[r.get('order',''),r['code'],float(dec(r['qty']))] for r in b['orders']['rows']])
        details=[]
        for l in b['lines']:
            for src in l['sources']: details.append([src['pcba'],src['version'],l['code'],l['spec'],l['package'],float(dec(src['usage']))])
        sheet('BOM 明细',['客户物料编码','BOM版本','物料编号','规格/型号','封装','单机用量'],details)
        sheet('物料需求汇总',['物料编号','规格/型号','总用量','损耗后需求','库存'],[[l['code'],l['spec'],float(dec(l['gross'])),float(dec(l['demand'])),float(dec(l['stock']))] for l in b['lines']])
        s=sheet('发料记录',['物料编号','规格型号','封装','最小采购单位','总用量','库存','自动发料','最终发料','理论结存','修改原因'],[])
        for i,l in enumerate(b['lines'],1):
            n=i+1; s.write_row(i,0,[l['code'],l['spec'],l['package'],float(dec(l.get('min_unit','0'))),float(dec(l['gross'])),float(dec(l['stock']))])
            s.write_formula(i,6,f'=MAX(0,E{n}*1.003-F{n})',None,float(dec(l['auto'])))
            s.write_number(i,7,float(dec(l['final']))); s.write_formula(i,8,f'=F{n}+H{n}-E{n}*1.003',None,float(dec(l['balance']))); s.write(i,9,l['reason'],formats['center'])
        po=b['purchase']; lookup={l['code']:l for l in b['lines']}
        if po:
            for order in po['orders']:
                safe=re.sub(r'[^0-9A-Za-z一-龥_-]+','_',order['supplier'])[:20]
                sheet('采购订单-'+safe,['采购单号','物料编号','规格/型号','封装','最小包装','采购数量','供应商','其他字段'],[[order['number'],l['code'],lookup[l['code']]['spec'],lookup[l['code']]['package'],float(dec(lookup[l['code']].get('min_unit','0'))),float(dec(l['qty'])),order['supplier'],po['basis']] for l in order['lines']])
        else:
            sheet('采购订单',['采购单号','物料编号','规格/型号','封装','最小包装','采购数量','供应商','其他字段'],[])
        lookup={l['code']:l for l in b['lines']}
        sim_rows=[[r['code'],float(dec(lookup[r['code']]['stock'])),float(dec(r.get('incoming','0'))),float(dec(r.get('outgoing','0'))),float(dec(r.get('production','0'))),float(dec(r['balance']))] for r in b['simulation']]
        sim=sheet('库存模拟',['物料编号','库存','模拟入库','模拟发料','模拟净需求','模拟结存'],sim_rows)
        for i,r in enumerate(b['simulation'],1):
            n=i+1;sim.write_row(i,0,[r['code'],float(dec(lookup[r['code']]['stock'])),float(dec(r.get('incoming','0'))),float(dec(r.get('outgoing','0'))),float(dec(r.get('production','0')))])
            sim.write_formula(i,5,f'=B{n}+C{n}-D{n}-E{n}*1.003',None,float(dec(r['balance'])))
        sheet('异常检查',['问题','状态'],[])
        sheet('计算记录',['参数','值'],[['批次',b['id']],['时间',b['created']],['损耗系数',b['factor']],['算法',b['algorithm']],['BOM版本及摘要',json.dumps(b['boms'],ensure_ascii=False)],['订单摘要',b['orders'].get('sha','')],['库存摘要',b['inventory'].get('sha','')],['采购口径',po['basis'] if po else '未生成'],['采购订单数',len(po['orders']) if po else 0]])
    sheet('库存明细',['物料编号','规格/型号','封装','库存','最小采购单位','供应商','来源'],[[r['code'],r.get('spec',''),r.get('package',''),float(dec(r['stock'])),float(dec(r.get('min_unit','0') or '0')),r.get('supplier',''),r['source']] for r in sorted(inventory,key=lambda x:x['code'])])
    w.close()
    return Response(output.getvalue(),media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':f'attachment; filename="pcba-{id[:8]}.xlsx"'})

@app.post('/api/demo')
def demo():
    result={}
    for kind,name,required in [('orders','客户订单-待处理.xls',['code','qty']),('inventory','物料库存总表.xls',['code','stock'])]:
        p=ROOT/name
        if not p.exists(): raise HTTPException(404,'仓库样例不存在，请导入文件')
        raw=p.read_bytes(); sha=digest(raw); (DATA/(sha+'.source')).write_bytes(raw)
        result[kind]={'name':name,'sha':sha,'rows':table(raw,ALIASES,required)}
    return result
