import os,tempfile,unittest
os.environ['PCBA_DATA_DIRECTORY']=tempfile.mkdtemp(prefix='pcba-tests-')
from fastapi.testclient import TestClient
from app.main import app
import app.main as main
main.INVENTORY_OUTPUT_DIR=main.DATA
from io import BytesIO
from openpyxl import load_workbook

class Workflow(unittest.TestCase):
 def setUp(self): self.c=TestClient(app)
 def batch(self):
  d=self.c.post('/api/demo').json()
  r=self.c.post('/api/calculate',json=d)
  self.assertEqual(r.status_code,200,r.text)
  return r.json()
 def test_complete_workflow(self):
  b=self.batch();self.assertEqual(len(b['lines']),12);self.assertEqual(len(b['supplemented']),6)
  resistor=next(l for l in b['lines'] if l['code']=='R-0402-100R-01')
  self.assertEqual(resistor['auto'],'111378.000'); self.assertEqual(resistor['final'],'111378.000');self.assertEqual(resistor['gross'],'126000')
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{resistor['code']:{'final':'111000','reason':'试算'}},'purchase_basis':'final','simulation':[{'code':resistor['code'],'incoming':'900','outgoing':'50','production':'1000'}]})
  self.assertEqual(r.status_code,200,r.text);new=r.json()
  self.assertNotEqual(new['id'],b['id']);self.assertEqual(next(l for l in new['lines'] if l['code']==resistor['code'])['balance'],'-378.000')
  self.assertEqual(new['simulation'][0]['balance'],'14847.000')
  self.assertIsNone(self.c.get('/api/batches/'+b['id']).json()['purchase'])
  out=self.c.get('/api/batches/'+new['id']+'/export');self.assertEqual(out.status_code,200)
  wb=load_workbook(BytesIO(out.content));self.assertGreaterEqual(len(wb.sheetnames),9)
  self.assertTrue(wb['发料记录']['I2'].value.startswith('='))
  cached=load_workbook(BytesIO(out.content),data_only=True)
  self.assertIsInstance(cached['发料记录']['I2'].value,(int,float))
  inv=load_workbook(BytesIO(self.c.get('/api/batches/'+b['id']+'/export?inventory_only=true').content))
  self.assertEqual(inv['库存明细'].max_row,13)
 def test_export_fields_formats_and_suppliers(self):
  from openpyxl.utils import get_column_letter
  b=self.batch()
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text)
  b=r.json()
  supplier=next(x for x in b['purchase']['lines'] if x['code']=='R-0402-100R-01')['supplier']
  self.assertEqual(supplier,'供应商1')
  wb=load_workbook(BytesIO(self.c.get('/api/batches/'+b['id']+'/export').content))
  self.assertEqual(wb['BOM 明细']['A1'].value,'客户物料编码')
  self.assertEqual(wb['订单汇总']['B1'].value,'客户物料编码')
  self.assertEqual([c.value for c in wb['库存明细'][1]][3:6],['库存','最小采购单位','供应商'])
  self.assertEqual({row[0].value for row in list(wb['BOM 明细'].rows)[1:]},{'A79A1900032','ZNDZ000042'})
  for name in ['发料记录']:
   headers=[c.value for c in wb[name][1]]
  po_sheets=[n for n in wb.sheetnames if n.startswith('采购订单-')]
  self.assertTrue(po_sheets)
  for name in po_sheets:
   headers=[c.value for c in wb[name][1]]
   code=headers.index('物料编号')
   self.assertEqual(headers[code+1:code+3],['规格型号','封装'])
  long_fields={'客户物料编码','物料编号','规格型号','规格/型号','封装','供应商','来源'}
  for sheet in wb:
   for j,c in enumerate(sheet[1],1):
    wide=c.value in long_fields
    self.assertEqual(c.alignment.horizontal,'left' if wide else 'center',(sheet.title,c.coordinate))
    letter=get_column_letter(j)
    dim=next(d for d in sheet.column_dimensions.values() if d.min<=j<=d.max)
    expected={'客户物料编码':16,'物料编号':19,'规格型号':36,'规格/型号':36,'BOM版本':12,'采购数量':15,'单机用量':15,'损耗后需求':16,'最小采购单位':15,'封装':30,'供应商':30,'来源':30}.get(c.value,12)
    self.assertGreaterEqual(dim.width,1.7109375)
    if sheet.max_row>1:
     self.assertEqual(sheet.cell(2,j).alignment.horizontal,'left' if wide else 'center',(sheet.title,letter))
 def test_dated_inventory_closing_and_overwrite(self):
  b=self.batch()
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{'R-0402-100R-01':{'final':'120000','reason':'整盘采购'}},'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text)
  first=r.json();self.assertRegex(first['closing_inventory']['name'],r'^库存表-\d{8}\.xlsx$')
  wb=load_workbook(BytesIO(self.c.get('/api/batches/'+first['id']+'/closing-inventory').content),data_only=True)
  records={row[0]:row[3] for row in wb.active.iter_rows(min_row=2,values_only=True)}
  self.assertEqual(records['R-0402-100R-01'],8622)
  self.assertEqual(len(records),12)
  target=main.INVENTORY_OUTPUT_DIR/first['closing_inventory']['name']
  self.assertTrue(target.exists())
  r=self.c.post('/api/batches/'+first['id']+'/revise',json={'adjustments':{'R-0402-100R-01':{'final':'121000','reason':'新增采购'}},'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text)
  wb=load_workbook(target,data_only=True)
  self.assertEqual({row[0]:row[3] for row in wb.active.iter_rows(min_row=2,values_only=True)}['R-0402-100R-01'],9622)
  historical=load_workbook(BytesIO(self.c.get('/api/batches/'+first['id']+'/closing-inventory').content),data_only=True)
  self.assertEqual({row[0]:row[3] for row in historical.active.iter_rows(min_row=2,values_only=True)}['R-0402-100R-01'],8622)
 def test_closing_inventory_material_format(self):
  b=self.batch(); r=self.c.post('/api/batches/'+b['id']+'/revise',json={'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text)
  wb=load_workbook(BytesIO(self.c.get('/api/batches/'+r.json()['id']+'/closing-inventory').content))
  ws=wb.active
  self.assertEqual(ws['A1'].value,'物料编号')
  self.assertAlmostEqual(ws.column_dimensions['A'].width,19.7109375,places=4)
  self.assertEqual(ws['A1'].alignment.horizontal,'left'); self.assertEqual(ws['A2'].alignment.horizontal,'left')
 def test_dated_inventory_write_failure(self):
  from unittest.mock import patch
  b=self.batch()
  with patch.object(main,'INVENTORY_OUTPUT_DIR',main.DATA/'missing-directory'):
   r=self.c.post('/api/batches/'+b['id']+'/revise',json={'purchase_basis':'final'})
  self.assertEqual(r.status_code,400)
  self.assertIsNone(self.c.get('/api/batches/'+b['id']).json()['purchase'])
 def test_block_bad_quantity(self):
  d=self.c.post('/api/demo').json();d['orders']['rows'][0]['qty']='0'
  self.assertEqual(self.c.post('/api/calculate',json=d).status_code,400)
 def test_adjustment_requires_reason(self):
  b=self.batch();l=b['lines'][0]
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{l['code']:{'final':'1'}}})
  self.assertEqual(r.status_code,400)
 def test_minimum_purchase_unit_rounds_automatic_final(self):
  d=self.c.post('/api/demo').json()
  for row in d['inventory']['rows']:
   if row['code']=='IC-99013F': row['min_unit']='1000'
  b=self.c.post('/api/calculate',json=d).json(); l=next(x for x in b['lines'] if x['code']=='IC-99013F')
  self.assertEqual(l['auto'],'6063.000'); self.assertEqual(l['final'],'7000'); self.assertEqual(l['balance'],'937.000')
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{l['code']:{'final':'6063','reason':'人工按实际数量'}},'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text); self.assertEqual(next(x for x in r.json()['lines'] if x['code']==l['code'])['final'],'6063')
 def test_supplier_comes_from_imported_inventory(self):
  d=self.c.post('/api/demo').json()
  for row in d['inventory']['rows']:
   row['supplier']='当前导入供应商-'+row['code']
  b=self.c.post('/api/calculate',json=d).json()
  self.assertEqual(next(l for l in b['lines'] if l['code']=='R-0402-100R-01')['supplier'],'当前导入供应商-R-0402-100R-01')
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'purchase_basis':'final'})
  self.assertEqual(r.status_code,200,r.text)
  po=r.json()['purchase']
  self.assertIn('当前导入供应商-R-0402-100R-01',{o['supplier'] for o in po['orders']})
  self.assertTrue(all(o['number'].startswith('PO-') for o in po['orders']))
 def test_upload_actual_formats(self):
  from app.main import ROOT
  for kind,name in [('order','客户订单-待处理.xls'),('stock','物料库存总表.xls')]:
   r=self.c.post('/api/import?kind='+kind,files={'file':(name,(ROOT/name).read_bytes())})
   self.assertEqual(r.status_code,200,r.text)
  d=self.c.get('/api/boms').json();self.assertEqual(len(d['boms']),2)

if __name__=='__main__':unittest.main()
