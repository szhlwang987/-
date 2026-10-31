import os,tempfile,unittest
os.environ['PCBA_DATA_DIRECTORY']=tempfile.mkdtemp(prefix='pcba-tests-')
from fastapi.testclient import TestClient
from app.main import app
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
  self.assertEqual(resistor['auto'],'111378.000');self.assertEqual(resistor['gross'],'126000')
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{resistor['code']:{'final':'111000','reason':'试算'}},'purchase_basis':'final','simulation':[{'code':resistor['code'],'incoming':'900','outgoing':'50','production':'1000'}]})
  self.assertEqual(r.status_code,200,r.text);new=r.json()
  self.assertNotEqual(new['id'],b['id']);self.assertEqual(next(l for l in new['lines'] if l['code']==resistor['code'])['balance'],'-378.000')
  self.assertEqual(new['simulation'][0]['balance'],'14847.000')
  self.assertIsNone(self.c.get('/api/batches/'+b['id']).json()['purchase'])
  out=self.c.get('/api/batches/'+new['id']+'/export');self.assertEqual(out.status_code,200)
  wb=load_workbook(BytesIO(out.content));self.assertEqual(len(wb.sheetnames),9)
  self.assertTrue(wb['发料记录']['F2'].value.startswith('='))
  cached=load_workbook(BytesIO(out.content),data_only=True)
  self.assertIsInstance(cached['发料记录']['F2'].value,(int,float))
  inv=load_workbook(BytesIO(self.c.get('/api/batches/'+b['id']+'/export?inventory_only=true').content))
  self.assertEqual(inv['库存明细'].max_row,13)
 def test_block_bad_quantity(self):
  d=self.c.post('/api/demo').json();d['orders']['rows'][0]['qty']='0'
  self.assertEqual(self.c.post('/api/calculate',json=d).status_code,400)
 def test_adjustment_requires_reason(self):
  b=self.batch();l=b['lines'][0]
  r=self.c.post('/api/batches/'+b['id']+'/revise',json={'adjustments':{l['code']:{'final':'1'}}})
  self.assertEqual(r.status_code,400)
 def test_upload_actual_formats(self):
  from app.main import ROOT
  for kind,name in [('order','客户订单-待处理.xls'),('stock','物料库存总表.xls')]:
   r=self.c.post('/api/import?kind='+kind,files={'file':(name,(ROOT/name).read_bytes())})
   self.assertEqual(r.status_code,200,r.text)
  d=self.c.get('/api/boms').json();self.assertEqual(len(d['boms']),2)

if __name__=='__main__':unittest.main()
