import importlib.util
import io
from pathlib import Path
import tempfile
import unittest
import zipfile
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('mailer', ROOT / 'main.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

class MailTests(unittest.TestCase):
    def setUp(self):
        m.CONFIG = {'archive_prefix':'next-exam', 'student_email_domain':'example.com', 'subject':'回傳', 'assignment_name':'作業二', 'body':'{name}：{assignment_name}'}
        m.SENDER = 'ta@example.com'
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name)
        w = openpyxl.Workbook()
        w.active.append(['Email','學號','姓名'])
        w.active.append(['S00000001@example.com',' S00000001 ','範例學生'])
        w.save(self.path/'students.xlsx')
        data = io.BytesIO()
        with zipfile.ZipFile(data,'w') as z:z.writestr('Main.java','class Main {}')
        self.data = data.getvalue()
    def archive(self,names):
        with zipfile.ZipFile(self.path/'submissions.zip','w') as z:
            for n in names:z.writestr(n,self.data)
    def load(self):
        return m.load_plan(self.path/'students.xlsx',self.path/'submissions.zip')
    def test_corrupted_name_slash_and_reordered_columns(self):
        self.archive(['next-exam/S00000001_亂/碼.zip'])
        student = self.load()[0]
        self.assertEqual(student['學號'],'S00000001')
        self.assertEqual(student['data'],self.data)
        for test in [True,False]:
            msg = m.make_message(student,test)
            self.assertEqual(msg['To'],m.SENDER if test else student['Email'])
            self.assertEqual(list(msg.iter_attachments())[0].get_payload(decode=True),self.data)
    def test_duplicate_attachment_stops(self):
        self.archive(['next-exam/S00000001_a.zip','next-exam/S00000001_b.zip'])
        with self.assertRaisesRegex(ValueError,'多份'):self.load()
    def test_missing_attachment_stops(self):
        self.archive(['other-exam/S00000001.zip'])
        with self.assertRaisesRegex(ValueError,'名單不符'):self.load()
    def test_extra_attachment_stops(self):
        self.archive(['next-exam/S00000001.zip','next-exam/S00000002.zip'])
        with self.assertRaisesRegex(ValueError,'名單不符'):self.load()
    def test_wrong_email_stops(self):
        w=openpyxl.load_workbook(self.path/'students.xlsx');w.active['A2']='S00000002@example.com';w.save(self.path/'students.xlsx')
        self.archive(['next-exam/S00000001.zip'])
        with self.assertRaisesRegex(ValueError,'Email'):self.load()

if __name__=='__main__':unittest.main()
