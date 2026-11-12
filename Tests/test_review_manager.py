import csv
import json
from pathlib import Path
from types import SimpleNamespace
import sys
import tempfile
import unittest
from unittest.mock import patch

APP_ROOT=Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))

import convert
import maintenance
from reports import build_excel,report_rows
import review_manager


class Tests(unittest.TestCase):
    def make_report(self,folder,document=None,pdf=None):
        document=document or folder/'Source'/'letter.docx';pdf=pdf or document.with_suffix('.pdf')
        review=folder/'#REVIEW'/'Source'/'letter.docx';journal=folder/'review.csvlog';report=folder/'review-report.xlsx'
        with open(journal,'w',newline='',encoding='utf-8-sig') as stream:
            writer=csv.writer(stream);writer.writerow(['Document','PDF','Outcome','Revision handling','Reasons for review','Review copy','Copy result'])
            writer.writerow([document,pdf,'CONVERTED','Visible revisions preserved','Visible tracked changes preserved',review,'Added'])
        build_excel(journal,report,'Review');return report,document,pdf

    def test_accept_undo_and_report_columns_are_persistent(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory);decisions=folder/'decisions'
            report,document,pdf=self.make_report(folder)
            with patch.object(review_manager,'DECISION_ROOT',decisions):
                rows=review_manager.review_rows(report,report);self.assertEqual(len(rows),1);self.assertTrue(rows[0]['canDecide'])
                args=SimpleNamespace(report=report,id=rows[0]['id'],action='accept',document=document,pdf=pdf,root=folder/'Source',originals=folder/'Originals',settings=folder/'settings.json')
                accepted=review_manager.perform_action(args);self.assertEqual(accepted['decision'],'Accepted')
                review_manager.apply_decisions(report)
                saved=list(report_rows(report))[0];self.assertEqual(saved['Review decision'],'Accepted');self.assertIn('left unchanged',saved['Decision result'])
                args.action='undo';undone=review_manager.perform_action(args);self.assertEqual(undone['decision'],'')
                review_manager.apply_decisions(report)
                saved=list(report_rows(report))[0];self.assertEqual(saved['Review decision'],'Awaiting decision');self.assertIn('Accepted decision cleared',saved['Decision result'])

    def test_revert_uses_verified_restore_record(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory);root=folder/'Source';originals=folder/'Originals';archive=originals/root.name
            document=root/'Nested'/'letter.docx';pdf=document.with_suffix('.pdf');saved=archive/'Nested'/'letter.docx'
            saved.parent.mkdir(parents=True);pdf.parent.mkdir(parents=True);saved.write_bytes(b'original');pdf.write_bytes(b'pdf')
            maintenance.save_restore_record(document,pdf,saved,convert.digest(saved),convert.digest(pdf),{'archive':str(archive)})
            report,_,_=self.make_report(folder,document,pdf)
            with patch.object(review_manager,'DECISION_ROOT',folder/'decisions'):
                row=review_manager.review_rows(report,report)[0]
                result=review_manager.perform_action(SimpleNamespace(report=report,id=row['id'],action='revert',document=document,pdf=pdf,root=root,originals=originals,settings=folder/'settings.json'))
                self.assertEqual(result['decision'],'Reverted');self.assertEqual(document.read_bytes(),b'original')
                self.assertFalse(pdf.exists());self.assertFalse(saved.exists())

    def test_report_with_no_review_reason_has_no_detailed_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory);journal=folder/'scan.csvlog';report=folder/'scan.xlsx'
            with open(journal,'w',newline='',encoding='utf-8-sig') as stream:csv.writer(stream).writerows([['Folder','Scan error'],[folder,'']])
            build_excel(journal,report,'Scan')
            with patch.object(review_manager,'DECISION_ROOT',folder/'decisions'):
                self.assertEqual(review_manager.review_rows(report,report),[])

    @unittest.skipUnless(sys.platform=='win32','Windows sharing semantics')
    def test_exclusive_report_open_detects_another_process_lock(self):
        with tempfile.TemporaryDirectory() as directory:
            report,_,_=self.make_report(Path(directory))
            with review_manager.exclusive_snapshot(report):
                with self.assertRaises(PermissionError):review_manager.exclusive_snapshot(report)


if __name__=='__main__':unittest.main()
