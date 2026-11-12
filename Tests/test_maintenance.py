import csv
import tempfile
import unittest
import sys
from pathlib import Path
from unittest.mock import patch
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert as core
import maintenance as m


class Tests(unittest.TestCase):
    def test_conflicting_word_file_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder)
            source.write_bytes(b'newer Word document')
            plan=m.plan_restore(root,originals)
            self.assertEqual(plan['ready'],0)
            self.assertEqual(m.restore(plan)[0][1],'NOT RESTORED')
            self.assertEqual(source.read_bytes(),b'newer Word document')
            self.assertTrue(pdf.exists()); self.assertTrue(saved.exists())

    def test_changed_log_and_restore_records_are_retained(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder)
            app=folder/'app'; (app/'Logs').mkdir(parents=True)
            log=app/'Logs'/'Client [time] - debug.jsonl'; log.write_text('first')
            with patch.object(core,'HERE',app):
                plan=m.plan_delete_logs(root,originals)
                self.assertEqual(plan['count'],1)
                log.write_text('changed')
                self.assertEqual(m.delete_logs(plan)[0][1],'RETAINED')
            self.assertEqual(log.read_text(),'changed')
            self.assertEqual(m.plan_restore(root,originals)['ready'],1)

    def setup_files(self,folder,record=True):
        root=folder/'Client'; root.mkdir()
        originals=folder/'#ORIGINALS'; archive=originals/'Client'; archive.mkdir(parents=True)
        source=root/'letter.docx'; pdf=root/'letter.pdf'; saved=archive/'letter.docx'
        saved.write_bytes(b'original Word bytes'); pdf.write_bytes(b'generated PDF')
        if record: m.save_restore_record(source,pdf,saved,core.digest(saved),core.digest(pdf),{'archive':str(archive)})
        return root,originals,source,pdf,saved

    def test_restore_checks_both_files_and_leaves_review_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder)
            review=folder/'#REVIEW'/'Client'/'letter.docx'; review.parent.mkdir(parents=True); review.write_bytes(saved.read_bytes())
            unrelated=root/'unrelated.pdf'; unrelated.write_bytes(b'keep')
            plan=m.plan_restore(root,originals)
            self.assertEqual(plan['ready'],1)
            rows=m.restore(plan)
            self.assertEqual(rows[0][1],'RESTORED')
            self.assertEqual(source.read_bytes(),b'original Word bytes')
            self.assertFalse(pdf.exists()); self.assertFalse(saved.exists())
            self.assertEqual(review.read_bytes(),b'original Word bytes'); self.assertEqual(unrelated.read_bytes(),b'keep')

    def test_changed_pdf_or_original_is_never_deleted(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder)
            plan=m.plan_restore(root,originals)
            pdf.write_bytes(b'edited PDF')
            self.assertEqual(m.restore(plan)[0][1],'NOT RESTORED')
            self.assertTrue(saved.exists()); self.assertEqual(pdf.read_bytes(),b'edited PDF')
            self.assertFalse(source.exists())

    def test_log_deletion_keeps_old_restore_proof_and_all_documents(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder,False)
            report=saved.parent/'Client [time] - conversion-report.csv'
            with open(report,'w',newline='',encoding='utf-8-sig') as stream:
                writer=csv.writer(stream); writer.writerow(['Source','PDF','Archive','Status','Details'])
                writer.writerow([source,pdf,saved,'CONVERTED','1 pages; SHA256 '+core.digest(saved)])
            app=folder/'app'; (app/'Logs').mkdir(parents=True)
            log=app/'Logs'/'Client [time] - debug.jsonl'; log.write_text('diagnostic')
            with patch.object(core,'HERE',app):
                plan=m.plan_delete_logs(root,originals)
                self.assertEqual(plan['count'],2)
                rows=m.delete_logs(plan)
            self.assertTrue(all(row[1]=='DELETED' for row in rows),rows)
            self.assertFalse(report.exists()); self.assertFalse(log.exists())
            self.assertTrue(pdf.exists()); self.assertTrue(saved.exists())
            restore=m.plan_restore(root,originals)
            self.assertEqual(restore['legacy'],1)
            self.assertEqual(m.restore(restore)[0][1],'NOT RESTORED')
            self.assertEqual(m.restore(restore,allow_legacy=True)[0][1],'RESTORED')
            self.assertEqual(source.read_bytes(),b'original Word bytes')

    def test_outside_paths_in_record_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root,originals,source,pdf,saved=self.setup_files(folder)
            plan=m.plan_restore(root,originals)
            outside=folder/'outside.pdf'; outside.write_bytes(b'keep')
            plan['entries'][0]['pdf']=str(outside)
            self.assertEqual(m.restore(plan)[0][1],'NOT RESTORED')
            self.assertEqual(outside.read_bytes(),b'keep'); self.assertTrue(saved.exists())


if __name__=='__main__': unittest.main()
