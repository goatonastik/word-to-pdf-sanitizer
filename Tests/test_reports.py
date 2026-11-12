"""Report fidelity, workbook compatibility, recovery and cleanup checks."""
import csv
import json
from pathlib import Path
import tempfile
import sys
import unittest
from unittest.mock import patch
import zipfile
import openpyxl  # Independent reader used only in tests, not by the application.
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert as core
import pipeline
import maintenance
import reports
from runlogs import RunLogs


def journal(path,rows):
    with open(path,'w',newline='',encoding='utf-8-sig') as stream:csv.writer(stream).writerows(rows)


class Tests(unittest.TestCase):
    def test_excel_is_readable_typed_and_literal_with_filters_and_freeze(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'data.csvlog';target=root/'data.xlsx'
            values=[['Document','Comments removed','Status','Details'],['=HYPERLINK("x","y")','2','CONVERTED','Unicode: Zoë 漢字 & < >\nLine 2 _x0041_'],['+cmd','0','FAILED; SOURCE RETAINED','Failure details']]
            journal(source,values)
            self.assertEqual(reports.build_excel(source,target),2)
            book=openpyxl.load_workbook(target);sheet=book.active
            self.assertEqual(sheet.freeze_panes,'A2');self.assertEqual(sheet.auto_filter.ref,'A1:D3')
            self.assertEqual(sheet['A2'].value,values[1][0]);self.assertEqual(sheet['A2'].data_type,'s')
            self.assertEqual(sheet['B2'].value,2);self.assertEqual(sheet['B2'].data_type,'n')
            self.assertEqual(sheet['D2'].value,values[1][3])
            self.assertEqual(sheet['C3'].fill.fgColor.rgb,'FFFCE4D6')
            self.assertEqual(sheet['A1'].font.color.rgb,'FFFFFFFF')
            self.assertGreater(sheet.column_dimensions['A'].width,40)
            book.close()
            self.assertEqual(list(reports.report_rows(target)),[dict(zip(values[0],row)) for row in values[1:]])
            self.assertTrue(source.exists())
            with zipfile.ZipFile(target) as z:self.assertIsNone(z.testzip())

    def test_large_queue_splits_sheets_and_retains_all_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'data.csvlog';target=root/'data.xlsx'
            rows=[['Document','Details']]+[[str(i),'line '+str(i)] for i in range(8)]
            journal(source,rows)
            with patch.object(reports,'MAX_ROWS',4):reports.build_excel(source,target,'conversion-report')
            book=openpyxl.load_workbook(target)
            self.assertEqual(len(book.worksheets),3);book.close()
            self.assertEqual([r['Document'] for r in reports.report_rows(target)],[str(i) for i in range(8)])

    def test_workbooks_are_only_built_after_processing_and_recovery_survives(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);outputs=root/'Originals';recovery=root/'Logs'/'Recovery'
            run=RunLogs('Client [time]',recovery)
            path=run.path(outputs,'conversion-report','csv')
            journal(path,[['Source','Status'],['original.docx','CONVERTED']])
            self.assertFalse(outputs.exists());self.assertTrue(path.is_relative_to(recovery))
            self.assertEqual(run.build_reports(),[])
            mapping=run.finalize(True,False)
            workbook=Path(mapping[str(path)])
            self.assertEqual(workbook.suffix,'.xlsx');self.assertTrue(workbook.exists())
            self.assertTrue(any(recovery.rglob('*.csvlog')))
            self.assertFalse(any(root.rglob('*.csv')))

    def test_workbook_failure_keeps_journal_and_flags_queue_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            def batch(*args):
                path=core.RUN_LOGS.path(root/'Originals','conversion-report','csv')
                journal(path,[['Source','Status'],['original.docx','CONVERTED']])
                core.progress(report=str(path),finished=True,total=1,done=1,converted=1,failed=0)
            with patch.object(core,'HERE',root),patch.object(core,'DEBUG_OVERRIDE',False),patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'_batch',side_effect=batch),patch.object(reports,'build_excel',side_effect=PermissionError('simulated locked workbook')):
                pipeline.batch(root/'Source',root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertTrue(state['queueCompleted']);self.assertTrue(state['queueErrors'])
            self.assertEqual(state['converted'],1);self.assertEqual(state['failed'],0)
            self.assertTrue(Path(state['report']).exists());self.assertEqual(Path(state['report']).suffix,'.csvlog')
            self.assertIn('simulated locked workbook',state['reportErrors'][0]['error'])
            self.assertEqual(state['reportErrors'][0]['journal'],state['report'])
            self.assertIn('ERRORS',Path(state['report']).stem)

    def test_excel_limit_failure_preserves_full_recovery_value(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'data.csvlog';target=root/'data.xlsx'
            journal(source,[['Details'],['x'*32768]])
            reports.build_excel(source,target)
            value=next(reports.report_rows(target))['Details']
            self.assertLessEqual(len(value),32767);self.assertIn('[Truncated; full text',value)
            self.assertEqual(len(next(reports.report_rows(source))['Details']),32768)

    def test_existing_workbook_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'data.csvlog';target=root/'data.xlsx'
            journal(source,[['Details'],['new']]);target.write_bytes(b'keep existing')
            with self.assertRaises(FileExistsError):reports.build_excel(source,target)
            self.assertEqual(target.read_bytes(),b'keep existing')

    def test_maintenance_cli_produces_excel_after_restore(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir();originals=root/'#ORIGINALS';archive=originals/'Client';archive.mkdir(parents=True)
            saved=archive/'letter.docx';saved.write_bytes(b'original');pdf=source/'letter.pdf';pdf.write_bytes(b'generated')
            maintenance.save_restore_record(source/'letter.docx',pdf,saved,core.digest(saved),core.digest(pdf),{'archive':str(archive)})
            plan=root/'restore.plan';plan.write_text(json.dumps(maintenance.plan_restore(source,originals)))
            with patch('sys.argv',['maintenance.py','restore','--plan',str(plan)]):maintenance.main()
            workbook=plan.with_suffix('.xlsx')
            self.assertTrue(workbook.exists());self.assertTrue(plan.with_suffix('.csvlog').exists())
            self.assertFalse(plan.with_suffix('.csv').exists())
            self.assertEqual(next(reports.report_rows(workbook))['Outcome'],'RESTORED')
            self.assertEqual((source/'letter.docx').read_bytes(),b'original');self.assertFalse(pdf.exists())

    def test_log_deletion_preserves_excel_restore_proof(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir();originals=root/'#ORIGINALS';archive=originals/'Client';archive.mkdir(parents=True)
            saved=archive/'letter.docx';saved.write_bytes(b'original');pdf=source/'letter.pdf';pdf.write_bytes(b'generated')
            recover=root/'data.csvlog';target=archive/'Client [time] - conversion-report.xlsx'
            journal(recover,[['Source','PDF','Archive','Status','Details'],[source/'letter.docx',pdf,saved,'CONVERTED','1 pages; SHA256 '+core.digest(saved)]])
            reports.build_excel(recover,target)
            with patch.object(core,'HERE',root/'app'):
                plan=maintenance.plan_delete_logs(source,originals)
                self.assertEqual(plan['count'],1)
                self.assertEqual(maintenance.delete_logs(plan)[0][1],'DELETED')
            self.assertFalse(target.exists());self.assertTrue(saved.exists());self.assertTrue(pdf.exists())
            self.assertEqual(maintenance.plan_restore(source,originals)['ready'],1)
            self.assertTrue(any((archive/maintenance.RECORDS/'legacy-reports').glob('*.xlsx')))


if __name__=='__main__':unittest.main()
