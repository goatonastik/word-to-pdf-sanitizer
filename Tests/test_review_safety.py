"""Regression cases from Claude's read-only review."""
import csv
import json
import os
from pathlib import Path
import shutil
import stat
import sys
import tempfile
import unittest
from unittest.mock import patch
from concurrent.futures import Future
from reportlab.pdfgen.canvas import Canvas
from pypdf import PdfReader,PdfWriter
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert as core
import pipeline
import maintenance
import fileops
import runlogs
from reports import report_rows


def crash_first_verification(source,*args):
    if source.stem=='aCrash':os._exit(93)
    return pipeline.verify_and_archive(source,*args)


def crash_after_completed_verification(source,*args):
    result=pipeline.verify_and_archive(source,*args)
    if source.stem=='aCrashAfter':os._exit(93)
    return result


class FakeWord:
    instances=0;forced=0
    def __init__(self,*args):type(self).instances+=1;self.ready={'ok':True}
    def export(self,source,output,tick,original_name=None):
        c=Canvas(str(output));c.drawString(72,720,'Preserve content');c.save();shutil.copyfile(output,str(output)+'.baseline.pdf')
        return dict(timings={'export':.01},audit={})
    def close(self,force=False):
        if force:type(self).forced+=1


class Tests(unittest.TestCase):
    def test_readonly_word_round_trip_and_zone_identifier(self):
        if os.name!='nt':self.skipTest('Windows attribute/ADS test')
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);starting=root/'Client';starting.mkdir();originals=root/'Originals';archive=originals/'Client';archive.mkdir(parents=True)
            source=starting/'letter.docx';source.write_bytes(b'original bytes')
            zone=b'[ZoneTransfer]\r\nZoneId=3\r\n'
            with open(str(source)+':Zone.Identifier','wb') as stream:stream.write(zone)
            os.chmod(source,stat.S_IREAD)
            raw=root/'raw.pdf';clean=root/'clean.pdf';target=source.with_suffix('.pdf');saved=archive/source.name
            c=Canvas(str(raw));c.drawString(72,720,'Page');c.save();shutil.copyfile(raw,str(raw)+'.baseline.pdf')
            result=pipeline.verify_and_archive(source,target,saved,raw,clean,core.digest(source),dict(pipeline.DEFAULTS,archive=str(archive)))
            self.assertEqual(result['status'],'CONVERTED',result)
            self.assertFalse(source.exists());self.assertFalse(saved.stat().st_mode & stat.S_IWRITE)
            with open(str(saved)+':Zone.Identifier','rb') as stream:self.assertEqual(stream.read(),zone)
            os.chmod(target,stat.S_IREAD)
            plan=maintenance.plan_restore(starting,originals);rows=maintenance.restore(plan)
            self.assertEqual(rows[0][1],'RESTORED',rows)
            self.assertEqual(source.read_bytes(),b'original bytes');self.assertFalse(source.stat().st_mode & stat.S_IWRITE)
            with open(str(source)+':Zone.Identifier','rb') as stream:self.assertEqual(stream.read(),zone)
            self.assertFalse(target.exists());self.assertFalse(saved.exists())
            self.assertEqual(maintenance.plan_restore(starting,originals)['blocked'],0)

    def test_timestamp_failure_is_warning_after_verified_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'file.docx';source.write_bytes(b'original')
            raw=root/'raw.pdf';clean=root/'clean.pdf';target=root/'file.pdf';saved=root/'Originals'/'file.docx'
            c=Canvas(str(raw));c.drawString(72,720,'Page');c.save();shutil.copyfile(raw,str(raw)+'.baseline.pdf')
            with patch.object(pipeline.shutil,'copystat',side_effect=PermissionError('share rejects timestamp')):
                result=pipeline.verify_and_archive(source,target,saved,raw,clean,core.digest(source),dict(pipeline.DEFAULTS,archive=str(saved.parent)))
            self.assertEqual(result['status'],'CONVERTED',result);self.assertIn('archiveTimestampWarning',result['audit'])
            self.assertFalse(pipeline.review_reasons(result));self.assertEqual(saved.read_bytes(),b'original')

    def test_cloud_reparse_tags_are_not_redirecting_links(self):
        from types import SimpleNamespace
        for tag,mode,expected in ((0x9000001A,stat.S_IFREG,False),(0x80000013,stat.S_IFREG,False),(0xA0000003,stat.S_IFDIR,True),(0xA000000C,stat.S_IFLNK,True)):
            with self.subTest(tag=hex(tag)),patch.object(core.os,'lstat',return_value=SimpleNamespace(st_mode=mode,st_file_attributes=0x400,st_reparse_tag=tag)):
                self.assertEqual(core.is_link(Path('C:/placeholder.docx')),expected)

    def test_review_and_original_archive_must_not_overlap(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            with self.assertRaises(ValueError):core.validate_review_root(source,root/'#REVIEW'/'Client',root/'#REVIEW'/'Client')
            with self.assertRaises(ValueError):maintenance.roots(source,root/'#REVIEW')

    def test_nonstandard_pdf_property_names_are_not_logged(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);base=root/'base.pdf';raw=root/'raw.pdf';out=root/'out.pdf'
            c=Canvas(str(base));c.drawString(72,720,'Page');c.save()
            writer=PdfWriter(clone_from=PdfReader(base));writer.add_metadata({'/OtherClientEntitySecret':'Other client value','/Author':'Secret author'})
            with open(raw,'wb') as stream:writer.write(stream)
            writer.close();audit={};core.STRIP_METADATA=True
            core.clean_and_check(raw,out,audit=audit)
            self.assertNotIn('OtherClientEntitySecret',json.dumps(audit));self.assertNotIn('Other client value',json.dumps(audit))
            self.assertIn('1 non-standard field',audit['pdf_property_names_removed'])
            self.assertNotIn(b'OtherClientEntitySecret',out.read_bytes())

    def test_finalization_lock_preserves_files_and_does_not_raise(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);run=runlogs.RunLogs('Client [time]');path=run.path(root,'debug','jsonl');path.write_text('diagnostics')
            with patch.object(runlogs.os,'rename',side_effect=PermissionError('locked')),patch.object(runlogs.time,'sleep'):
                mapping=run.finalize(True,False)
            self.assertEqual(Path(mapping[str(path)]),path);self.assertEqual(path.read_text(),'diagnostics')
            self.assertEqual(len(run.finalization_errors),1)

    def test_large_root_name_gets_bounded_stable_log_label(self):
        name='Client '+('x'*210);label=runlogs.root_label(name)
        self.assertLessEqual(len(label.encode('utf-16-le'))//2,100);self.assertEqual(label,runlogs.root_label(name))
        self.assertNotEqual(label,runlogs.root_label(name+'y'))

    def test_word_document_error_keeps_reusable_session(self):
        class DocumentErrorWord(FakeWord):
            def export(self,source,*args):
                if source.stem=='aBad':raise pipeline.WordDocumentError('Rejected protected document')
                return super().export(source,*args)
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('aBad','bGood'):(source/(name+'.docx')).write_bytes(b'original')
            DocumentErrorWord.instances=0;DocumentErrorWord.forced=0
            with patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',DocumentErrorWord):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual((state['converted'],state['failed']),(1,1));self.assertEqual(DocumentErrorWord.instances,1);self.assertEqual(DocumentErrorWord.forced,0)

    def test_dead_verifier_is_rebuilt_without_restarting_word(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('aCrash','bGood','cGood'):(source/(name+'.docx')).write_bytes(b'original')
            settings=root/'settings.json';settings.write_text(json.dumps({'queueLimit':1}))
            FakeWord.instances=0;FakeWord.forced=0
            with patch.object(core,'SETTINGS_FILE',settings),patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',FakeWord),patch.object(pipeline,'verify_and_archive',crash_first_verification):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual((state['converted'],state['failed'],state['done']),(2,1,3),state)
            self.assertEqual(FakeWord.instances,1);self.assertEqual(FakeWord.forced,0)
            self.assertTrue((source/'aCrash.docx').exists());self.assertTrue((source/'bGood.pdf').exists())

    def test_aborting_queue_still_records_pending_completed_conversion(self):
        class AbortWord(FakeWord):
            def export(self,source,*args):
                if source.stem=='bAbort':raise KeyboardInterrupt('simulated queue interruption')
                result=super().export(source,*args);result['audit']['visibleRevisionsPreserved']=True;return result
        class RunningPool:
            def __init__(self,*args,**kwargs):self.jobs=[]
            def submit(self,fn,*args):
                future=Future();self.jobs.append((future,fn,args));return future
            def shutdown(self,**kwargs):
                for future,fn,args in self.jobs:
                    if not future.done():future.set_result(fn(*args))
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('aRunning','bAbort'):(source/(name+'.docx')).write_bytes(b'original')
            with patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',AbortWord),patch.object(pipeline,'ProcessPoolExecutor',RunningPool):
                with self.assertRaises(KeyboardInterrupt):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertFalse(state['queueCompleted']);self.assertEqual(state['converted'],1);self.assertEqual(state['reviewAdded'],1)
            self.assertTrue((root/'#REVIEW'/'Client'/'aRunning.docx').exists())
            rows=list(report_rows(state['report']));self.assertEqual(rows[0]['Status'],'CONVERTED');self.assertIn('Visible tracked changes',rows[0]['Review reasons'])

    def test_completed_work_is_recovered_after_worker_dies_before_reply(self):
        class VisibleWord(FakeWord):
            def export(self,source,*args):
                result=super().export(source,*args)
                result['audit']['visibleRevisionsPreserved']=source.stem=='aCrashAfter'
                return result
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('aCrashAfter','bGood'):(source/(name+'.docx')).write_bytes(b'original')
            settings=root/'settings.json';settings.write_text(json.dumps({'queueLimit':1}))
            with patch.object(core,'SETTINGS_FILE',settings),patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',VisibleWord),patch.object(pipeline,'verify_and_archive',crash_after_completed_verification):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual((state['converted'],state['failed'],state['reviewAdded']),(2,0,1),state)
            rows=list(report_rows(state['report']));self.assertIn('recovered from restore record',rows[0]['Details'])
            privacy=list(report_rows(state['privacyReport']));self.assertTrue(privacy[0]['Original final-view appearance'].startswith('Passed'))
            self.assertEqual((root/'#REVIEW'/'Client'/'aCrashAfter.docx').read_bytes(),b'original')

    def test_two_word_startup_failures_stop_queue_without_failing_every_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('a','b','c','d'):(source/(name+'.docx')).write_bytes(b'original')
            with patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',side_effect=pipeline.WordSessionError('simulated startup failure')) as start:
                with self.assertRaises(pipeline.QueueFatalError):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual(start.call_count,2);self.assertEqual(state['failed'],2);self.assertEqual(state['unprocessed'],2)
            self.assertFalse(state['queueCompleted']);self.assertEqual(len(list(source.glob('*.docx'))),4)

    def test_report_write_error_does_not_lose_other_outcomes_or_kill_word(self):
        writer=csv.writer;failed=[]
        class FailingRowWriter:
            def __init__(self,*args,**kwargs):self.writer=writer(*args,**kwargs)
            def writerow(self,row):
                if not failed and row and isinstance(row[0],Path) and str(row[0]).endswith('.docx'):
                    failed.append(True);raise OSError('simulated conversion journal write failure')
                return self.writer.writerow(row)
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('a','b'):(source/(name+'.docx')).write_bytes(b'original')
            FakeWord.instances=0;FakeWord.forced=0
            with patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',FakeWord),patch.object(pipeline.csv,'writer',FailingRowWriter):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual((state['converted'],state['done'],state['reviewAdded']),(2,2,0),state)
            self.assertTrue(state['queueErrors']);self.assertEqual(state['reportWriteErrors'],1)
            self.assertEqual(FakeWord.instances,1);self.assertEqual(FakeWord.forced,0)
            self.assertIn('report_row_fallback',Path(state['diagnostics']).read_text())

    def test_source_changed_during_restore_record_write_is_not_deleted(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'letter.docx';source.write_bytes(b'original')
            raw=root/'raw.pdf';clean=root/'clean.pdf';target=root/'letter.pdf';saved=root/'Originals'/'letter.docx'
            c=Canvas(str(raw));c.drawString(72,720,'Page');c.save();shutil.copyfile(raw,str(raw)+'.baseline.pdf')
            original=maintenance.save_restore_record
            def changed(*args):
                original(*args);source.write_bytes(b'newer user content')
            with patch.object(maintenance,'save_restore_record',side_effect=changed):
                result=pipeline.verify_and_archive(source,target,saved,raw,clean,core.digest(source),dict(pipeline.DEFAULTS,archive=str(saved.parent)))
            self.assertEqual(result['status'],'PDF CREATED; SOURCE RETAINED',result)
            self.assertEqual(source.read_bytes(),b'newer user content');self.assertEqual(saved.read_bytes(),b'original')

    def test_log_deletion_ignores_user_spreadsheet_and_includes_word_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir();originals=root/'Originals';archive=originals/'Client';archive.mkdir(parents=True)
            user=archive/'Client privacy-report.xlsx';user.write_bytes(b'user workbook')
            log=archive/'Client [time] - word-errors.log';log.write_text('error')
            with patch.object(core,'HERE',root/'app'):
                plan=maintenance.plan_delete_logs(source,originals)
                self.assertEqual(plan['count'],1);self.assertEqual(plan['files'][0]['path'],str(log))


if __name__=='__main__':unittest.main()
