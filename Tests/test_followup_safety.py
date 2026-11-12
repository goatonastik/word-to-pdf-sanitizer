import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import time
import unittest
from unittest.mock import patch,Mock
import zipfile
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from pypdf.generic import DictionaryObject,NameObject,TextStringObject,ArrayObject,NumberObject
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert as core
import pipeline
import fileops
import maintenance
from fieldcache import freeze_saved_fields
from debugtrace import DebugLog
from test_review_safety import FakeWord


def crash_during_publish(source,*args):
    if source.stem!='aCrashPublish':return pipeline.verify_and_archive(source,*args)
    def interrupted(src,dst,*unused):dst.write(src.read(80));dst.flush();os.fsync(dst.fileno());os._exit(94)
    with patch.object(fileops.shutil,'copyfileobj',side_effect=interrupted):return pipeline.verify_and_archive(source,*args)


def crash_during_archive(source,*args):
    if source.stem!='aCrashArchive':return pipeline.verify_and_archive(source,*args)
    copy=shutil.copyfileobj
    def interrupted(src,dst,*rest):
        if str(dst.name).lower().endswith('.docx'):
            dst.write(src.read(3));dst.flush();os.fsync(dst.fileno());os._exit(95)
        return copy(src,dst,*rest)
    with patch.object(fileops.shutil,'copyfileobj',side_effect=interrupted):return pipeline.verify_and_archive(source,*args)


class Tests(unittest.TestCase):
    def test_verifier_worker_auto_and_explicit_limits(self):
        with patch.object(pipeline.os,'cpu_count',return_value=20):
            self.assertEqual(pipeline.verifier_worker_count({'verifierWorkers':0}),6)
        self.assertEqual(pipeline.verifier_worker_count({'verifierWorkers':3}),3)
        for invalid in (-1,17):
            with self.assertRaises(ValueError):pipeline.verifier_worker_count({'verifierWorkers':invalid})

    def test_dynamic_field_flags_preserve_cache_and_leave_page_fields_unlocked(self):
        with tempfile.TemporaryDirectory() as directory:
            file=Path(directory)/'fields.docx';doc=Document()
            for code,value in (('DATE','1999-12-31'),('PAGE','3'),('USERNAME','Saved Author')):
                f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),' '+code+' ');f.set(qn('w:dirty'),'true');r=OxmlElement('w:r');t=OxmlElement('w:t');t.text=value;r.append(t);f.append(r);doc.add_paragraph()._p.append(f)
            # A complex DATE instruction split across runs must also be frozen.
            p=doc.add_paragraph()
            for tag,attribute,text in [('fldChar','begin',None),('instrText',None,' DA'),('instrText',None,'TE '),('fldChar','separate',None),('t',None,'1980-01-01'),('fldChar','end',None)]:
                r=OxmlElement('w:r');e=OxmlElement('w:'+tag)
                if attribute:e.set(qn('w:fldCharType'),attribute)
                if text:e.text=text
                r.append(e);p._p.append(r)
            doc.save(file)
            with zipfile.ZipFile(file) as z:before=z.read('word/styles.xml')
            result=freeze_saved_fields(file)
            self.assertEqual(result['dynamicFieldsLocked'],3)
            with zipfile.ZipFile(file) as z:
                xml=etree.fromstring(z.read('word/document.xml'));self.assertEqual(z.read('word/styles.xml'),before)
                fields=xml.findall('.//'+qn('w:fldSimple'))
                self.assertEqual(fields[0].get(qn('w:fldLock')),'true');self.assertIsNone(fields[1].get(qn('w:fldLock')))
                self.assertEqual(fields[0].find('.//'+qn('w:t')).text,'1999-12-31')
                begin=xml.find('.//'+qn('w:fldChar'));self.assertEqual(begin.get(qn('w:fldLock')),'true')

    def test_dynamic_field_without_cache_is_withheld(self):
        with tempfile.TemporaryDirectory() as directory:
            file=Path(directory)/'fields.docx';doc=Document();f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),' DATE ');doc.add_paragraph()._p.append(f);doc.save(file)
            original=file.read_bytes()
            with self.assertRaisesRegex(ValueError,'no saved display value'):freeze_saved_fields(file)
            self.assertEqual(file.read_bytes(),original)

    def test_signed_package_is_not_rewritten(self):
        with tempfile.TemporaryDirectory() as directory:
            file=Path(directory)/'fields.docx';doc=Document();f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),' DATE ');r=OxmlElement('w:r');t=OxmlElement('w:t');t.text='Old date';r.append(t);f.append(r);doc.add_paragraph()._p.append(f);doc.save(file)
            with zipfile.ZipFile(file,'a') as z:z.writestr('_xmlsignatures/sig1.xml','<Signature/>')
            original=file.read_bytes()
            with self.assertRaisesRegex(ValueError,'Signed Word package'):freeze_saved_fields(file)
            self.assertEqual(file.read_bytes(),original)

    def test_reply_written_during_slow_result_recording_is_not_timed_out(self):
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory);session=pipeline.WordSession.__new__(pipeline.WordSession)
            session.workspace=folder;session.timeout=.001;session.debug=DebugLog();session.process=Mock();session.process.poll.return_value=None;session.close=Mock()
            def slow_tick():time.sleep(.02);(folder/'result.json').write_text(json.dumps({'ok':True}))
            self.assertTrue(session.wait('result.json',slow_tick)['ok']);session.close.assert_not_called()

    def test_retained_links_allow_web_and_internal_but_reject_file_actions(self):
        for kind,target,allowed in (('/URI','https://example.test/a',True),('/URI','mailto:user@example.test',True),('/URI','file:///C:/Users/Test/Temp/a.pdf',False),('/URI','relative/a.pdf',False),('/Launch',None,False),('/GoToR',None,False),('/ImportData',None,False),('/GoTo',None,True)):
            action=DictionaryObject({NameObject('/S'):NameObject(kind)})
            if target is not None:action[NameObject('/URI')]=TextStringObject(target)
            annotation=DictionaryObject({NameObject('/A'):action})
            if allowed:core.validate_link(annotation)
            else:
                with self.assertRaises(ValueError):core.validate_link(annotation)
        action=DictionaryObject({NameObject('/S'):NameObject('/URI'),NameObject('/URI'):TextStringObject('https://example.test/C:/Temp/word-pdf-export-abc/file.pdf')})
        with patch.object(core,'LINK_WORKSPACE','C:/Temp/word-pdf-export-abc'):
            with self.assertRaisesRegex(ValueError,'workspace'):core.validate_link(DictionaryObject({NameObject('/A'):action}))

    def test_parent_removes_only_owned_partial_pdf_after_native_worker_exit(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir()
            for name in ('aCrashPublish','bGood'):(source/(name+'.docx')).write_bytes(b'original')
            settings=root/'settings.json';settings.write_text(json.dumps({'queueLimit':1}))
            with patch.object(core,'SETTINGS_FILE',settings),patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',FakeWord),patch.object(pipeline,'verify_and_archive',crash_during_publish):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual((state['failed'],state['converted']),(1,1),state)
            self.assertFalse((source/'aCrashPublish.pdf').exists());self.assertEqual((source/'aCrashPublish.docx').read_bytes(),b'original')

    def test_native_archive_exit_removes_partial_archive_and_persists_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir();word=source/'aCrashArchive.docx';word.write_bytes(b'original')
            settings=root/'settings.json';settings.write_text(json.dumps({'queueLimit':1}))
            with patch.object(core,'SETTINGS_FILE',settings),patch.object(core,'PROGRESS_FILE',root/'progress.json'),patch.object(pipeline,'WordSession',FakeWord),patch.object(pipeline,'verify_and_archive',crash_during_archive):pipeline.batch(source,root/'Originals')
            state=json.loads((root/'progress.json').read_text())
            self.assertEqual(state['partial'],1,state);self.assertFalse((root/'Originals'/'aCrashArchive.docx').exists());self.assertEqual(word.read_bytes(),b'original')
            self.assertTrue((source/'aCrashArchive.pdf').exists())
            with patch.object(core,'SETTINGS_FILE',settings),patch.object(core,'PROGRESS_FILE',root/'retry.json'),patch.object(pipeline,'WordSession') as session:pipeline.batch(source,root/'Originals');session.assert_not_called()
            again=json.loads((root/'retry.json').read_text());self.assertEqual(again['skipped'],0);self.assertEqual(again['partial'],1);self.assertEqual(again['reviewFlagged'],1)
            # Manual removal of the retained output resolves the marker and permits retry.
            (source/'aCrashArchive.pdf').unlink()
            self.assertIsNone(maintenance.incomplete_conversion_reason(word,word.with_suffix('.pdf'),root/'Originals'/'aCrashArchive.docx'))

    def test_original_text_change_is_withheld_even_with_identical_white_pixels(self):
        from reportlab.pdfgen.canvas import Canvas
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);before=root/'before.pdf';after=root/'after.pdf'
            for path,text in ((before,'Original text'),(after,'Changed text')):
                c=Canvas(str(path));c.setFillColorRGB(1,1,1);c.drawString(72,720,text);c.save()
            with self.assertRaisesRegex(ValueError,'Text changed from original reference'):core.compare_reference(before,after)

    def test_unsafe_retained_link_is_rejected_by_complete_pdf_verification(self):
        from reportlab.pdfgen.canvas import Canvas
        from pypdf import PdfReader,PdfWriter
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);base=root/'base.pdf';raw=root/'raw.pdf';out=root/'out.pdf'
            c=Canvas(str(base));c.drawString(72,720,'Visible text unchanged');c.save()
            writer=PdfWriter(clone_from=PdfReader(base))
            annotation=DictionaryObject({NameObject('/Type'):NameObject('/Annot'),NameObject('/Subtype'):NameObject('/Link'),NameObject('/Rect'):ArrayObject([NumberObject(0),NumberObject(0),NumberObject(10),NumberObject(10)]),NameObject('/A'):DictionaryObject({NameObject('/S'):NameObject('/Launch'),NameObject('/F'):TextStringObject('C:/Users/Test/AppData/Local/Temp/private.pdf')})})
            writer.add_annotation(0,annotation)
            with open(raw,'wb') as stream:writer.write(stream)
            writer.close()
            with patch.object(core,'REMOVE_LINKS',False):
                with self.assertRaisesRegex(ValueError,'local-file PDF link action'):core.clean_and_check(raw,out)
            self.assertFalse(out.exists())

    def test_owned_partial_cleanup_never_deletes_unowned_prefix_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);reference=root/'clean.pdf';target=root/'client.pdf';reference.write_bytes(b'complete bytes');target.write_bytes(b'complete')
            self.assertFalse(fileops.remove_owned_partial(target,reference,root/'missing.json',fileops.digest(reference)))
            self.assertEqual(target.read_bytes(),b'complete')


if __name__=='__main__':unittest.main()
