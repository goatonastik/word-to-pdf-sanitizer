"""Internal regression checks; creates samples only in a fresh temporary folder."""
import tempfile
from pathlib import Path
import subprocess
import sys
import unittest
from unittest.mock import patch

APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':
    APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))

import convert
from reports import report_rows
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas


def sample(path):
    d = Document()
    # A deliberately saved No Markup fixture, rather than Word's app default.
    view=OxmlElement('w:revisionView'); view.set(qn('w:markup'),'0')
    d.settings.element.append(view)
    d.core_properties.author = 'PRIVATE_AUTHOR_MARKER'
    d.core_properties.comments = 'PRIVATE_PROPERTY_MARKER'
    d.add_paragraph('Core document text must remain.')
    p = d.add_paragraph('Final version: ')
    for tag, text in [('del', 'DELETED_SECRET_MARKER'), ('ins', 'Approved text')]:
        revision = OxmlElement('w:' + tag)
        revision.set(qn('w:id'), '1' if tag == 'del' else '2')
        revision.set(qn('w:author'), 'PRIVATE_REVIEWER_MARKER')
        r = OxmlElement('w:r'); t = OxmlElement('w:delText' if tag == 'del' else 'w:t')
        t.text = text; r.append(t); revision.append(r); p._p.append(revision)
    run = d.add_paragraph().add_run('HIDDEN_SECRET_MARKER'); run.font.hidden = True
    comment_run = d.add_paragraph().add_run('Comment anchor stays.')
    d.add_comment(comment_run, 'COMMENT_SECRET_MARKER', author='Private Reviewer')
    d.sections[0].header.paragraphs[0].text = 'Header must remain'
    d.add_table(rows=2, cols=2).cell(0, 0).text = 'Table must remain'
    d.save(path)


class Checks(unittest.TestCase):
    def test_non_word_email_is_not_stat_inspected(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            name=('Delivery of documents '+('x'*200)+'.eml')
            # Simulate the exact failure: an enumerated non-input whose stat fails.
            def walk(*args,**kwargs): yield str(root),[],[name,'valid.DOCX']
            def inspect(path):
                if path.suffix=='.eml': raise FileNotFoundError(3,'long email path',str(path))
                return False
            audit={}
            with patch.object(convert.os,'walk',side_effect=walk),patch.object(convert,'is_link',side_effect=inspect) as link:
                files=list(convert.documents(root,audit))
            self.assertEqual(files,[root/'valid.DOCX'])
            self.assertEqual(link.call_count,1)
            self.assertEqual(audit['errors'],[])
            self.assertEqual(audit['otherFiles'],1)

    def test_scanner_walks_and_stats_extended_windows_paths(self):
        import os
        if os.name!='nt': self.skipTest('Windows-only path test')
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            nested=root/('Long directory '*5).strip()/('Another directory '*5).strip()
            import os
            os.makedirs(convert.native_path(nested))
            word=nested/('Long document '+('x'*90)+'.DOCX')
            email=nested/('Long email '+('x'*100)+'.eml')
            for path in (word,email):
                with open(convert.native_path(path),'wb') as stream: stream.write(b'test')
            self.assertGreater(len(str(word)),260)
            self.assertFalse(convert.is_link(word))
            audit={}; files=list(convert.documents(root,audit))
            self.assertEqual(files,[word])
            self.assertEqual(audit['errors'],[])
            self.assertEqual(audit['otherFiles'],1)

    def test_eligible_file_inspection_failure_is_a_scan_error(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            def walk(*args,**kwargs): yield str(root),[],['missing.docx']
            audit={}
            with patch.object(convert.os,'walk',side_effect=walk),patch.object(convert,'is_link',side_effect=FileNotFoundError(3,'missing Word input',str(root/'missing.docx'))):
                self.assertEqual(list(convert.documents(root,audit)),[])
            self.assertEqual(len(audit['errors']),1)
    def test_debug_paths_and_exceptions_without_document_contents(self):
        from debugtrace import DebugLog
        import json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); source=folder/'sample.docx'; source.write_text('PRIVATE_DOCUMENT_BODY_MARKER')
            log=folder/'debug.jsonl'; debug=DebugLog(log)
            debug.stage('copying document',{'source':source})
            try: raise FileNotFoundError(3,'simulated path error')
            except FileNotFoundError as error: debug.exception('copying document',error)
            rows=[json.loads(line) for line in log.read_text().splitlines()]
            self.assertEqual(rows[0]['paths']['source']['characters'],len(str(source)))
            self.assertTrue(rows[0]['paths']['source']['exists'])
            self.assertIn('FileNotFoundError',rows[1]['traceback'])
            self.assertNotIn('PRIVATE_DOCUMENT_BODY_MARKER',log.read_text())
            with patch('builtins.open',side_effect=PermissionError('simulated log lock')):
                debug.event('must not raise')
            self.assertEqual(debug.failures,1)

    def test_debug_option_can_be_disabled(self):
        import pipeline,json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root=folder/'Source'; root.mkdir()
            settings=folder/'settings.json'; settings.write_text(json.dumps({'debugLogging':False}))
            with patch.object(convert,'SETTINGS_FILE',settings),patch.object(convert,'DEBUG_OVERRIDE',None),patch.object(convert,'PROGRESS_FILE',folder/'progress.json'):
                pipeline.batch(root,folder/'Originals',preview=True)
            state=json.loads((folder/'progress.json').read_text())
            self.assertNotIn('debugLog',state)
            self.assertNotIn('wordDebugLog',state)
            self.assertNotIn('verificationDebugLog',state)
    def test_progress_lock_does_not_fail_conversion_and_recovers(self):
        import json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); path=folder/'progress.json'
            with patch.object(convert,'PROGRESS_FILE',path),patch.object(convert,'PROGRESS',{}):
                with patch.object(convert.os,'replace',side_effect=PermissionError(5,'simulated progress lock')),patch.object(convert.time,'sleep'):
                    convert.progress(done=1)
                self.assertEqual(convert.PROGRESS['done'],1)
                self.assertEqual(convert.PROGRESS['progressWriteWarnings'],1)
                convert.progress(done=2)
                self.assertEqual(json.loads(path.read_text())['done'],2)
                self.assertEqual(list(folder.glob('*.tmp')),[])

    def test_long_original_filename_gets_short_word_working_path(self):
        import pipeline
        workspace=Path(tempfile.gettempdir())/'word-pdf-export-test'
        source=Path('C:/Export/Starting')/('Long letter name '+('x'*175)+'.DOCX')
        result=pipeline.working_copy_path(workspace,source)
        self.assertLess(len(str(result)),230)
        self.assertNotEqual(result.name,source.name)
        self.assertEqual(result.suffix,'.docx')

    def test_recursive_scan_and_exact_existing_pdf_skip(self):
        import pipeline,json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root=folder/'Starting'; nested=root/'2'/'PROBATE'; nested.mkdir(parents=True)
            names=[root/'a.doc',root/'2'/'b.docm',nested/'3.DOCX']
            for path in names:
                path.write_bytes(b'unchanged original')
                path.with_suffix('.pdf').write_bytes(b'existing PDF')
            with patch.object(convert,'PROGRESS_FILE',folder/'progress.json'),patch.object(pipeline,'WordSession') as session:
                pipeline.batch(root,folder/'Originals')
                session.assert_not_called()
            state=json.loads((folder/'progress.json').read_text())
            self.assertEqual((state['total'],state['done'],state['skipped'],state['reviewAdded']),(3,3,3,0))
            self.assertEqual(state['scanFolders'],3)
            self.assertEqual(state['reviewCopyErrors'],0)
            self.assertIn('Existing PDF:',str(list(report_rows(state['report']))))
            self.assertFalse((folder/'#REVIEW'/'Starting').exists())
            self.assertEqual(names[-1].read_bytes(),b'unchanged original')
            self.assertEqual(names[-1].with_suffix('.pdf').read_bytes(),b'existing PDF')

    def test_old_skip_review_setting_is_ignored(self):
        import pipeline,json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root=folder/'Starting'; root.mkdir()
            source=root/'letter.docx'; source.write_bytes(b'original')
            source.with_suffix('.pdf').write_bytes(b'existing PDF')
            settings=folder/'settings.json'; settings.write_text(json.dumps({'flagAlreadyConvertedReview':True,'debugLogging':False}))
            with patch.object(convert,'SETTINGS_FILE',settings),patch.object(convert,'PROGRESS_FILE',folder/'progress.json'),patch.object(pipeline,'WordSession') as session:
                pipeline.batch(root,folder/'Originals')
                session.assert_not_called()
            state=json.loads((folder/'progress.json').read_text())
            self.assertEqual(state['reviewAdded'],0)
            self.assertFalse((folder/'#REVIEW'/'Starting').exists())
            self.assertNotIn('reviewReport',state)
            self.assertEqual(source.with_suffix('.pdf').read_bytes(),b'existing PDF')

    def test_inaccessible_subfolder_stops_before_conversion(self):
        import pipeline,json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root=folder/'Starting'; root.mkdir()
            def failed_walk(*args,**kwargs):
                kwargs['onerror'](PermissionError(13,'Access denied',str(root/'Denied')))
                yield str(root),[],[]
            with patch.object(convert,'PROGRESS_FILE',folder/'progress.json'),patch.object(convert.os,'walk',side_effect=failed_walk),patch.object(pipeline,'WordSession') as session:
                with self.assertRaisesRegex(RuntimeError,'scan was incomplete'):
                    pipeline.batch(root,folder/'Originals')
                session.assert_not_called()
            state=json.loads((folder/'progress.json').read_text())
            self.assertFalse(state['queueCompleted'])
            self.assertIn('Access denied',str(list(report_rows(state['scanReport']))))

    def test_long_review_path_and_timestamp_failure_preserve_verified_copy(self):
        import pipeline
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); source=folder/'source.docx'; source.write_bytes(b'original')
            review=folder/'Originals'/'REVIEW'
            relative=Path(*([('Long folder name '*3).strip()]*5))/'3.DOCX'
            audit={}
            with patch.object(pipeline.shutil,'copystat',side_effect=FileNotFoundError('simulated timestamp failure')):
                path,added=pipeline.review_copy(source,folder/'unused',review,relative,'queue',audit)
            self.assertTrue(added); self.assertGreater(len(path),260)
            with open(convert.native_path(path),'rb') as stream: self.assertEqual(stream.read(),b'original')
            self.assertIn('timestampWarning',audit)
            self.assertEqual(source.read_bytes(),b'original')

    def test_aborted_queue_retains_both_filename_flags(self):
        import pipeline,json
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory)
            def abort(*args):
                path=convert.RUN_LOGS.path(folder,'conversion-report','csv')
                path.write_text('unfinished report')
                convert.progress(report=str(path))
                raise RuntimeError('simulated queue abort')
            with patch.object(convert,'PROGRESS_FILE',folder/'progress.json'),patch.object(convert,'QUEUE_NAME','Root ['+folder.name+']'),patch.object(pipeline,'_batch',side_effect=abort):
                with self.assertRaises(RuntimeError): pipeline.batch(folder,folder/'archive')
            state=json.loads((folder/'progress.json').read_text())
            self.assertFalse(state['queueCompleted']); self.assertTrue(state['queueErrors'])
            self.assertTrue(Path(state['report']).stem.endswith('FAILED ERRORS'))
            self.assertTrue(Path(state['report']).exists())

    def test_log_status_suffixes_are_independent(self):
        from runlogs import RunLogs,status_suffix
        self.assertEqual(status_suffix(True,False),'')
        self.assertEqual(status_suffix(True,True),' ERRORS')
        self.assertEqual(status_suffix(False,False),' FAILED')
        self.assertEqual(status_suffix(False,True),' FAILED ERRORS')
        with tempfile.TemporaryDirectory() as directory:
            run=RunLogs('Client Folder [2026-11-09 14-00-00.000]')
            path=run.path(Path(directory),'privacy-report','csv')
            path.write_text('test')
            self.assertIn('FAILED',path.stem)
            mapping=run.finalize(False,True)
            renamed=Path(mapping[str(path)])
            self.assertTrue(renamed.stem.endswith('FAILED ERRORS'))
            self.assertEqual(renamed.read_text(),'test')
    def test_review_copy_deduplicates_and_keeps_versions(self):
        import pipeline
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); source=folder/'original.docx'; source.write_bytes(b'first')
            root=folder/'Originals'/'REVIEW'/'Source'
            relative=Path('Nested')/'original.docx'
            path,added=pipeline.review_copy(source,folder/'unused',root,relative,'batch1')
            self.assertTrue(added); self.assertEqual(Path(path).read_bytes(),b'first')
            again,added=pipeline.review_copy(source,folder/'unused',root,relative,'batch2')
            self.assertFalse(added); self.assertEqual(path,again)
            source.write_bytes(b'second')
            version,added=pipeline.review_copy(source,folder/'unused',root,relative,'batch3')
            self.assertTrue(added); self.assertEqual(Path(version).read_bytes(),b'second')
            self.assertEqual(Path(path).read_bytes(),b'first')
    def test_changed_appearance_is_withheld_and_reported(self):
        import pipeline
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); source=folder/'source.docx'; source.write_bytes(b'original')
            raw=folder/'raw.pdf'; clean=folder/'clean.pdf'; target=folder/'result.pdf'; archive=folder/'archive'/'source.docx'
            for path,x in ((raw,72),(Path(str(raw)+'.baseline.pdf'),96)):
                c=Canvas(str(path)); c.drawString(x,720,'Same text, changed position'); c.save()
            result=pipeline.verify_and_archive(source,target,archive,raw,clean,convert.digest(source),dict(pipeline.DEFAULTS,archive=str(folder/'archive')))
            self.assertEqual(result['status'],'FAILED; SOURCE RETAINED')
            self.assertIn('Failed',result['audit']['original_appearance'])
            difference=result['audit']['appearance_difference']
            self.assertEqual(difference['page'],1)
            self.assertEqual(difference['dpi'],144)
            self.assertIsNotNone(difference['changed_bounds'])
            self.assertFalse(target.exists()); self.assertFalse(archive.exists())
            self.assertEqual(source.read_bytes(),b'original')

    def test_pipeline_recovers_and_keeps_failed_original(self):
        import pipeline
        import json
        class FakeSession:
            def __init__(self, *args): self.ready = {'ok': True}
            def export(self, source, output, tick, original_name=None):
                if source.stem == 'b': raise TimeoutError('simulated export timeout')
                c=Canvas(str(output)); c.drawString(72,720,'Preserved text'); c.save()
                import shutil
                shutil.copyfile(output,str(output)+'.baseline.pdf')
                return {'timings': {'export': .01}}
            def close(self, **kwargs): pass
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory); root=folder/'root'; root.mkdir()
            for name in ('a','b','c'): (root/(name+'.docx')).write_bytes(b'original bytes')
            settings=folder/'settings.json'; settings.write_text(json.dumps({'queueLimit':1}))
            with patch.object(convert,'SETTINGS_FILE',settings), patch.object(convert,'PROGRESS_FILE',folder/'progress.json'), patch.object(pipeline,'WordSession',FakeSession):
                pipeline.batch(root,folder/'archive')
            state=json.loads((folder/'progress.json').read_text())
            self.assertEqual((state['done'],state['converted'],state['failed']),(3,2,1))
            self.assertEqual((root/'b.docx').read_bytes(),b'original bytes')
            self.assertFalse((root/'b.pdf').exists())
            self.assertEqual((folder/'archive'/'c.docx').read_bytes(),b'original bytes')
            self.assertEqual((folder/'#REVIEW'/'root'/'b.docx').read_bytes(),b'original bytes')
            self.assertEqual(state['reviewAdded'],1)
            self.assertTrue(Path(state['privacyReport']).stem.endswith('ERRORS'))
            self.assertNotIn('FAILED',Path(state['privacyReport']).stem)
            self.assertIn('simulated export timeout',Path(state['diagnostics']).read_text())

    def test_only_requested_conditions_trigger_review(self):
        import pipeline
        cases=[
            ('CONVERTED',{'revisionsAccepted':2},False),
            ('CONVERTED',{'commentsRemoved':2},False),
            ('CONVERTED',{'sourcePropertyTypesFound':'Inspection unavailable'},False),
            ('CONVERTED',{'sourcePropertyTypesFound':'Not inspected (legacy Word format)'},False),
            ('SKIPPED',{},False),
            ('CONVERTED',{'visibleRevisionsPreserved':True},True),
            ('CONVERTED',{'visibleCommentsPreserved':True},True),
            ('FAILED; SOURCE RETAINED',{},True),
            ('PDF CREATED; ARCHIVE FAILED',{},True),
        ]
        for status,audit,expected in cases:
            with self.subTest(status=status,audit=audit):
                result=dict(status=status,audit=audit,details='Test outcome')
                self.assertEqual(bool(pipeline.review_reasons(result)),expected)
        self.assertFalse(pipeline.review_reasons(dict(status='SKIPPED',skip_kind='name_collision',details='Collision')))

    def test_cleanup(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory); raw = folder / 'raw.pdf'; clean = folder / 'clean.pdf'
            c = Canvas(str(raw)); c.setAuthor('SECRET_AUTHOR'); c.drawString(72, 720, 'Keep this text'); c.save()
            self.assertEqual(convert.clean_and_check(raw, clean), 1)
            r = PdfReader(clean)
            self.assertFalse(r.metadata)
            self.assertIn('Keep this text', r.pages[0].extract_text())
            self.assertNotIn(b'SECRET_AUTHOR', clean.read_bytes())

    def test_roots(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for archive in [root, root / 'archive', root.parent]:
                with self.assertRaises(ValueError):
                    convert.validate_roots(root, archive)



def integration():
    folder = Path(tempfile.mkdtemp(prefix='word-pdf-integration-'))
    root = folder / 'root'; archive = folder / 'archive'
    (root / 'nested').mkdir(parents=True); archive.mkdir()
    source = root / 'nested' / 'Sample.docx'; sample(source)
    sample(root / 'Second.docx')
    original = convert.digest(source)
    convert.PROGRESS_FILE = folder / 'progress.json'
    print('Test folder:', folder, flush=True)
    from pipeline import batch
    batch(root, archive, timeout=90)
    target = source.with_suffix('.pdf')
    assert target.exists(), 'No output PDF; inspect report'
    assert not source.exists(), 'Source not archived; inspect report'
    assert convert.digest(archive / 'nested' / 'Sample.docx') == original
    r = PdfReader(target)
    text = '\n'.join(p.extract_text() for p in r.pages)
    for marker in ['Core document text must remain.', 'Approved text', 'Header must remain', 'Table must remain', 'Comment anchor stays.']:
        assert marker in text, (marker, text)
    for marker in ['DELETED_SECRET_MARKER', 'HIDDEN_SECRET_MARKER', 'COMMENT_SECRET_MARKER', 'PRIVATE_AUTHOR_MARKER', 'PRIVATE_REVIEWER_MARKER', 'PRIVATE_PROPERTY_MARKER']:
        assert marker not in text, (marker, text)
        assert marker.encode() not in target.read_bytes(), marker
    assert not r.metadata
    import json
    progress = json.loads(convert.PROGRESS_FILE.read_text())
    assert progress['finished'] and progress['done'] == 2 and progress['converted'] == 2
    assert progress['failed'] == 0 and progress['unprocessed'] == 0
    diagnostics = [json.loads(line) for line in Path(progress['diagnostics']).read_text().splitlines()]
    import csv
    rows=list(report_rows(progress['privacyReport']))
    assert len(rows)==2
    for row in rows:
        assert row['Comments removed']=='1', row
        assert int(row['Revisions accepted'])>=2, row
        assert row['Original final-view appearance'].startswith('Passed'), row
        assert row['Word properties export']=='Excluded'
        assert 'Author' in row['Source standard property types found'], row
    assert diagnostics[-1]['event'] == 'summary'
    assert (root / 'Second.pdf').exists(), 'Second document did not convert'
    assert (archive / 'Second.docx').exists(), 'Second document did not archive'
    assert len([r for r in diagnostics if r['event']=='word_session']) == 1
    verifier_pids={r['verifier_pid'] for r in diagnostics if r['event']=='document' and r['status']=='CONVERTED'}
    assert 1<=len(verifier_pids)<=progress['verifierWorkers'],verifier_pids
    assert all('export' in r['timings'] for r in diagnostics if r['event']=='document' and r['status']=='CONVERTED')
    print('PASS: persistent direct export, parallel reusable verifiers, privacy markers, timings and verified archiving.')


if __name__ == '__main__':
    if '--options-integration' in sys.argv:
        import json
        folder = Path(tempfile.mkdtemp(prefix='word-pdf-options-'))
        root = folder / 'Starting Folder'; root.mkdir()
        source = root / 'Sample.docx'; sample(source)
        settings = folder / 'settings.json'
        settings.write_text(json.dumps(dict(hideText=False, excludeProperties=False, excludeTags=False, stripMetadata=False, removeLinks=False, queueLimit=1)))
        result = subprocess.run([sys.executable, str(convert.HERE / 'convert.py'), '--root', str(root), '--archive', str(folder / 'Originals'), '--settings', str(settings), '--preserve-root'], capture_output=True, text=True)
        print(result.stdout, result.stderr)
        # Disabling removal must preserve hidden formatting, not reveal the text.
        assert not source.exists()
        assert (root / 'Sample.pdf').exists()
        assert (folder / 'Originals' / 'Starting Folder' / 'Sample.docx').exists()
        visible='\n'.join(p.extract_text() for p in PdfReader(root/'Sample.pdf').pages)
        assert 'HIDDEN_SECRET_MARKER' not in visible
        assert 'CONVERTED:' in result.stdout
        print('PASS: disabled hidden-text removal keeps saved hidden formatting and appearance.')
    elif '--integration' in sys.argv:
        integration()
    else:
        unittest.main()
