import csv
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
from docx import Document
from pypdf import PdfReader
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert
from reports import report_rows
import pipeline
from test_converter import sample

def main():
    import argparse
    parser=argparse.ArgumentParser()
    parser.add_argument('--legacy-fixture',type=Path,help='Optional .doc fixture with no comments or revisions')
    args=parser.parse_args()
    folder=Path(tempfile.mkdtemp(prefix='word-pdf-savedviews-'))
    source=folder/'Source'; originals=folder/'Originals'
    sub=source/'Nested'; sub.mkdir(parents=True)
    for name in ('Hidden','Visible','Simple','Original'): sample(sub/(name+'.docx'))
    doc=Document(); doc.add_paragraph('Clean document with no revisions.'); doc.save(sub/'Clean.docx')
    doc=Document(); run=doc.add_paragraph().add_run('Visible comment anchor.')
    doc.add_comment(run,'COMMENT_SECRET_MARKER',author='Private Reviewer'); doc.save(sub/'CommentsVisible.docx')
    print('Fixture folder:',folder,flush=True)
    subprocess.run(['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(convert.HERE/'Tests'/'Test-SavedViews.ps1'),'-Folder',str(sub)],check=True)
    # Let Word finish its ordinary shutdown; do not terminate other instances.
    time.sleep(1)
    hashes={p.name:convert.digest(p) for p in sub.glob('*.docx')}
    legacy_hash=None
    if args.legacy_fixture:
        import shutil
        shutil.copyfile(args.legacy_fixture,sub/'Legacy.doc')
        legacy_hash=convert.digest(args.legacy_fixture)
        hashes['Legacy.doc']=legacy_hash
    progress=folder/'progress.json'
    subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(source),'--archive',str(originals),'--preserve-root','--progress',str(progress)],check=True)
    state=json.loads(progress.read_text())
    for key in ('debugLog','wordDebugLog','verificationDebugLog'):
        assert Path(state[key]).exists(),(key,state)
        assert Path(state[key]).stat().st_size>0,(key,state)
        data=Path(state[key]).read_text(encoding='utf-8-sig')
        assert 'PRIVATE_AUTHOR_MARKER' not in data
        assert 'COMMENT_SECRET_MARKER' not in data
        assert 'DELETED_SECRET_MARKER' not in data
    word_events=[json.loads(line) for line in Path(state['wordDebugLog']).read_text(encoding='utf-8-sig').splitlines()]
    assert any(event['stage']=='exporting prepared PDF' for event in word_events)
    verify_events=[json.loads(line) for line in Path(state['verificationDebugLog']).read_text().splitlines()]
    assert any(event.get('stage')=='copying original into archive' for event in verify_events)
    assert Path(state['privacyReport']).name.startswith('Source ['),state
    assert 'FAILED' not in Path(state['privacyReport']).stem,state
    assert 'ERRORS' not in Path(state['privacyReport']).stem,state
    assert Path(state['reviewReport']).is_relative_to(folder/'#REVIEW'),state
    review_rows=list(report_rows(state['reviewReport']))
    rows={Path(row['Document']).stem:row for row in report_rows(state['privacyReport'])}
    for name,row in rows.items(): print(name,row['Outcome'],row['Revision handling'],row['Revisions accepted'],flush=True)
    if legacy_hash:
        legacy=rows['Legacy']
        assert legacy['Outcome']=='CONVERTED',legacy
        assert legacy['Original final-view appearance'].startswith('Passed'),legacy
        assert legacy['Review reasons']=='',legacy
        assert 'Author' in legacy['Source standard property types found'],legacy
        assert int(legacy['Source custom properties found'])>=0,legacy
        assert convert.digest(originals/'Source'/'Nested'/'Legacy.doc')==legacy_hash
        assert convert.digest(args.legacy_fixture)==legacy_hash
        assert not (folder/'#REVIEW'/'Source'/'Nested'/'Legacy.doc').exists()
        print('PASS: legacy Word inventory, mixed-session appearance, archive bytes and source attachment unchanged.',flush=True)
    hidden=rows['Hidden']; visible=rows['Visible']; simple=rows['Simple']
    assert hidden['Outcome']=='CONVERTED',hidden
    assert int(hidden['Revisions accepted'])>=2,hidden
    assert visible['Outcome']=='CONVERTED',visible
    assert visible['Revisions accepted']=='0',visible
    assert simple['Revisions accepted']=='0',simple
    # Word may reopen a saved Original view as No Markup; follow the state
    # Word actually restores, rather than assuming an unsaved UI state persists.
    expected_review={'Visible','Simple'}
    if rows['Original']['Outcome']!='CONVERTED' or rows['Original']['Revision handling']=='Saved revision display preserved': expected_review.add('Original')
    assert len(review_rows)==len(expected_review),review_rows
    for name in expected_review:
        review=folder/'#REVIEW'/'Source'/'Nested'/(name+'.docx')
        assert review.exists(),(name,state)
        assert convert.digest(review)==hashes[name+'.docx']
    assert not (folder/'#REVIEW'/'Source'/'Nested'/'Clean.docx').exists()
    assert rows['CommentsVisible']['Outcome']=='CONVERTED',rows['CommentsVisible']
    assert rows['CommentsVisible']['Comments removed']=='1'
    assert not (folder/'#REVIEW'/'Source'/'Nested'/'CommentsVisible.docx').exists()
    for name in ('Hidden','Visible'):
        archived=originals/'Source'/'Nested'/(name+'.docx')
        assert convert.digest(archived)==hashes[name+'.docx']
    for name in ('Hidden','Visible','Simple','Original'):
        assert rows[name]['Comments removed']=='1',rows[name]
        pdf=sub/(name+'.pdf')
        if pdf.exists():
            text='\n'.join(page.extract_text() for page in PdfReader(pdf).pages)
            assert 'COMMENT_SECRET_MARKER' not in text,(name,text)
    text='\n'.join(p.extract_text() for p in PdfReader(sub/'Visible.pdf').pages)
    assert 'DELETED_SECRET_MARKER' in text,text
    text='\n'.join(p.extract_text() for p in PdfReader(sub/'Hidden.pdf').pages)
    assert 'DELETED_SECRET_MARKER' not in text,text
    assert state['reviewCopyErrors']==0,state
    assert not (folder/'#REVIEW'/'Source'/'Nested'/'Hidden.docx').exists()
    assert state['reviewAdded']==len(expected_review),state
    print('PASS: hidden accepted, visible/simple preserved, clean not flagged, review structure and original bytes preserved.')
    # Detailed-review choices persist in the selected workbook. Revert uses the
    # verified restore record; Undo reconverts only that document.
    import review_manager
    review_manager.DECISION_ROOT=folder/'review-decisions'
    detailed=review_manager.review_rows(state['reviewReport'],state['reviewReport'])
    by_name={Path(row['document']).stem:row for row in detailed}
    action_settings=folder/'review-action-settings.json';action_settings.write_text('{}')
    from types import SimpleNamespace
    accepted=by_name['Simple']; action=SimpleNamespace(report=Path(state['reviewReport']),id=accepted['id'],action='accept',document=Path(accepted['document']),pdf=Path(accepted['pdf']),root=source,originals=originals,settings=action_settings)
    assert review_manager.perform_action(action)['decision']=='Accepted'
    action.action='undo';assert review_manager.perform_action(action)['decision']==''
    reverted=by_name['Visible'];action.id=reverted['id'];action.document=Path(reverted['document']);action.pdf=Path(reverted['pdf']);action.action='revert'
    assert review_manager.perform_action(action)['decision']=='Reverted'
    assert action.document.exists() and not action.pdf.exists() and not (originals/'Source'/'Nested'/'Visible.docx').exists()
    action.action='undo';undo=review_manager.perform_action(action)
    assert undo['decision']=='' and not action.document.exists() and action.pdf.exists() and (originals/'Source'/'Nested'/'Visible.docx').exists()
    review_manager.apply_decisions(Path(state['reviewReport']))
    detailed_row={Path(row['Document']).stem:row for row in report_rows(state['reviewReport'])}
    assert detailed_row['Visible']['Review decision']=='Awaiting decision',detailed_row['Visible']
    assert 'converted and archived again' in detailed_row['Visible']['Decision result'],detailed_row['Visible']
    print('PASS: detailed Accept, Revert and Undo persist; Undo reconverts only the selected document.',flush=True)
    # Turning preservation off explicitly requests clean final-view output.
    clean_root=folder/'DiscardAll'; clean_root.mkdir()
    import shutil
    shutil.copyfile(originals/'Source'/'Nested'/'Visible.docx',clean_root/'Visible.docx')
    settings=folder/'discard-settings.json'; settings.write_text(json.dumps({'preserveVisibleRevisions':False}))
    second_progress=folder/'discard-progress.json'
    subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(clean_root),'--archive',str(originals),'--preserve-root','--settings',str(settings),'--progress',str(second_progress)],check=True)
    second=json.loads(second_progress.read_text())
    row=next(report_rows(second['privacyReport']))
    assert row['Outcome']=='CONVERTED',row
    assert int(row['Revisions accepted'])>=2,row
    assert 'discarded' in row['Revision handling'],row
    assert 'DELETED_SECRET_MARKER' not in '\n'.join(p.extract_text() for p in PdfReader(clean_root/'Visible.pdf').pages)
    assert not (folder/'#REVIEW'/'DiscardAll'/'Visible.docx').exists()
    assert second['reviewFlagged']==0,second
    print('PASS: disabled preservation discards visible revision history, exports final text, and logs changes without review copies.')
    comments_root=folder/'KeepComments'; comments_root.mkdir()
    shutil.copyfile(originals/'Source'/'Nested'/'CommentsVisible.docx',comments_root/'CommentsVisible.docx')
    comments_settings=folder/'comments-settings.json'; comments_settings.write_text(json.dumps({'removeAllComments':False}))
    comments_progress=folder/'comments-progress.json'
    subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(comments_root),'--archive',str(originals),'--preserve-root','--settings',str(comments_settings),'--progress',str(comments_progress)],check=True)
    comments_state=json.loads(comments_progress.read_text())
    assert comments_state['converted']==1,comments_state
    assert comments_state['reviewAdded']==1,comments_state
    comment_row=next(report_rows(comments_state['reviewReport']))
    assert 'Visible comments preserved' in comment_row['Reasons for review'],comment_row
    print('PASS: removed visible comments do not trigger review; preserved visible comments do.',flush=True)
    combined_root=folder/'DiscardKeepComments';combined_root.mkdir()
    shutil.copyfile(originals/'Source'/'Nested'/'Visible.docx',combined_root/'Visible.docx')
    combined_settings=folder/'combined-settings.json';combined_settings.write_text(json.dumps({'removeAllComments':False,'preserveVisibleRevisions':False}))
    combined_progress=folder/'combined-progress.json'
    subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(combined_root),'--archive',str(originals),'--preserve-root','--settings',str(combined_settings),'--progress',str(combined_progress)],check=True)
    combined=json.loads(combined_progress.read_text())
    assert combined['converted']==1,combined
    combined_row=next(report_rows(combined['privacyReport']))
    assert combined_row['Comments removed']=='0',combined_row
    assert int(combined_row['Revisions accepted'])>=2,combined_row
    assert combined['reviewAdded']==1,combined
    assert 'Visible comments preserved' in combined_row['Review reasons'],combined_row
    print('PASS: discarding revisions preserves visible comments independently when comment removal is off.',flush=True)
    long_root=folder/'Long Names'; long_root.mkdir()
    long_name='Letter '+('x'*185)+'.docx'
    long_source=long_root/long_name
    sample(long_source)
    long_hash=convert.digest(long_source)
    long_progress=folder/'long-progress.json'
    subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(long_root),'--archive',str(originals),'--preserve-root','--progress',str(long_progress)],check=True)
    long_state=json.loads(long_progress.read_text())
    assert long_state['converted']==1,long_state
    assert (long_root/Path(long_name).with_suffix('.pdf')).exists()
    assert convert.digest(originals/'Long Names'/long_name)==long_hash
    events=[json.loads(line) for line in Path(long_state['diagnostics']).read_text().splitlines()]
    assert any(event.get('event')=='export_paths' and event['workingFilenameShortened'] for event in events)
    print('PASS: long original names use short working paths, retain final filenames and archive original bytes.')
    import maintenance
    restore_plan=maintenance.plan_restore(folder/'Source',originals)
    assert restore_plan['ready']==len(hashes),restore_plan
    restored=maintenance.restore(restore_plan)
    assert all(row[1]=='RESTORED' for row in restored),restored
    for name,expected_hash in hashes.items():
        assert convert.digest(sub/name)==expected_hash
        assert not (sub/Path(name).with_suffix('.pdf')).exists()
    assert (folder/'#REVIEW'/'Source'/'Nested'/'Visible.docx').exists()
    print('PASS: real conversion records restore original Word bytes, delete only paired PDFs, and retain #REVIEW copies.')

if __name__=='__main__': main()
