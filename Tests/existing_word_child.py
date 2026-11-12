"""Helper invoked by Test-ExistingWord.ps1 while its unsaved document is open."""
from pathlib import Path
import json,subprocess,sys,tempfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import convert
import pipeline
from docx import Document

def main():
    user_pid=int(sys.argv[1]);folder=Path(tempfile.mkdtemp(prefix='word-pdf-existing-word-'))
    root=folder/'Starting';root.mkdir();originals=folder/'Originals'
    source=root/'Sample.docx';doc=Document();doc.add_paragraph('Generated conversion fixture.');doc.save(source)
    expected=convert.digest(source);progress=folder/'progress.json'
    result=subprocess.run([sys.executable,str(convert.HERE/'convert.py'),'--root',str(root),'--archive',str(originals),'--preserve-root','--progress',str(progress)],capture_output=True,text=True,encoding='utf-8',errors='replace')
    (folder/'console.log').write_text(result.stdout+'\n'+result.stderr,encoding='utf-8')
    assert result.returncode==0,result.stderr
    state=json.loads(progress.read_text());assert state['converted']==1,state
    archived=originals/'Starting'/'Sample.docx';assert convert.digest(archived)==expected
    events=[json.loads(line) for line in Path(state['diagnostics']).read_text().splitlines()]
    owned=next(row['ready'] for row in events if row.get('event')=='word_session')
    assert owned['pid']!=user_pid and owned['existingWordCount']>=1,owned
    converted=next(row for row in events if row.get('event')=='document' and row.get('status')=='CONVERTED')
    assert converted['audit']['wordGuardChecks']>=4,converted['audit']
    assert converted['timings']['word_guard_checks']>=0
    workspace=folder/'forced-worker';workspace.mkdir();settings=folder/'settings.json';settings.write_text(json.dumps(pipeline.DEFAULTS))
    session=pipeline.WordSession(workspace,settings,90)
    assert session.ready['pid']!=user_pid and session.ready['existingWordCount']>=1
    session.close(force=True)
    assert not (workspace/'options.json').exists(),'Forced recovery did not restore settings.'
    print('PROBE_RESULT: '+json.dumps(dict(normalConversion=True,archiveVerified=True,forcedRecovery=True,userPid=user_pid,conversionPid=owned['pid'],folder=str(folder))))

if __name__=='__main__':main()
