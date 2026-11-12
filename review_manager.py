"""Persistent, checksum-safe actions for the desktop detailed REVIEW window."""
import argparse
import csv
from datetime import datetime, timezone
import hashlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile

import convert as core
import maintenance
from fileops import remove_matching
from reports import build_excel,report_rows


DECISION_ROOT=core.HERE/'Logs'/maintenance.RECORDS/'review-decisions'
DECISION_HEADERS=('Review decision','Decision time','Decision result','Decision history')


def decision_path(report):
    key=hashlib.sha256(str(Path(report).resolve()).casefold().encode()).hexdigest()
    return DECISION_ROOT/(key+'.json')


def entry_id(document,pdf):
    value=str(Path(document).resolve()).casefold()+'|'+str(Path(pdf).resolve()).casefold()
    return hashlib.sha256(value.encode()).hexdigest()


def load_decisions(report):
    path=decision_path(report)
    if not path.exists():return {'version':1,'report':str(Path(report).resolve()),'entries':{}}
    with open(core.native_path(path),encoding='utf-8') as stream:value=json.load(stream)
    if Path(value['report']).resolve()!=Path(report).resolve():raise ValueError('Review decision record does not match this report.')
    value.setdefault('entries',{})
    return value


def save_decisions(report,value):
    path=decision_path(report);os.makedirs(core.native_path(path.parent),exist_ok=True)
    temp=path.with_suffix('.tmp')
    with open(core.native_path(temp),'w',encoding='utf-8') as stream:
        json.dump(value,stream,ensure_ascii=False,indent=2);stream.flush();os.fsync(stream.fileno())
    os.replace(core.native_path(temp),core.native_path(path))


def review_rows(report,identity_report):
    decisions=load_decisions(identity_report)['entries'];result=[]
    for row in report_rows(report):
        reasons=row.get('Reasons for review') or row.get('Review reasons') or ''
        if not reasons:continue
        document=row.get('Document') or row.get('Source') or ''
        pdf=row.get('PDF') or (str(Path(document).with_suffix('.pdf')) if document else '')
        if not document or not pdf:continue
        identifier=entry_id(document,pdf);state=decisions.get(identifier,{})
        original=row.get('Review copy') or row.get('Archive') or document
        decision=state.get('decision','')
        decision_result=state.get('result','')
        result.append(dict(id=identifier,document=document,pdf=pdf,original=original,
                           originalLabel=Path(original).name,convertedLabel=Path(pdf).name,
                           reason=reasons,decision=decision,decisionTime=state.get('time',''),
                           decisionResult=decision_result,decisionDisplay=((decision+': ') if decision else '')+decision_result,
                           canDecide=not bool(decision),canUndo=bool(decision)))
    return result


def set_state(report,identifier,decision,result):
    values=load_decisions(report);states=values['entries'];previous=states.get(identifier,{})
    history=list(previous.get('history',[]));now=datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')
    if previous.get('decision') or previous.get('result'):
        history.append(dict(decision=previous.get('decision',''),time=previous.get('time',''),result=previous.get('result','')))
    states[identifier]=dict(decision=decision,time=now,result=result,history=history)
    save_decisions(report,values)
    return states[identifier]


def revert_document(document,pdf,root,originals):
    document,pdf=Path(document).resolve(),Path(pdf).resolve()
    plan=maintenance.plan_restore(root,originals)
    matches=[entry for entry in plan['entries'] if Path(entry['source']).resolve()==document]
    if matches:
        chosen=matches[0]
        if chosen.get('blocked'):raise ValueError(chosen['blocked'])
        rows=maintenance.restore(dict(action='restore',root=plan['root'],originals=plan['originals'],entries=[chosen]),False)
        if not rows or rows[0][1]!='RESTORED':raise ValueError(rows[0][2] if rows else 'Restoration produced no result.')
        return rows[0][2]
    if not os.path.isfile(core.native_path(document)):
        raise ValueError('No verified archived original or retained source is available to restore.')
    if os.path.isfile(core.native_path(pdf)):
        marker=maintenance.incomplete_record_path(document)
        if not os.path.isfile(core.native_path(marker)):
            raise ValueError('The PDF has no checksum record proving that this program owns it; it was retained.')
        with open(core.native_path(marker),encoding='utf-8') as stream:proof=json.load(stream)
        if Path(proof['source']).resolve()!=document or Path(proof['pdf']).resolve()!=pdf:
            raise ValueError('Incomplete-conversion record paths do not match this document.')
        remove_matching(pdf,proof['pdf_sha256'])
        maintenance.clear_incomplete_conversion(document)
    return 'Original was already retained; matching program-owned PDF was removed if present.'


def reconvert_document(document,root,originals,settings):
    command=[sys.executable,'-X','utf8','-u',str(core.HERE/'convert.py'),'--root',str(root),'--archive',str(originals),
             '--settings',str(settings),'--preserve-root','--only',str(document)]
    completed=subprocess.run(command,cwd=core.HERE,text=True,encoding='utf-8',errors='replace',capture_output=True)
    if completed.returncode:
        message=(completed.stderr or completed.stdout or 'One-document conversion failed.').strip()
        raise RuntimeError(message[-4000:])
    pdf=Path(document).with_suffix('.pdf')
    archive=Path(originals)/Path(root).name/Path(document).resolve().relative_to(Path(root).resolve())
    if not os.path.isfile(core.native_path(pdf)) or not os.path.isfile(core.native_path(archive)) or os.path.exists(core.native_path(document)):
        raise RuntimeError('Reconversion did not finish with a verified PDF and archived original.')
    return 'Undo completed; the document was converted and archived again and is awaiting a new decision.'


def perform_action(args):
    values=load_decisions(args.report);previous=values['entries'].get(args.id,{})
    if args.action=='accept':
        state=set_state(args.report,args.id,'Accepted','Conversion accepted; files were left unchanged.')
    elif args.action=='revert':
        details=revert_document(args.document,args.pdf,args.root,args.originals)
        state=set_state(args.report,args.id,'Reverted',details)
    elif args.action=='undo':
        if not previous.get('decision'):raise ValueError('This entry has no decision to undo.')
        if previous['decision']=='Reverted':details=reconvert_document(args.document,args.root,args.originals,args.settings)
        else:details='Accepted decision cleared; files were already unchanged.'
        state=set_state(args.report,args.id,'',details)
    else:raise ValueError('Unknown review action.')
    return dict(decision=state['decision'],decisionTime=state['time'],decisionResult=state['result'],
                decisionDisplay=((state['decision']+': ') if state['decision'] else '')+state['result'],
                canDecide=not bool(state['decision']),canUndo=bool(state['decision']))


def exclusive_snapshot(path):
    """Read while denying other opens on Windows; caller closes before replacement."""
    path=Path(path)
    if os.name!='nt':
        stream=open(path,'rb');return stream
    import ctypes,msvcrt
    from ctypes import wintypes
    create=ctypes.windll.kernel32.CreateFileW
    create.argtypes=[wintypes.LPCWSTR,wintypes.DWORD,wintypes.DWORD,wintypes.LPVOID,wintypes.DWORD,wintypes.DWORD,wintypes.HANDLE]
    create.restype=wintypes.HANDLE
    handle=create(str(path),0x80000000,0,None,3,0x80,None)
    if handle==wintypes.HANDLE(-1).value:
        raise PermissionError(ctypes.WinError(ctypes.get_last_error()))
    return os.fdopen(msvcrt.open_osfhandle(handle,os.O_RDONLY),'rb')


def apply_decisions(report):
    report=Path(report);decisions=load_decisions(report)['entries']
    handle,journal=tempfile.mkstemp(prefix='word-pdf-decisions-',suffix='.csvlog');os.close(handle)
    candidate=report.with_name('.'+report.name+'.review-update-'+os.urandom(6).hex()+'.xlsx')
    snapshot=Path(journal).with_suffix('.snapshot.xlsx')
    try:
        with exclusive_snapshot(report) as locked:
            with open(snapshot,'wb') as target:
                while True:
                    block=locked.read(1024*1024)
                    if not block:break
                    target.write(block)
                target.flush();os.fsync(target.fileno())
            rows=list(report_rows(snapshot))
            headers=list(rows[0]) if rows else []
            for header in DECISION_HEADERS:
                if header not in headers:headers.append(header)
            with open(journal,'w',newline='',encoding='utf-8-sig') as stream:
                writer=csv.DictWriter(stream,fieldnames=headers);writer.writeheader()
                for row in rows:
                    document=row.get('Document') or row.get('Source') or ''
                    pdf=row.get('PDF') or (str(Path(document).with_suffix('.pdf')) if document else '')
                    state=decisions.get(entry_id(document,pdf),{}) if document and pdf else {}
                    row['Review decision']=state.get('decision','') or ('Awaiting decision' if state.get('result') else '')
                    row['Decision time']=state.get('time','')
                    row['Decision result']=state.get('result','')
                    row['Decision history']=' | '.join(f"{h.get('time','')} {h.get('decision','')}: {h.get('result','')}" for h in state.get('history',[]))
                    writer.writerow(row)
                stream.flush();os.fsync(stream.fileno())
            build_excel(journal,candidate,'Detailed Review')
        os.replace(core.native_path(candidate),core.native_path(report))
    finally:
        for path in (Path(journal),snapshot,candidate):
            try:
                if path.exists():path.unlink()
            except OSError:pass


def main():
    parser=argparse.ArgumentParser()
    sub=parser.add_subparsers(dest='command',required=True)
    listing=sub.add_parser('list');listing.add_argument('--report',type=Path,required=True);listing.add_argument('--identity-report',type=Path,required=True)
    action=sub.add_parser('action');action.add_argument('--report',type=Path,required=True);action.add_argument('--id',required=True)
    action.add_argument('--action',choices=('accept','revert','undo'),required=True);action.add_argument('--document',type=Path,required=True)
    action.add_argument('--pdf',type=Path,required=True);action.add_argument('--root',type=Path,required=True);action.add_argument('--originals',type=Path,required=True);action.add_argument('--settings',type=Path,required=True)
    apply=sub.add_parser('apply');apply.add_argument('--report',type=Path,required=True)
    args=parser.parse_args()
    if args.command=='list':result=review_rows(args.report,args.identity_report)
    elif args.command=='action':result=perform_action(args)
    else:apply_decisions(args.report);result={'updated':str(args.report)}
    print(json.dumps(result,ensure_ascii=False))


if __name__=='__main__':
    try:main()
    except Exception as error:
        print(type(error).__name__+': '+str(error),file=sys.stderr);sys.exit(1)
