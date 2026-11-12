"""Plan and perform explicit restoration/log cleanup; never infer PDF ownership."""
import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
import convert as core
from fileops import copy_verified,remove_readonly,remove_matching
from reports import report_rows,build_excel

RECORDS='.sanitizer-records'


def saved_record_path(source,archive):
    key=hashlib.sha256(str(Path(source).resolve()).casefold().encode()).hexdigest()
    return Path(archive)/RECORDS/(key+'.json')


def save_restore_record(source,pdf,archive,source_hash,pdf_hash,config,audit=None):
    folder=Path(config['archive'])/RECORDS
    os.makedirs(core.native_path(folder),exist_ok=True)
    path=saved_record_path(source,config['archive'])
    value=dict(source=str(source),pdf=str(pdf),archive=str(archive),source_sha256=source_hash,pdf_sha256=pdf_hash,audit=audit or {})
    temp=path.with_suffix('.tmp')
    with open(core.native_path(temp),'w',encoding='utf-8') as stream:
        json.dump(value,stream);stream.flush();os.fsync(stream.fileno())
    os.replace(core.native_path(temp),core.native_path(path))


def incomplete_record_path(source):
    key=hashlib.sha256(str(Path(source).resolve()).casefold().encode()).hexdigest()
    return core.HERE/'Logs'/RECORDS/'incomplete-conversions'/(key+'.json')


def clear_incomplete_conversion(source):
    path=incomplete_record_path(source)
    if os.path.exists(core.native_path(path)):os.unlink(core.native_path(path))


def remember_incomplete_conversion(source,pdf,archive,reason):
    if not os.path.isfile(core.native_path(pdf)):return
    path=incomplete_record_path(source);os.makedirs(core.native_path(path.parent),exist_ok=True)
    value=dict(source=str(Path(source).resolve()),pdf=str(Path(pdf).resolve()),pdf_sha256=core.digest(pdf),archive=str(Path(archive).resolve()),archive_sha256=core.digest(archive) if os.path.isfile(core.native_path(archive)) else None,reason=str(reason))
    temporary=path.with_suffix('.tmp')
    with open(core.native_path(temporary),'w',encoding='utf-8') as stream:json.dump(value,stream);stream.flush();os.fsync(stream.fileno())
    os.replace(core.native_path(temporary),core.native_path(path))


def incomplete_conversion_reason(source,pdf,archive):
    path=incomplete_record_path(source)
    if not os.path.exists(core.native_path(path)):return None
    with open(core.native_path(path),encoding='utf-8') as stream:value=json.load(stream)
    if Path(value['source']).resolve()!=Path(source).resolve() or Path(value['pdf']).resolve()!=Path(pdf).resolve():raise ValueError('Incomplete-conversion record paths do not match.')
    unchanged_pdf=os.path.isfile(core.native_path(pdf)) and core.digest(pdf)==value['pdf_sha256']
    unchanged_archive=Path(value['archive']).resolve()==Path(archive).resolve() and value.get('archive_sha256') and os.path.isfile(core.native_path(archive)) and core.digest(archive)==value['archive_sha256']
    if unchanged_pdf or unchanged_archive:return 'A previous conversion did not complete. Its retained output has not changed: '+value['reason']
    clear_incomplete_conversion(source)
    return None


def roots(root,originals):
    root,originals=core.validate_roots(Path(root),Path(originals))
    if root.name.casefold() in ('#review','#originals'): raise ValueError('Select a document folder as Starting directory.')
    if originals.name.casefold()=='#review':raise ValueError('Choose an Originals directory distinct from #REVIEW.')
    archive=originals/root.name
    review=originals.parent/'#REVIEW'/root.name
    core.validate_review_root(root,archive,review)
    return root,originals,archive,review


QUEUE_LOG=re.compile(r'^.+ \[[^\[\]]+\] - (?:conversion-report|review-report|privacy-report|scan-report|debug|word-debug|word-errors|verification-debug|diagnostics|output|console-errors|progress|interface-errors)(?: FAILED)?(?: ERRORS)?\.(?:csv|xlsx|csvlog|jsonl|log|json)$',re.I)


def is_generated_log(path):
    name=path.name
    if QUEUE_LOG.fullmatch(name) or re.fullmatch(r'startup-failure-\d{8}-\d{6}-\d{3}(?: FAILED)?(?: ERRORS)?\.log',name,re.I) or name.lower()=='interface-unhandled-errors.log':return True
    if path.suffix.lower()=='.csv' and 'conversion-report' in name.lower():
        try:
            with open(core.native_path(path),newline='',encoding='utf-8-sig') as stream:
                return {'Source','PDF','Archive','Status','Details'}.issubset(next(csv.reader(stream),[]))
        except (OSError,csv.Error):return False
    return False


def log_files(base):
    if not base.exists(): return []
    result=[]
    for directory,dirs,files in os.walk(core.native_path(base),followlinks=False):
        parent=core.display_path(directory)
        dirs[:]=[d for d in dirs if d!=RECORDS and not core.is_link(parent/d)]
        for name in files:
            path=parent/name; lower=name.lower()
            if core.is_link(path): continue
            if is_generated_log(path):result.append(path)
    return result


def valid_entry(value,root,archive):
    source=Path(value['source']).resolve(); pdf=Path(value['pdf']).resolve(); saved=Path(value['archive']).resolve()
    relative=source.relative_to(root)
    if source.suffix.lower() not in core.EXTENSIONS or pdf!=source.with_suffix('.pdf') or saved!=archive/relative:
        raise ValueError('Conversion record does not match selected folders.')
    if not core.inside(saved,archive.resolve()): raise ValueError('Archive path is redirected outside selected originals.')
    digest=value.get('source_sha256','')
    if not re.fullmatch('[0-9a-fA-F]{64}',digest): raise ValueError('Record lacks original checksum.')
    return dict(value,source=str(source),pdf=str(pdf),archive=str(saved))


def conversion_records(root,archive):
    entries={}
    store=archive/RECORDS
    if store.exists():
        for path in store.glob('*.json'):
            try:
                with open(core.native_path(path),encoding='utf-8') as stream: value=valid_entry(json.load(stream),root,archive)
                entries[value['source']]=value
            except (ValueError,KeyError,OSError): continue
    reports=[p for p in log_files(archive) if 'conversion-report' in p.name.lower()]
    legacy=store/'legacy-reports'
    if legacy.exists(): reports+=[p for p in legacy.iterdir() if p.suffix.lower() in ('.csv','.xlsx','.csvlog')]
    for path in reports:
        try:
            for row in report_rows(path):
                if row.get('Status')!='CONVERTED': continue
                match=re.search(r'SHA256\s+([0-9a-f]{64})',row.get('Details',''),re.I)
                if not match: continue
                try:
                    value=valid_entry(dict(source=row['Source'],pdf=row['PDF'],archive=row['Archive'],source_sha256=match[1],pdf_sha256=None),root,archive)
                    entries.setdefault(value['source'],value)
                except (ValueError,KeyError): continue
        except (OSError,csv.Error,ValueError,ET.ParseError,zipfile.BadZipFile): continue
    return list(entries.values())


def restore_issue(entry):
    source,pdf,saved=map(Path,(entry['source'],entry['pdf'],entry['archive']))
    if not os.path.isfile(core.native_path(saved)): return 'Archived original is missing.'
    if core.digest(saved)!=entry['source_sha256']: return 'Archived original changed since conversion.'
    if os.path.exists(core.native_path(source)) and core.digest(source)!=entry['source_sha256']: return 'Different Word document already exists in Starting.'
    if os.path.exists(core.native_path(pdf)) and entry.get('pdf_sha256') and core.digest(pdf)!=entry['pdf_sha256']:
        return 'PDF changed since conversion; will not delete it.'
    return ''


def plan_restore(root,originals):
    root,originals,archive,review=roots(root,originals)
    candidates=conversion_records(root,archive);entries=[];already=0
    for entry in candidates:
        try:
            source,pdf,saved=map(Path,(entry['source'],entry['pdf'],entry['archive']))
            if not os.path.exists(core.native_path(saved)) and not os.path.exists(core.native_path(pdf)) and os.path.isfile(core.native_path(source)) and core.digest(source)==entry['source_sha256']:
                already+=1;continue
            entry['blocked']=restore_issue(entry)
        except OSError as error:entry['blocked']=str(error)
        entries.append(entry)
    ready=[e for e in entries if not e['blocked']]
    return dict(action='restore',root=str(root),originals=str(originals),entries=entries,ready=len(ready),blocked=len(entries)-len(ready),alreadyRestored=already,legacy=sum(not e.get('pdf_sha256') for e in ready))


def restore(plan,allow_legacy=False):
    root,originals,archive,review=roots(plan['root'],plan['originals'])
    rows=[]
    for candidate in plan['entries']:
        try:
            entry=valid_entry(candidate,root,archive)
            issue=restore_issue(entry)
            if issue: raise ValueError(issue)
            if not entry.get('pdf_sha256') and not allow_legacy: raise ValueError('Older PDF has no recorded checksum; explicit confirmation required.')
            source,pdf,saved=map(Path,(entry['source'],entry['pdf'],entry['archive']))
            os.makedirs(core.native_path(source.parent),exist_ok=True)
            if not os.path.exists(core.native_path(source)):
                copy_verified(saved,source,entry['source_sha256'])
                if core.digest(source)!=entry['source_sha256']: raise ValueError('Restored Word checksum failed; original archive retained.')
                try: shutil.copystat(core.native_path(saved),core.native_path(source))
                except OSError: pass
            if core.digest(source)!=entry['source_sha256'] or core.digest(saved)!=entry['source_sha256']:
                raise ValueError('Word source or archive changed during restoration; all remaining files retained.')
            # Recheck PDF immediately before deleting it; source restoration is verified first.
            if os.path.exists(core.native_path(pdf)):
                if entry.get('pdf_sha256') and core.digest(pdf)!=entry['pdf_sha256']: raise ValueError('PDF changed during restoration.')
                remove_matching(pdf,entry.get('pdf_sha256') or core.digest(pdf))
            remove_matching(saved,entry['source_sha256'])
            details='Original verified; matching converted PDF deleted if present.'
            record=saved_record_path(source,archive)
            try:
                if os.path.exists(core.native_path(record)):os.unlink(core.native_path(record))
            except OSError as error:details+=' Restore record retained: '+str(error)
            try:clear_incomplete_conversion(source)
            except OSError as error:details+=' Incomplete-conversion marker retained: '+str(error)
            rows.append([entry['source'],'RESTORED',details])
        except Exception as error: rows.append([candidate.get('source',''),'NOT RESTORED',str(error)])
    return rows


def plan_delete_logs(root,originals):
    root,originals,archive,review=roots(root,originals)
    locations=[core.HERE/'Logs',archive,review,originals/'REVIEW'/root.name]
    files={str(p.resolve()):p for base in locations for p in log_files(base)}
    return dict(action='delete-logs',root=str(root),originals=str(originals),files=[dict(path=str(p),sha256=core.digest(p)) for p in files.values()],count=len(files))


def delete_logs(plan):
    root,originals,archive,review=roots(plan['root'],plan['originals'])
    allowed={str(p.resolve()) for base in (core.HERE/'Logs',archive,review,originals/'REVIEW'/root.name) for p in log_files(base)}
    rows=[]
    for item in plan['files']:
        path=Path(item['path'])
        try:
            if str(path.resolve()) not in allowed: raise ValueError('Not an eligible log in the selected scope.')
            if core.digest(path)!=item['sha256']: raise ValueError('Log changed since confirmation; retained.')
            # Keep the historical proof required for safe restore, separately from logs.
            if 'conversion-report' in path.name.lower() and core.inside(path.resolve(),archive.resolve()):
                backup=archive/RECORDS/'legacy-reports'/(item['sha256']+path.suffix)
                os.makedirs(core.native_path(backup.parent),exist_ok=True)
                if not backup.exists(): shutil.copyfile(core.native_path(path),core.native_path(backup))
                if core.digest(backup)!=item['sha256']: raise ValueError('Restore-record backup failed; log retained.')
            os.unlink(core.native_path(path)); rows.append([str(path),'DELETED',''])
        except Exception as error: rows.append([str(path),'RETAINED',str(error)])
    return rows


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('action',choices=['plan-restore','restore','plan-delete-logs','delete-logs'])
    parser.add_argument('--root'); parser.add_argument('--originals'); parser.add_argument('--plan',required=True)
    parser.add_argument('--allow-legacy',action='store_true')
    args=parser.parse_args(); path=Path(args.plan)
    if args.action.startswith('plan-'):
        plan=plan_restore(args.root,args.originals) if args.action=='plan-restore' else plan_delete_logs(args.root,args.originals)
        path.write_text(json.dumps(plan),encoding='utf-8'); return
    plan=json.loads(path.read_text())
    rows=restore(plan,args.allow_legacy) if args.action=='restore' else delete_logs(plan)
    report=path.with_suffix('.csvlog')
    with open(report,'w',newline='',encoding='utf-8-sig') as stream:
        writer=csv.writer(stream); writer.writerow(['File','Outcome','Details']); writer.writerows(rows)
        stream.flush();os.fsync(stream.fileno())
    for row in rows: print(' | '.join(row),flush=True)
    try:
        workbook=path.with_suffix('.xlsx')
        build_excel(report,workbook,'Maintenance')
        print('Maintenance report: '+str(workbook),flush=True)
    except Exception as error:
        print('Excel maintenance report failed: '+str(error)+'; recovery journal: '+str(report),file=sys.stderr,flush=True)
    print('Completed: '+str(sum(r[1] in ('RESTORED','DELETED') for r in rows))+'; retained/blocked: '+str(sum(r[1] not in ('RESTORED','DELETED') for r in rows)),flush=True)


if __name__=='__main__': main()
