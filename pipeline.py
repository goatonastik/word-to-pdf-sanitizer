"""Persistent Word export and reusable, isolated PDF verification."""
import csv
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import uuid
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ProcessPoolExecutor,ThreadPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from fieldcache import freeze_saved_fields
import convert as core
from fileops import copy_verified,copy_hashed,remove_readonly,remove_matching,owned_copy_proof,remove_owned_partial
from runlogs import RunLogs,queue_name
from debugtrace import DebugLog,configure_library_logging

DEFAULTS = dict(hideText=True, excludeProperties=True, preserveVisibleRevisions=True,debugLogging=True,
                excludeTags=True, stripMetadata=True, removeLinks=True, queueLimit=0, verifierWorkers=0, prefetchDocuments=2, removeAllComments=True)


def verifier_worker_count(config):
    requested=int(config.get('verifierWorkers',0))
    if requested<0 or requested>16:raise ValueError('Verification workers must be between 0 and 16.')
    if requested:return requested
    # Automatic sizing for mixed Word/PDF workloads. Leave several logical CPUs
    # for WINWORD, the coordinator and Windows. Six workers fits a 20-thread,
    # 16-GB workstation while avoiding the memory spikes of one worker per thread.
    logical=os.cpu_count() or 4
    return max(1,min(6,max(2,(logical-2)//3)))


def review_reasons(result):
    """Only failures and deliberately preserved visible review content need review."""
    audit=result.get('audit',{})
    reasons=[]
    if audit.get('visibleRevisionsPreserved'):
        reasons.append('Visible tracked changes preserved using saved display settings; review redlines and reviewer labels.')
    if audit.get('visibleCommentsPreserved'):
        reasons.append('Visible comments preserved because comment removal is disabled; review reviewer identities.')
    if result['status'] not in ('CONVERTED','SKIPPED'):
        reasons.append('Document not fully converted/archived: '+concise_error(result.get('details','')))
    return reasons


def review_copy(source, destination, review_root, relative, run_id, audit=None):
    """Keep an unchanged, verified review copy without overwriting earlier copies."""
    if audit is None: audit={}
    original = source if os.path.isfile(core.native_path(source)) else destination
    audit.update(source=str(original),review_root=str(review_root),relative=str(relative),stage='Hashing original')
    if not os.path.isfile(core.native_path(original)):
        raise FileNotFoundError('Original unavailable for REVIEW copy.')
    expected = core.digest(original)
    target = review_root / relative
    if os.path.exists(core.native_path(target)):
        if target.is_symlink(): raise ValueError('REVIEW file is a filesystem link.')
        if core.digest(target) == expected: return str(target), False
        target = review_root / '_versions' / run_id / relative
    audit.update(target=str(target),targetCharacters=len(str(target)),sourceCharacters=len(str(original)),stage='Creating REVIEW parent directories')
    os.makedirs(core.native_path(target.parent),exist_ok=True)
    if not core.inside(target.parent.resolve(),review_root.resolve()):
        raise ValueError('REVIEW subfolder points outside the review folder.')
    created=False
    try:
        audit['stage']='Copying original bytes'
        with open(core.native_path(original),'rb') as src, open(core.native_path(target),'xb') as dst:
            created=True
            shutil.copyfileobj(src,dst,1024*1024)
            dst.flush(); os.fsync(dst.fileno())
        audit['stage']='Checking REVIEW checksum'
        if core.digest(target)!=expected or core.digest(original)!=expected:
            raise ValueError('REVIEW copy checksum mismatch.')
        audit['stage']='Preserving timestamps'
        try: shutil.copystat(core.native_path(original),core.native_path(target))
        except OSError as error:
            # The verified document bytes remain useful even if timestamps cannot be copied.
            audit['timestampWarning']=str(error)
    except Exception:
        if created and os.path.exists(core.native_path(target)): os.unlink(core.native_path(target))
        raise
    return str(target),True


def source_properties(path):
    """Inspect property categories only; never retain their values or custom names."""
    if path.suffix.lower() not in ('.docx','.docm'):
        return dict(sourcePropertyTypesFound='Not inspected (legacy Word format)', customPropertiesFound='Not measured')
    names={'creator':'Author','lastModifiedBy':'Last saved by','title':'Title','subject':'Subject','keywords':'Keywords','description':'Comments','Company':'Company','Manager':'Manager','Template':'Template','HyperlinkBase':'Hyperlink base','created':'Created','modified':'Modified'}
    found=set(); custom=0
    with zipfile.ZipFile(path) as package:
        for name in ('docProps/core.xml','docProps/app.xml','docProps/custom.xml'):
            if name not in package.namelist(): continue
            root=ET.fromstring(package.read(name))
            if name.endswith('custom.xml'):
                custom=len(root)
            else:
                for element in root.iter():
                    label=names.get(element.tag.rsplit('}',1)[-1])
                    if label and element.text and element.text.strip(): found.add(label)
    return dict(sourcePropertyTypesFound='; '.join(sorted(found)),customPropertiesFound=custom)


def atomic_json(path, value):
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(value), encoding='utf-8')
    os.replace(temp, path)


def concise_error(details):
    lines=details.strip().splitlines()
    for line in lines:
        if line.startswith(('RuntimeError:','ValueError:','TimeoutError:','PermissionError:','FileNotFoundError:','OSError:','FileExistsError:')):
            return line
    return lines[-1] if lines else 'No details available'


def working_copy_path(workspace, source):
    candidate=workspace/uuid.uuid4().hex/source.name
    # Word itself can reject paths that Python can access. Keep its input short.
    if len(str(candidate))>230:
        return workspace/(uuid.uuid4().hex[:12]+source.suffix.lower())
    return candidate


class WordDocumentError(RuntimeError):
    def __init__(self,message,audit=None):
        super().__init__(message);self.audit=audit or {}


class WordSessionError(RuntimeError):pass
class QueueFatalError(RuntimeError):pass


class WordSession:
    def __init__(self, workspace, settings, timeout):
        self.debug=core.DEBUG_LOG or DebugLog()
        self.debug.stage('starting Word worker',dict(workspace=workspace,settings=settings),timeout_seconds=timeout)
        self.workspace = workspace
        self.settings=settings
        self.timeout = timeout
        for name in ('job.json','result.json','ready.json','stop','owner.json'):
            (workspace/name).unlink(missing_ok=True)
        self.errors = open(workspace / 'word-errors.log', 'ab')
        self.process = subprocess.Popen(['powershell.exe', '-NoProfile', '-STA',
            '-ExecutionPolicy', 'Bypass', '-File', str(core.HERE / 'Export-Session.ps1'),
            '-Workspace', str(workspace), '-SettingsFile', str(settings),'-ParentPid',str(os.getpid())],
            stdout=subprocess.DEVNULL, stderr=self.errors,
            creationflags=subprocess.CREATE_NO_WINDOW)
        self.debug.event('worker_started',worker='Word',process_id=self.process.pid)
        try:
            self.ready = self.wait('ready.json', lambda: None)
        except Exception:
            self.close(force=False)
            raise

    def wait(self, name, tick):
        deadline = time.monotonic() + self.timeout
        next_heartbeat=time.monotonic()+5
        path = self.workspace / name
        while not path.exists():
            tick_started=time.monotonic()
            tick()
            deadline+=time.monotonic()-tick_started
            if path.exists():break
            if time.monotonic()>=next_heartbeat:
                self.debug.stage('waiting for Word response',dict(response=path,owner_record=self.workspace/'owner.json'),worker_pid=self.process.pid,remaining_seconds=round(deadline-time.monotonic(),1))
                next_heartbeat=time.monotonic()+5
            if self.process.poll() is not None:
                raise WordSessionError('Word session exited unexpectedly; see word-errors.log')
            if time.monotonic() > deadline:
                if path.exists():break
                self.close(force=True)
                raise TimeoutError('Word session timed out. Original retained.')
            time.sleep(.01)
        result = json.loads(path.read_text(encoding='utf-8-sig'))
        self.debug.event('word_response',response=name,ok=result.get('ok'),timings=result.get('timings'),audit=result.get('audit'))
        path.unlink()
        if not result['ok']:
            if name=='result.json':raise WordDocumentError(result['error'],result.get('audit'))
            raise WordSessionError(result['error'])
        return result

    def export(self, source, output, tick, original_name=None):
        self.debug.stage('submitting Word export',dict(source=source,output=output))
        atomic_json(self.workspace / 'job.json', dict(source=str(source), output=str(output),originalName=original_name or source.name))
        return self.wait('result.json', tick)

    def close(self, force=False):
        if getattr(self,'closed',False):return
        self.closed=True
        self.debug.event('closing_word_worker',worker_pid=self.process.pid,force=force)
        forced=force
        try:
            if self.process.poll() is None:
                (self.workspace/'stop').touch()
                if not force:
                    try:self.process.wait(timeout=15)
                    except subprocess.TimeoutExpired:forced=True
                if self.process.poll() is None:self.process.kill();self.process.wait();forced=True
            elif (self.workspace/'options.json').exists():forced=True
        except (OSError,subprocess.SubprocessError) as error:self.debug.exception('closing Word child',error)
        finally:self.errors.close()
        owner=self.workspace/'owner.json'
        if owner.exists():
            script=self.workspace/'stop-word.ps1'
            script.write_text("param($record) $r=Get-Content -LiteralPath $record|ConvertFrom-Json; $p=Get-Process -Id $r.pid -ErrorAction SilentlyContinue; if(!$r.released -and $p -and $p.ProcessName -eq 'WINWORD' -and $p.StartTime.ToUniversalTime().ToString('o') -eq $r.started){Stop-Process -Id $p.Id -Force; Wait-Process -Id $p.Id -Timeout 15 -ErrorAction SilentlyContinue}")
            try:subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),str(owner)],capture_output=True,timeout=20,creationflags=subprocess.CREATE_NO_WINDOW)
            except (OSError,subprocess.TimeoutExpired) as error:self.debug.exception('owned Word cleanup',error)
        if forced and (self.workspace/'options.json').exists():
            args=['powershell.exe','-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',str(core.HERE/'Export-Session.ps1'),'-Workspace',str(self.workspace),'-SettingsFile',str(self.settings),'-RestoreOnly']
            try:
                result=subprocess.run(args,capture_output=True,timeout=30,creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode:self.debug.event('Word option recovery failed',exit_code=result.returncode)
            except (OSError,subprocess.TimeoutExpired) as error:
                self.debug.exception('Word option recovery',error)
                if owner.exists():
                    try:subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),str(owner)],capture_output=True,timeout=20,creationflags=subprocess.CREATE_NO_WINDOW)
                    except (OSError,subprocess.TimeoutExpired):pass
            if (self.workspace/'options.json').exists():
                core.PROGRESS['wordRecoveryErrors']=core.PROGRESS.get('wordRecoveryErrors',0)+1
                backup=core.HERE/'Logs'/'Recovery'/('word-options-recovery-'+uuid.uuid4().hex+'.json')
                try:
                    backup.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(self.workspace/'options.json',backup)
                    print('Word settings could not be fully restored; recovery snapshot: '+str(backup),file=sys.stderr,flush=True)
                except OSError as error:self.debug.exception('saving Word option recovery snapshot',error)



def verify_and_archive(source, target, destination, raw, clean, original_hash, config, word_audit=None):
    debug=DebugLog(config.get('_verifierDebug'))
    configure_library_logging(debug)
    stage='starting PDF verification'
    debug.stage(stage,dict(source=source,target=target,archive=destination,raw_pdf=raw,clean_pdf=clean))
    core.STRIP_METADATA = config['stripMetadata']
    core.REMOVE_LINKS = config['removeLinks']
    timings = {}
    audit = dict(word_audit or {})
    audit.update(cleanup_verification='Not completed', original_appearance='Not checked')
    started = time.monotonic()
    try:
        stage='PDF cleanup, structure, text and render checks'; debug.stage(stage)
        pages = core.clean_and_check(raw, clean, timings, audit)
        audit['cleanup_verification']='Passed'
        tick=time.monotonic()
        try:
            stage='original-view appearance comparison'; debug.stage(stage,dict(reference=Path(str(raw)+'.baseline.pdf'),clean_pdf=clean))
            core.compare_reference(Path(str(raw)+'.baseline.pdf'),clean,audit)
            audit['original_appearance']='Passed (text and 144 dpi pixel comparison)'
        except Exception:
            audit['original_appearance']='Failed or unavailable; PDF withheld'
            raise
        timings['original_appearance_verification']=time.monotonic()-tick
        stage='checking source unchanged'; debug.stage(stage,dict(source=source))
        if core.digest(source) != original_hash:
            raise ValueError('Source changed during conversion; original retained.')
        if os.path.exists(core.native_path(target)) or os.path.exists(core.native_path(destination)):
            raise FileExistsError('Output or archive appeared during processing; source retained.')
        tick = time.monotonic()
        stage='publishing verified PDF'; debug.stage(stage,dict(clean_pdf=clean,published_pdf=target))
        # Publish only a complete verified PDF; no temporary files in source/archive.
        clean_hash=core.digest(clean)
        pdf_hash=copy_verified(clean,target,clean_hash,Path(str(raw)+'.publish-owner.json'),dict(audit=audit))
        timings['publish'] = time.monotonic()-tick
        tick = time.monotonic()
        try:
            stage='creating archive directories'; debug.stage(stage,dict(archive=destination))
            os.makedirs(core.native_path(destination.parent), exist_ok=True)
            if not core.inside(destination.parent.resolve(), Path(config['archive'])):
                raise ValueError('Archive folder resolves outside the selected archive.')
            stage='copying original into archive'; debug.stage(stage,dict(source=source,archive=destination))
            archive_hash=copy_verified(source,destination,original_hash,Path(str(raw)+'.archive-owner.json'))
            stage='checking archive checksum'; debug.stage(stage)
            if archive_hash != original_hash or core.digest(source) != original_hash:
                raise ValueError('Archive checksum mismatch; source retained.')
            stage='preserving archive timestamps'; debug.stage(stage)
            try:shutil.copystat(core.native_path(source), core.native_path(destination))
            except OSError as error:audit['archiveTimestampWarning']=str(error)
            stage='saving restore record'; debug.stage(stage)
            from maintenance import save_restore_record
            save_restore_record(source,target,destination,original_hash,pdf_hash,config,audit)
            stage='removing verified source copy'; debug.stage(stage,dict(source=source,archive=destination))
            remove_matching(source,original_hash)
        except Exception as error:
            debug.exception(stage,error)
            return dict(status='PDF CREATED; SOURCE RETAINED', details=traceback.format_exc(), timings=timings, audit=audit)
        timings['archive'] = time.monotonic()-tick
        return dict(status='CONVERTED', details=f'{pages} pages; SHA256 {original_hash}', timings=timings, verifier_pid=os.getpid(), audit=audit)
    except Exception as error:
        debug.exception(stage,error)
        return dict(status='FAILED; SOURCE RETAINED', details=traceback.format_exc(), timings=timings, audit=audit)
    finally:
        debug.event('verification_worker_finished',document=str(source),last_stage=stage,timings=timings,audit=audit)
        timings['verification_worker_total'] = time.monotonic()-started
        for path in (raw,clean,Path(str(raw)+'.baseline.pdf'),Path(str(raw)+'.publish-owner.json'),Path(str(raw)+'.archive-owner.json')):
            try:path.unlink(missing_ok=True)
            except OSError as error:debug.event('temporary_file_cleanup_delayed',path=str(path),error=str(error))


def batch(root, archive, timeout=180, preview=False):
    catalog=RunLogs(core.QUEUE_NAME or queue_name(root),core.HERE/'Logs'/'Recovery')
    core.RUN_LOGS=catalog
    core.PROGRESS.clear()
    core.DEBUG_LOG=DebugLog()
    completed=False
    try:
        settings=json.loads(core.SETTINGS_FILE.read_text(encoding='utf-8-sig')) if core.SETTINGS_FILE else {}
        enabled=core.DEBUG_OVERRIDE if core.DEBUG_OVERRIDE is not None else settings.get('debugLogging',True)
        if enabled:
            logs=core.HERE/'Logs'; logs.mkdir(exist_ok=True)
            debug_path=catalog.path(logs,'debug','jsonl')
            core.DEBUG_LOG=DebugLog(debug_path)
            core.progress(debugLog=str(debug_path))
            version_path=core.HERE/'VERSION.txt'
            core.DEBUG_LOG.event('environment',program_version=version_path.read_text().strip() if version_path.exists() else 'unknown',python=sys.version,executable=sys.executable,platform=sys.platform,temporary_directory=tempfile.gettempdir(),settings=settings)
            core.DEBUG_LOG.stage('queue started',dict(source_root=root,archive_root=archive))
            core.DEBUG_LOG.event('program_fingerprints',sha256={name:core.digest(core.HERE/name) for name in ('convert.py','pipeline.py','debugtrace.py','maintenance.py','PropertyInventory.ps1','WordSessionSupport.ps1','Export-Session.ps1','Interface.ps1','ProcessLauncher.ps1','LogTail.ps1','runlogs.py','reports.py','fileops.py','fieldcache.py')})
        configure_library_logging(core.DEBUG_LOG)
        _batch(root,archive,timeout,preview)
        completed=bool(core.PROGRESS.get('finished')) and core.PROGRESS.get('unprocessed',0)==0
    except BaseException as error:
        core.DEBUG_LOG.exception('queue aborted',error)
        core.progress(finished=False,phase='Queue did not complete',fatalError=traceback.format_exc(),unprocessed=max(0,core.PROGRESS.get('total',0)-core.PROGRESS.get('done',0)))
        raise
    finally:
        core.DEBUG_LOG.event('queue_finished',completed=completed,progress=core.PROGRESS)
        previous_phase=core.PROGRESS.get('phase','Queue finished')
        core.progress(phase='Building Excel reports')
        report_started=time.monotonic()
        report_errors=catalog.build_reports()
        report_seconds=time.monotonic()-report_started
        core.DEBUG_LOG.event('excel_reports_finished',seconds=report_seconds,errors=report_errors)
        for failure in report_errors:
            core.DEBUG_LOG.event('excel_report_error',**failure)
        core.progress(reportErrors=report_errors,phase=previous_phase)
        errors=bool(core.PROGRESS.get('failed',0) or core.PROGRESS.get('partial',0) or core.PROGRESS.get('reviewCopyErrors',0) or core.PROGRESS.get('fatalError') or core.PROGRESS.get('reportWriteErrors',0) or core.PROGRESS.get('wordRecoveryErrors',0) or report_errors)
        mapping=catalog.finalize(completed,errors)
        if catalog.finalization_errors:
            errors=True
            mapping=catalog.finalize(completed,True)
            report_errors.extend(catalog.finalization_errors)
            for failure in catalog.finalization_errors:print('Log finalization error; retained original name: '+failure['path']+'; '+failure['error'],file=sys.stderr,flush=True)
        for failure in report_errors:
            if 'journal' in failure:
                failure['journal']=mapping.get(failure['journal'],failure['journal'])
                print('Excel report could not be created: '+failure['error']+'; recovery journal retained: '+failure['journal'],file=sys.stderr,flush=True)
        for key in ('debugLog','wordDebugLog','verificationDebugLog'):
            value=core.PROGRESS.get(key)
            if value and value not in mapping and not os.path.exists(core.native_path(value)):
                core.PROGRESS.pop(key,None)
        updates={key:mapping[value] for key,value in core.PROGRESS.items() if isinstance(value,str) and value in mapping}
        core.progress(queueCompleted=completed,queueErrors=errors,queueName=catalog.name,reportErrors=report_errors,reportGenerationSeconds=round(report_seconds,3),**updates)
        for key in ('debugLog','wordDebugLog','verificationDebugLog','scanReport','report','privacyReport','diagnostics','reviewReport'):
            if core.PROGRESS.get(key): print(f'{key}: {core.PROGRESS[key]}',flush=True)
        print('QUEUE_RESULT: '+json.dumps(core.PROGRESS,default=str),flush=True)


def _batch(root, archive, timeout=180, preview=False):
    started = time.monotonic()
    debug=core.DEBUG_LOG or DebugLog()
    config = DEFAULTS.copy()
    if core.SETTINGS_FILE:
        config.update(json.loads(core.SETTINGS_FILE.read_text(encoding='utf-8-sig')))
    config['debugLogging']=bool(core.DEBUG_LOG and core.DEBUG_LOG.path)
    if config['debugLogging']:
        config['_wordDebug']=str(core.RUN_LOGS.path(core.HERE/'Logs','word-debug','jsonl'))
        config['_verifierDebug']=str(core.RUN_LOGS.path(core.HERE/'Logs','verification-debug','jsonl'))
        core.progress(wordDebugLog=config['_wordDebug'],verificationDebugLog=config['_verifierDebug'])
    limit = int(config.get('queueLimit', 0))
    if limit < 0:
        raise ValueError('Queue limit must be 0 or a positive integer.')
    worker_count=verifier_worker_count(config)
    root, archive = core.validate_roots(root, archive)
    config['archive'] = str(archive)
    review_root=Path(core.REVIEW_ROOT or archive.parent/'#REVIEW'/root.name)
    core.validate_review_root(root,archive,review_root)
    # Reject redirected REVIEW roots that escape the intended originals location.
    if review_root.resolve()!=review_root.absolute():
        raise ValueError('REVIEW folder must not be a filesystem link.')
    core.progress(phase='Scanning folders', total=0, done=0, exported=0, queued=0, converted=0, skipped=0, failed=0, partial=0, finished=False, reviewAdded=0, reviewFlagged=0, reviewCopyErrors=0, reviewRoot=str(review_root),verifierWorkers=worker_count,maxPendingVerification=0)
    logs = core.HERE / 'Logs'; logs.mkdir(exist_ok=True)
    scan_path=core.RUN_LOGS.path(logs,'scan-report','csv')
    scan={}
    debug.stage('scanning folders',dict(source=root,archive=archive,review=review_root))
    files = list(core.documents(root,scan))
    only=getattr(core,'ONLY_DOCUMENT',None)
    if only:
        files=[path for path in files if path.resolve()==Path(only).resolve()]
        if not files:raise FileNotFoundError('The selected document is no longer available for one-document conversion.')
    debug.event('scan_completed',scan=scan)
    with open(scan_path,'x',newline='',encoding='utf-8-sig') as scan_stream:
        writer=csv.writer(scan_stream)
        writer.writerow(['Folder','Eligible Word files','Other files','Other extensions/counts','Scan error'])
        for row in scan['folders']:
            writer.writerow([row['folder'],row['wordDocuments'],row['otherFiles'],json.dumps(row['otherExtensions']),''])
        for error in scan['errors']: writer.writerow([error['path'],'','','',error['error']])
        scan_stream.flush();os.fsync(scan_stream.fileno())
    core.progress(total=len(files),scanReport=str(scan_path),scanFolders=scan['foldersScanned'],scanOtherFiles=scan['otherFiles'])
    print(f'Starting directory: {root}\nArchive directory: {archive}\nREVIEW directory: {review_root}\nScan: {scan["foldersScanned"]} folders, {len(files)} eligible Word files (.doc, .docx, .docm), {scan["otherFiles"]} other files, {scan["lockFiles"]} lock files, {scan["excludedLinks"]} filesystem links excluded.',flush=True)
    if scan['errors']:
        raise RuntimeError('Folder scan was incomplete. No documents converted. See scan report: '+str(scan_path))
    if preview:
        for source in files:
            print(f'{source}\n  PDF: {source.with_suffix(".pdf")}\n  Archive: {archive / source.relative_to(root)}')
        print(f'{len(files)} documents found. No files changed.')
        core.progress(finished=True, phase='Preview complete')
        return
    os.makedirs(core.native_path(archive),exist_ok=True)
    run_id = core.RUN_LOGS.name
    report = core.RUN_LOGS.path(archive,'conversion-report','csv')
    diagnostic_path = core.RUN_LOGS.path(logs,'diagnostics','jsonl')
    privacy_path = core.RUN_LOGS.path(logs,'privacy-report','csv')
    review_report=core.RUN_LOGS.path(review_root,'review-report','csv')
    word_errors=core.RUN_LOGS.path(logs,'word-errors','log')
    counts = dict(converted=0, skipped=0, failed=0, partial=0)
    pending = {}; reserved = set(); done = 0; exported = 0; session = None
    review_added=0; review_flagged=0; review_errors=0
    def log_failure(label,error):
        core.PROGRESS['reportWriteErrors']=core.PROGRESS.get('reportWriteErrors',0)+1
        debug.exception(label,error)
        print('REPORT WRITE ERROR ('+label+'): '+str(error),file=sys.stderr,flush=True)
    def diagnostic(value):
        debug.event('diagnostic_record',record=value)
        try:
            with open(core.native_path(diagnostic_path),'a',encoding='utf-8') as stream:
                stream.write(json.dumps(value,default=str)+'\n');stream.flush();os.fsync(stream.fileno())
        except OSError as error:log_failure('diagnostics',error)
    def journal_row(writer,stream,values,label):
        try:writer.writerow(values);stream.flush();os.fsync(stream.fileno())
        except (OSError,csv.Error) as error:
            log_failure(label,error)
            diagnostic(dict(event='report_row_fallback',report=label,row=values))
    diagnostic(dict(event='run', config=config, python=sys.version, platform=sys.platform, root=root, archive=archive))
    core.progress(report=str(report), diagnostics=str(diagnostic_path), privacyReport=str(privacy_path))
    with tempfile.TemporaryDirectory(prefix='word-pdf-export-',ignore_cleanup_errors=True) as temp, open(report,'x',newline='',encoding='utf-8-sig') as stream, open(privacy_path,'x',newline='',encoding='utf-8-sig') as privacy_stream:
        workspace = Path(temp).resolve()
        if core.inside(workspace, root) or core.inside(workspace, archive):
            raise ValueError('Windows temporary directory is inside a selected folder. Choose different folders.')
        settings = workspace / 'settings.json'; atomic_json(settings, config)
        log = csv.writer(stream)
        log.writerow(['Source','PDF','Archive','Status','Details','Stage timings (seconds)','Review reasons','Review copy','Review copy result'])
        privacy=csv.writer(privacy_stream)
        privacy.writerow(['Batch ID','Document','PDF','Outcome','Data types removed or excluded','Word properties export','Source custom properties found','Comments removed','Revisions accepted','Hidden ranges removed','PDF property fields removed','XMP streams removed','Hyperlinks removed','Private app dictionaries removed','Object dates removed','Accessibility tags export','Cleanup verification','Original final-view appearance','Notes','Source standard property types found','PDF property field names removed','Revision handling','Review reasons','Review copy','Review copy result'])
        def record_impl(source, result):
            nonlocal done,review_added,review_flagged,review_errors
            status = result['status']; done += 1
            key = 'converted' if status=='CONVERTED' else 'skipped' if status=='SKIPPED' else 'partial' if status.startswith('PDF CREATED') else 'failed'
            counts[key] += 1
            a=result.get('audit',{})
            reasons=review_reasons(result)
            try:
                from maintenance import clear_incomplete_conversion,remember_incomplete_conversion
                if status=='CONVERTED':clear_incomplete_conversion(source)
                elif status!='SKIPPED' and not result.get('prior_incomplete'):
                    remember_incomplete_conversion(source,source.with_suffix('.pdf'),archive/source.relative_to(root),result['details'])
            except (OSError,ValueError) as error:log_failure('incomplete-conversion safety record',error)
            inventory_warning=''
            if a.get('sourcePropertyTypesFound') in ('Not inspected (legacy Word format)','Inspection unavailable'):
                inventory_warning='Source property inventory unavailable; PDF sanitation and verification remain independent.'
                print('WARNING: '+str(source)+': '+inventory_warning,flush=True)
            copy_path=''; copy_state='Not required'
            if reasons:
                review_flagged+=1
                copy_audit={}
                try:
                    archived=archive/source.relative_to(root)
                    copy_from=archived if status=='CONVERTED' else source
                    copy_path,added=review_copy(copy_from,archived,review_root,source.relative_to(root),run_id,copy_audit)
                    review_added+=int(added)
                    copy_state='Added and checksum verified' if added else 'Identical review copy already exists'
                    if copy_audit.get('timestampWarning'): copy_state+='; timestamp warning (see diagnostics)'
                    diagnostic(dict(event='review_copy',document=str(source),audit=copy_audit))
                except Exception:
                    review_errors+=1; copy_state='FAILED: '+traceback.format_exc().strip().splitlines()[-1]
                    diagnostic(dict(event='review_copy_error',document=str(source),audit=copy_audit,traceback=traceback.format_exc()))
                result.update(review_reasons=reasons,review_copy=copy_path,review_copy_result=copy_state)
                try:
                    review_root.mkdir(parents=True,exist_ok=True)
                    new_report=not review_report.exists()
                    with open(review_report,'a',newline='',encoding='utf-8-sig') as review_stream:
                        review_writer=csv.writer(review_stream)
                        if new_report: review_writer.writerow(['Document','PDF','Outcome','Revision handling','Reasons for review','Review copy','Copy result'])
                        review_writer.writerow([source,source.with_suffix('.pdf'),status,a.get('revisionPolicy','Not measured'),'; '.join(reasons),copy_path,copy_state])
                        review_stream.flush();os.fsync(review_stream.fileno())
                    core.progress(reviewReport=str(review_report))
                except Exception:
                    review_errors+=1
                    diagnostic(dict(event='review_log_error',source=source,traceback=traceback.format_exc()))
            journal_row(log,stream,[source,source.with_suffix('.pdf'),archive/source.relative_to(root),status,result['details'],json.dumps(result.get('timings',{})),'; '.join(reasons),copy_path,copy_state],'conversion report')
            types=[]
            for key,label in [('commentsRemoved','Word comments'),('revisionsAccepted','Word revision history'),('hiddenRangesRemoved','hidden-font ranges'),('pdf_property_fields_removed','PDF properties'),('xmp_streams_removed','XMP metadata'),('hyperlinks_removed','hyperlink destinations'),('private_app_dictionaries_removed','private application data'),('object_dates_removed','object modification dates')]:
                if a.get(key,0): types.append(label)
            if a.get('propertiesExcluded'): types.append('Word properties excluded at export')
            if a.get('tagsExcluded'): types.append('accessibility tags/alt text excluded at export')
            notes=concise_error(result['details'])
            if status=='CONVERTED': notes='No removed values recorded. Image payload metadata is not inspected; visible property values remain.'
            if inventory_warning: notes+=' '+inventory_warning
            if a.get('dynamicFieldsLocked'):notes+=f" {a['dynamicFieldsLocked']} environment-dependent field(s) frozen to saved display values."
            if a.get('secondExportSkipped'):notes+=' Unchanged Word appearance reference reused; redundant second Word export skipped.'
            if a.get('archiveTimestampWarning'):notes+=' Archive bytes verified; timestamp/attribute copy warning (see diagnostics).'
            journal_row(privacy,privacy_stream,[run_id,source,source.with_suffix('.pdf'),status,'; '.join(types) or 'None recorded / not processed', 'Excluded' if a.get('propertiesExcluded') else 'Included' if a else 'Not processed',a.get('customPropertiesFound',''),*[a.get(k,'') for k in ('commentsRemoved','revisionsAccepted','hiddenRangesRemoved','pdf_property_fields_removed','xmp_streams_removed','hyperlinks_removed','private_app_dictionaries_removed','object_dates_removed')], 'Excluded' if a.get('tagsExcluded') else 'Included' if a else 'Not processed',a.get('cleanup_verification','Not performed'),a.get('original_appearance','Not performed'),notes,a.get('sourcePropertyTypesFound',''),a.get('pdf_property_names_removed',''),a.get('revisionPolicy','Not measured'),'; '.join(reasons),copy_path,copy_state],'privacy report')
            diagnostic(dict(event='document', source=source, **result))
            print(f'{status}: {source}\n{result["details"]}', flush=True)
            if reasons: print('REVIEW: '+copy_state+' '+copy_path,flush=True)
            core.progress(done=done, exported=exported, queued=len(pending), elapsed=round(time.monotonic()-started,1), reviewAdded=review_added,reviewFlagged=review_flagged,reviewCopyErrors=review_errors, **counts)
        def record(source,result):
            before=done
            try:record_impl(source,result)
            except Exception as error:
                # Reporting failures must not discard another future or kill Word.
                if done==before:
                    counts['failed']+=1
                log_failure('document result',error)
                diagnostic(dict(event='unrecorded_document_result',document=str(source),result=result))
                print('RESULT RECOVERY: '+str(source)+'; '+result.get('status','UNKNOWN'),flush=True)

        def unknown_result(job,error):
            source,target,destination=job['source'],job['target'],job['destination']
            states={name:os.path.exists(core.native_path(value)) for name,value in (('source',source),('pdf',target),('archive',destination))}
            if not states['source'] and states['pdf'] and states['archive']:
                # A durable restore record is written only after successful checks.
                try:
                    from maintenance import saved_record_path
                    with open(core.native_path(saved_record_path(source,destination.parent if not job.get('archive_root') else Path(job['archive_root']))),encoding='utf-8') as stream:proof=json.load(stream)
                    if all(Path(proof[key]).resolve()==Path(job[field]).resolve() for key,field in (('source','source'),('pdf','target'),('archive','destination'))) and proof['source_sha256']==job['original_hash'] and core.digest(destination)==proof['source_sha256'] and core.digest(target)==proof['pdf_sha256']:
                        recovered=dict(proof.get('audit') or job['audit'])
                        recovered.setdefault('cleanup_verification','Passed (recovered from verified restore record)')
                        recovered.setdefault('original_appearance','Passed (recovered from verified restore record)')
                        return dict(status='CONVERTED',details='Verified completed conversion recovered from restore record after worker exit.',audit=recovered)
                except (OSError,ValueError,KeyError):pass
            recovered_audit=dict(job['audit']);notes=[]
            clean=job['clean'];publish_proof=Path(str(job['raw'])+'.publish-owner.json');archive_proof=Path(str(job['raw'])+'.archive-owner.json')
            try:
                if states['pdf'] and os.path.isfile(core.native_path(clean)):
                    expected=core.digest(clean);proof=owned_copy_proof(target,publish_proof)
                    if proof and proof.get('expected_hash')==expected and core.digest(target)==expected:
                        recovered_audit.update(proof.get('context',{}).get('audit',{}))
                        notes.append('Published PDF matches its verified working copy; archiving did not finish.')
                    elif remove_owned_partial(target,clean,publish_proof,expected):
                        states['pdf']=False;notes.append("Removed this run's identified incomplete PDF copy.")
                if states['archive'] and states['source'] and core.digest(source)==job['original_hash']:
                    if remove_owned_partial(destination,source,archive_proof,job['original_hash']):
                        states['archive']=False;notes.append("Removed this run's identified incomplete archive copy.")
            except (OSError,ValueError) as recovery_error:notes.append('Copy recovery could not complete: '+str(recovery_error))
            status='PDF CREATED; VERIFICATION STATE UNKNOWN' if states['pdf'] else 'FAILED; SOURCE RETAINED' if states['source'] else 'FAILED; SOURCE STATE UNKNOWN'
            return dict(status=status,details='Verification worker did not return a result: '+str(error)+'; filesystem state: '+json.dumps(states)+'; '+' '.join(notes),audit=recovered_audit)

        def restart_lane(index,generation):
            lane=verifiers[index]
            if lane['generation']!=generation:return
            try:lane['executor'].shutdown(wait=True,cancel_futures=True)
            except Exception:pass
            lane['restarts']+=1
            if lane['restarts']>3:raise QueueFatalError('A verification worker repeatedly exited; queue stopped. Unprocessed originals remain in place.')
            lane['generation']+=1
            lane['executor']=ProcessPoolExecutor(max_workers=1)
            debug.event('restarting_verification_worker',lane=index+1,restart=lane['restarts'])

        def collect(wait=False):
            heartbeat=time.monotonic()+5
            while True:
                ready = [future for future in pending if future.done()]
                if ready or not wait or not pending:
                    break
                time.sleep(.05)
                if time.monotonic()>=heartbeat:
                    debug.event('verification_heartbeat',pending_count=len(pending),pending_documents=[str(pending[f][0]) for f in list(pending)[:10]])
                    heartbeat=time.monotonic()+5
            for future in ready:
                source,timings,job,lane_index,generation=pending.pop(future)
                verifiers[lane_index]['outstanding']-=1
                try:
                    result = future.result()
                    if verifiers[lane_index]['generation']==generation:verifiers[lane_index]['restarts']=0
                except Exception as error:
                    if isinstance(error,BrokenProcessPool):restart_lane(lane_index,generation)
                    result=unknown_result(job,error)
                result['timings'] = dict(timings, **result.get('timings',{}))
                record(source, result)
        verifiers=[dict(executor=ProcessPoolExecutor(max_workers=1),outstanding=0,restarts=0,generation=0) for _ in range(worker_count)]
        startup_failures=0
        def submit_job(job):
            args=(job['source'],job['target'],job['destination'],job['raw'],job['clean'],job['original_hash'],config,job['audit'])
            lane_index=min(range(len(verifiers)),key=lambda i:verifiers[i]['outstanding'])
            lane=verifiers[lane_index];generation=lane['generation']
            try:future=lane['executor'].submit(verify_and_archive,*args)
            except BrokenProcessPool:
                restart_lane(lane_index,generation);lane=verifiers[lane_index];generation=lane['generation']
                future=lane['executor'].submit(verify_and_archive,*args)
            lane['outstanding']+=1
            return future,lane_index,generation
        prefetch_count=max(0,min(4,int(config.get('prefetchDocuments',2))))
        stage_pool=ThreadPoolExecutor(max_workers=max(1,prefetch_count),thread_name_prefix='word-source-stage')
        def stage_source(source):
            began=time.monotonic();working=working_copy_path(workspace,source)
            working.parent.mkdir(parents=True,exist_ok=True)
            original_hash=copy_hashed(source,working)
            return working,original_hash,time.monotonic()-began
        stage_futures={};prefetch_reserved=set()
        if prefetch_count:
            for source in files:
                target=source.with_suffix('.pdf');destination=archive/source.relative_to(root);key=str(target).lower()
                if not os.path.exists(core.native_path(target)) and not os.path.exists(core.native_path(destination)) and key not in prefetch_reserved:
                    prefetch_reserved.add(key);stage_futures[source]=stage_pool.submit(stage_source,source)
        try:
            for index, source in enumerate(files):
                collect()
                while limit and len(pending)>=limit:
                    collect(wait=True)
                target=source.with_suffix('.pdf'); destination=archive/source.relative_to(root)
                from maintenance import incomplete_conversion_reason
                prior=incomplete_conversion_reason(source,target,destination)
                if prior:
                    record(source,dict(status='PDF CREATED; PREVIOUS CONVERSION INCOMPLETE' if os.path.isfile(core.native_path(target)) else 'FAILED; PREVIOUS CONVERSION INCOMPLETE',details=prior,prior_incomplete=True));continue
                if os.path.exists(core.native_path(target)):
                    record(source,dict(status='SKIPPED',skip_kind='pdf_exists',details=f'Existing PDF: {target}. Not sanitized or verified by this run; Word original retained and PDF not overwritten.')); continue
                if os.path.exists(core.native_path(destination)):
                    record(source,dict(status='SKIPPED',skip_kind='archive_exists',details=f'Existing archived original: {destination}. Not converted or verified by this run; neither original overwritten.')); continue
                if str(target).lower() in reserved:
                    record(source,dict(status='SKIPPED',skip_kind='name_collision',details=f'Another Word file in this queue maps to the same PDF: {target}.')); continue
                reserved.add(str(target).lower())
                timings={}; tick=time.monotonic()
                current_stage='hashing source'
                try:
                    debug.stage(current_stage,dict(source=source,target=target,archive=destination))
                    current_stage='copying source into workspace'; debug.stage(current_stage,dict(source=source))
                    if source in stage_futures:working_source,original_hash,stage_seconds=stage_futures.pop(source).result()
                    else:working_source,original_hash,stage_seconds=stage_source(source)
                    timings['source_hash']=0.0
                    current_stage='freezing saved dynamic fields'
                    field_audit=freeze_saved_fields(working_source)
                    timings['source_copy']=stage_seconds
                    if session is None:
                        current_stage='starting Word session'
                        tick=time.monotonic(); session=WordSession(workspace,settings,timeout)
                        startup_failures=0
                        diagnostic(dict(event='word_session', ready=session.ready, startup_seconds=time.monotonic()-tick))
                    raw=workspace/(uuid.uuid4().hex+'.pdf'); clean=workspace/(uuid.uuid4().hex+'.pdf')
                    core.progress(phase='Exporting in Word; verification runs in background', current=str(source))
                    diagnostic(dict(event='export_paths',source=str(source),sourceCharacters=len(str(source)),workingCharacters=len(str(working_source)),workingFilenameShortened=working_source.name!=source.name))
                    current_stage='Word export'
                    result=session.export(working_source,raw,collect,source.name)
                    result.setdefault('audit',{}).update(field_audit)
                    try:
                        current_stage='source property inventory'; debug.stage(current_stage,dict(working_copy=working_source))
                        audit=result.setdefault('audit',{})
                        if working_source.suffix.lower() in ('.docx','.docm') or not audit.get('propertyInventoryAvailable'):
                            audit.update(source_properties(working_source))
                    except (zipfile.BadZipFile,ET.ParseError,OSError) as error:
                        debug.exception(current_stage,error)
                        result.setdefault('audit',{}).update(sourcePropertyTypesFound='Inspection unavailable',customPropertiesFound='Not measured',propertyInspectionError=type(error).__name__+': '+str(error))
                    timings.update(result['timings']); exported += 1
                    current_stage='submitting verification'
                    job=dict(source=source,target=target,destination=destination,raw=raw,clean=clean,original_hash=original_hash,audit=result.get('audit',{}),archive_root=str(archive))
                    future,lane_index,generation=submit_job(job)
                    debug.event('verification_submitted',document=str(source),pending_count=len(pending)+1)
                    pending[future]=(source,timings,job,lane_index,generation)
                    core.progress(exported=exported,queued=len(pending),maxPendingVerification=max(core.PROGRESS.get('maxPendingVerification',0),len(pending)))
                except Exception as error:
                    debug.exception(current_stage,error)
                    record(source,dict(status='FAILED; SOURCE RETAINED',details=traceback.format_exc(),timings=timings,audit=getattr(error,'audit',{})))
                    if isinstance(error,QueueFatalError):raise
                    if current_stage=='starting Word session':
                        startup_failures+=1
                        if startup_failures>=2:raise QueueFatalError('Word could not start twice; queue stopped. Remaining originals were not processed.') from error
                    if current_stage in ('starting Word session','Word export') and not isinstance(error,WordDocumentError):
                        if session:session.close(force=isinstance(error,TimeoutError))
                        session=None
            core.progress(phase='Finishing verification and archiving', current='')
            while pending: collect(wait=True)
        finally:
            # Stop scheduling, then retain every pending outcome before leaving TEMP.
            stage_pool.shutdown(wait=True,cancel_futures=True)
            for lane in verifiers:lane['executor'].shutdown(wait=True,cancel_futures=True)
            collect()
            if session:session.close()
            errors=workspace/'word-errors.log'
            if errors.exists():
                try:shutil.copyfile(errors,word_errors)
                except OSError as error:log_failure('Word stderr copy',error)
    if workspace.exists():
        debug.event('temporary_workspace_retained',path=str(workspace))
        core.PROGRESS['temporaryWorkspace']=str(workspace)
    summary=dict(**counts,total=len(files),done=done,exported=exported,queued=0,unprocessed=len(files)-done,elapsed=round(time.monotonic()-started,1),reviewAdded=review_added,reviewFlagged=review_flagged,reviewCopyErrors=review_errors)
    diagnostic(dict(event='summary',**summary))
    print(f'Summary: {json.dumps(summary)}',flush=True)
    core.progress(finished=True,phase='Complete - review counts and report',**summary)


