"""Queue-scoped log naming, with incomplete runs marked from creation."""
from datetime import datetime, timezone, timedelta
from pathlib import Path
import re
import os
import hashlib
import time
from fileops import native_path


def root_label(name):
    if len(name.encode('utf-16-le'))//2<=100:return name
    prefix=name.encode('utf-16-le')[:160].decode('utf-16-le','ignore')
    return prefix+' ['+hashlib.sha256(name.encode('utf-8')).hexdigest()[:10]+']'


def queue_name(root):
    stamp=datetime.now(timezone(timedelta(hours=-7))).strftime('%Y-%m-%d %H-%M-%S.%f')
    return f'{root_label(Path(root).name)} [{stamp}]'


def status_suffix(completed, errors):
    return ('' if completed else ' FAILED')+(' ERRORS' if errors else '')


class RunLogs:
    def __init__(self, name, recovery_root=None):
        if not name or re.search(r'[<>:"/\\|?*\x00-\x1f]',name):
            raise ValueError('Invalid queue log name.')
        self.name=name
        self.entries=[]
        self.recovery_root=Path(recovery_root) if recovery_root else None
        self.reports={}
        self.workbooks={}
        self.finalization_errors=[]
        self.rename_history={}

    def path(self, folder, kind, extension):
        if extension=='csv':
            recovery=(self.recovery_root or Path(folder)/'.report-recovery')/self.name
            os.makedirs(native_path(recovery),exist_ok=True)
            path=recovery/f'{self.name} - {kind} FAILED.csvlog'
            self.reports[str(path)]=(Path(folder),kind)
            extension='csvlog'
        else:
            path=Path(folder)/f'{self.name} - {kind} FAILED.{extension}'
        self.entries.append((path,kind,extension))
        return path

    def build_reports(self):
        from reports import build_excel
        failures=[]
        for journal,(folder,kind) in self.reports.items():
            path=Path(journal)
            if not os.path.exists(native_path(path)):continue
            workbook=folder/f'{self.name} - {kind} FAILED.xlsx'
            try:
                build_excel(path,workbook,kind)
                self.entries.append((workbook,kind,'xlsx'))
                self.workbooks[journal]=str(workbook)
            except Exception as error:
                failures.append(dict(journal=journal,output=str(workbook),error=type(error).__name__+': '+str(error)))
        return failures

    def finalize(self, completed, errors):
        renamed={};entries=[]
        for old,kind,extension in self.entries:
            actual=old
            new=old.with_name(f'{self.name} - {kind}{status_suffix(completed,errors)}.{extension}')
            try:
                if os.path.exists(native_path(old)) and old!=new:
                    for attempt in range(5):
                        try:
                            if os.path.exists(native_path(new)):raise FileExistsError('Final log name already exists.')
                            os.rename(native_path(old),native_path(new));actual=new;break
                        except OSError:
                            if attempt==4:raise
                            time.sleep(.05*(attempt+1))
                if os.path.exists(native_path(actual)):renamed[str(old)]=str(actual)
            except OSError as error:
                self.finalization_errors.append(dict(path=str(old),output=str(new),error=type(error).__name__+': '+str(error)))
                if os.path.exists(native_path(old)):renamed[str(old)]=str(old)
            entries.append((actual,kind,extension))
        self.entries=entries
        self.rename_history={key:renamed.get(value,value) for key,value in self.rename_history.items()}
        self.rename_history.update(renamed)
        for journal,workbook in self.workbooks.items():
            if workbook in self.rename_history:self.rename_history[journal]=self.rename_history[workbook]
        return dict(self.rename_history)
