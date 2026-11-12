"""Best-effort operational tracing. Never reads document text or property values."""
from datetime import datetime, timezone
import json
import logging
import os
from pathlib import Path
import sys
import threading
import time
import traceback
import hashlib
import stat


def native(path):
    value=os.path.abspath(os.fspath(path))
    if os.name=='nt' and not value.startswith('\\\\?\\'):
        return '\\\\?\\UNC\\'+value[2:] if value.startswith('\\\\') else '\\\\?\\'+value
    return value


def path_state(path):
    value=os.fspath(path)
    state=dict(path=value,characters=len(value),filename_characters=len(Path(value).name))
    try:
        raw=native(value)
        try:
            info=os.stat(raw);exists=True
            state.update(exists=True,is_file=stat.S_ISREG(info.st_mode),is_directory=stat.S_ISDIR(info.st_mode),bytes=info.st_size)
        except FileNotFoundError:
            exists=False;state.update(exists=False,is_file=False,is_directory=False)
        try:state['parent_exists']=stat.S_ISDIR(os.stat(native(Path(value).parent)).st_mode)
        except OSError:state['parent_exists']=False
    except OSError as error:
        state.update(error=str(error),winerror=getattr(error,'winerror',None),errno=error.errno)
    return state


class DebugLog:
    def __init__(self,path=None):
        self.path=Path(path) if path else None
        self.failures=0
        self._lock=threading.Lock()
        self._mutex=None
        if self.path and os.name=='nt':
            import ctypes
            from ctypes import wintypes
            kernel=ctypes.WinDLL('kernel32',use_last_error=True)
            kernel.CreateMutexW.argtypes=[wintypes.LPVOID,wintypes.BOOL,wintypes.LPCWSTR]
            kernel.CreateMutexW.restype=wintypes.HANDLE
            kernel.WaitForSingleObject.argtypes=[wintypes.HANDLE,wintypes.DWORD]
            kernel.ReleaseMutex.argtypes=[wintypes.HANDLE]
            kernel.CloseHandle.argtypes=[wintypes.HANDLE]
            self._kernel=kernel
            name='Local\\WordPdfDebug-'+hashlib.sha256(str(self.path.resolve()).casefold().encode()).hexdigest()[:32]
            self._mutex=kernel.CreateMutexW(None,False,name)

    def __del__(self):
        if getattr(self,'_mutex',None):
            self._kernel.CloseHandle(self._mutex)
            self._mutex=None

    def _acquire_process_lock(self):
        if self._mutex:
            import ctypes
            result=self._kernel.WaitForSingleObject(self._mutex,30000)
            if result not in (0,0x80):raise OSError('Timed out waiting for the shared debug log.')

    def _release_process_lock(self):
        if self._mutex:
            import ctypes
            self._kernel.ReleaseMutex(self._mutex)

    def event(self,event,**fields):
        if not self.path: return
        record=dict(timestamp_utc=datetime.now(timezone.utc).isoformat(),monotonic_seconds=time.monotonic(),pid=os.getpid(),thread=threading.get_ident(),event=event,**fields)
        try:
            with self._lock:
                self._acquire_process_lock()
                try:
                    with open(native(self.path),'a',encoding='utf-8') as stream:
                        stream.write(json.dumps(record,default=str,ensure_ascii=True)+'\n')
                        stream.flush()
                finally:self._release_process_lock()
        except OSError as error:
            self.failures+=1
            if self.failures==1: print('Debug log could not be written: '+str(error),file=sys.stderr,flush=True)

    def stage(self,stage,paths=None,**fields):
        if self.path:
            self.event('stage',stage=stage,paths={key:path_state(value) for key,value in (paths or {}).items()},**fields)

    def exception(self,stage,error):
        self.event('exception',stage=stage,exception_type=type(error).__name__,error=str(error),errno=getattr(error,'errno',None),winerror=getattr(error,'winerror',None),traceback=traceback.format_exc())


class LibraryLogHandler(logging.Handler):
    def __init__(self,debug):super().__init__(logging.WARNING);self.debug=debug
    def emit(self,record):self.debug.event('library_warning',logger=record.name,message=record.getMessage())


def configure_library_logging(debug):
    logger=logging.getLogger('pypdf')
    for handler in list(logger.handlers):
        if isinstance(handler,LibraryLogHandler):logger.removeHandler(handler)
    logger.addHandler(LibraryLogHandler(debug));logger.propagate=False
