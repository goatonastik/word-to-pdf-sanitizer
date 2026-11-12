"""Exclusive, checksum-verified copies that clean up their own failed outputs."""
import hashlib
import os
import shutil
import stat
import json
from pathlib import Path


def native_path(path):
    value=os.path.abspath(os.fspath(path))
    if os.name=='nt' and not value.startswith('\\\\?\\'):
        return '\\\\?\\UNC\\'+value[2:] if value.startswith('\\\\') else '\\\\?\\'+value
    return value


def digest(path):
    result=hashlib.sha256()
    with open(native_path(path),'rb') as stream:
        for block in iter(lambda:stream.read(1024*1024),b''):result.update(block)
    return result.hexdigest()


def copy_hashed(source,target):
    """Copy a temporary working file while hashing the one source read."""
    identity=None
    try:
        result=hashlib.sha256()
        with open(native_path(source),'rb') as src,open(native_path(target),'xb') as dst:
            created=os.fstat(dst.fileno());identity=(created.st_dev,created.st_ino)
            for block in iter(lambda:src.read(1024*1024),b''):
                result.update(block);dst.write(block)
            dst.flush();os.fsync(dst.fileno())
        shutil.copystat(native_path(source),native_path(target))
        expected=result.hexdigest()
        if digest(target)!=expected:raise ValueError('Working-copy checksum mismatch; source retained.')
        return expected
    except BaseException:
        if identity is not None:
            try:
                current=os.stat(native_path(target),follow_symlinks=False)
                if (current.st_dev,current.st_ino)==identity:os.unlink(native_path(target))
            except OSError:pass
        raise


def copy_verified(source,target,expected_hash=None,ownership_record=None,context=None):
    """Never overwrite; remove a failed copy only if it is still the file we created.

    The caller owns directory validation and decides when the source may be removed.
    This handles caught failures; forced termination can still interrupt a final copy.
    """
    identity=None
    try:
        with open(native_path(source),'rb') as src,open(native_path(target),'xb') as dst:
            created=os.fstat(dst.fileno());identity=(created.st_dev,created.st_ino)
            if ownership_record:
                with open(native_path(ownership_record),'x',encoding='utf-8') as proof:
                    json.dump(dict(target=str(Path(target).resolve()),identity=identity,expected_hash=expected_hash,context=context or {}),proof)
                    proof.flush();os.fsync(proof.fileno())
            shutil.copyfileobj(src,dst,1024*1024)
            dst.flush();os.fsync(dst.fileno())
        copy_zone_identifier(source,target)
        actual=digest(target)
        if actual!=(expected_hash or digest(source)):
            raise ValueError('Copied file checksum mismatch; source retained.')
        return actual
    except BaseException as error:
        if identity is not None:
            try:
                current=os.stat(native_path(target),follow_symlinks=False)
                if identity[1] and (current.st_dev,current.st_ino)==identity:
                    os.unlink(native_path(target))
                else:
                    error.add_note('Failed-copy path no longer identifies the file created by this operation; path retained: '+str(target))
            except FileNotFoundError:pass
            except OSError as cleanup_error:
                error.add_note('Failed-copy cleanup also failed: '+str(cleanup_error))
        raise


def copy_zone_identifier(source,target):
    """Keep Windows Mark-of-the-Web on archived/restored originals, if present."""
    if os.name!='nt':return
    try:
        with open(native_path(str(source)+':Zone.Identifier'),'rb') as stream:data=stream.read()
    except FileNotFoundError:return
    except OSError as error:
        if getattr(error,'winerror',None) in (1,123) or error.errno==22:return
        raise
    with open(native_path(str(target)+':Zone.Identifier'),'xb') as stream:
        stream.write(data);stream.flush();os.fsync(stream.fileno())
    with open(native_path(str(target)+':Zone.Identifier'),'rb') as stream:
        if stream.read()!=data:raise ValueError('Mark-of-the-Web stream verification failed; original retained.')


def remove_readonly(path):
    """Remove an explicitly selected file; restore read-only status if deletion fails."""
    original=os.stat(native_path(path),follow_symlinks=False)
    changed=os.name=='nt' and not (original.st_mode & stat.S_IWRITE)
    if changed:os.chmod(native_path(path),original.st_mode | stat.S_IWRITE)
    try:os.unlink(native_path(path))
    except BaseException as error:
        if changed:
            try:
                current=os.stat(native_path(path),follow_symlinks=False)
                if (current.st_dev,current.st_ino)==(original.st_dev,original.st_ino):os.chmod(native_path(path),original.st_mode)
            except OSError as restore_error:error.add_note('Could not restore read-only attribute: '+str(restore_error))
        raise


def remove_matching(path,expected_hash):
    """Recheck a selected file immediately before removal; never discard changed bytes."""
    if digest(path)!=expected_hash:raise ValueError('File changed before removal; retained: '+str(path))
    remove_readonly(path)


def owned_copy_proof(target,record):
    try:
        with open(native_path(record),encoding='utf-8') as stream:proof=json.load(stream)
        info=os.stat(native_path(target),follow_symlinks=False)
        if not stat.S_ISREG(info.st_mode) or Path(proof['target']).resolve()!=Path(target).resolve() or tuple(proof['identity'])!=(info.st_dev,info.st_ino):return None
        return proof
    except (OSError,ValueError,KeyError,TypeError):return None


def remove_owned_partial(target,reference,record,expected_hash):
    """After worker death, delete only our identified copy with an exact partial prefix."""
    proof=owned_copy_proof(target,record)
    if not proof or proof.get('expected_hash')!=expected_hash or digest(reference)!=expected_hash:return False
    before=os.stat(native_path(target),follow_symlinks=False)
    if before.st_size>=os.stat(native_path(reference)).st_size:return False
    with open(native_path(target),'rb') as partial,open(native_path(reference),'rb') as original:
        for block in iter(lambda:partial.read(1024*1024),b''):
            if block!=original.read(len(block)):return False
    current=os.stat(native_path(target),follow_symlinks=False)
    if (current.st_dev,current.st_ino,current.st_size,current.st_mtime_ns)!=(before.st_dev,before.st_ino,before.st_size,before.st_mtime_ns):return False
    remove_readonly(target)
    return True
