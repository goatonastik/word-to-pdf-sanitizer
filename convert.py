"""Export Word documents, validate clean PDFs, and archive unchanged originals."""
import argparse
import csv
import re
from urllib.parse import urlsplit,unquote
import hashlib
import logging
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import uuid
import traceback
import platform
import stat

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DictionaryObject, ArrayObject, IndirectObject
import pypdfium2 as pdfium

HERE = Path(__file__).resolve().parent
SETTINGS_FILE = None
STRIP_METADATA = True
REMOVE_LINKS = True
PROGRESS_FILE = None
PROGRESS = {}
REVIEW_ROOT = None
QUEUE_NAME = None
RUN_LOGS = None
DEBUG_LOG = None
DEBUG_OVERRIDE = None


def progress(**values):
    PROGRESS.update(values)
    if PROGRESS_FILE:
        pending = PROGRESS_FILE.with_name(PROGRESS_FILE.name+'.'+uuid.uuid4().hex+'.tmp')
        try:
            with open(native_path(pending),'x',encoding='utf-8') as stream:
                json.dump(PROGRESS,stream)
            for attempt in range(5):
                try:
                    os.replace(native_path(pending),native_path(PROGRESS_FILE))
                    return
                except OSError:
                    if attempt==4: raise
                    time.sleep(.025*(attempt+1))
        except OSError as error:
            if DEBUG_LOG: DEBUG_LOG.exception('progress snapshot write',error)
            # Progress is advisory: a reader/antivirus lock must not fail a document.
            PROGRESS['progressWriteWarnings']=PROGRESS.get('progressWriteWarnings',0)+1
            PROGRESS['lastProgressWriteError']=str(error)
            if PROGRESS['progressWriteWarnings']==1:
                print('Progress display update delayed; conversion continues: '+str(error),file=sys.stderr,flush=True)
        finally:
            try: os.unlink(native_path(pending))
            except OSError: pass
EXTENSIONS = {'.doc', '.docx', '.docm'}
LINK_WORKSPACE = None
FORBIDDEN = {'/Annots', '/AcroForm', '/EmbeddedFiles', '/AF', '/JavaScript',
             '/JS', '/OpenAction', '/AA', '/OCProperties'}
PRIVATE = {'/Metadata', '/PieceInfo', '/LastModified'}


class PrinterWhitespaceFilter(logging.Filter):
    def filter(self, record):
        # Microsoft Print to PDF pads object headers; this is harmless and
        # otherwise generates a warning for almost every object on every page.
        return not record.getMessage().startswith('Superfluous whitespace found in object header')


logging.getLogger('pypdf._reader').addFilter(PrinterWhitespaceFilter())


def digest(path):
    h = hashlib.sha256()
    with open(native_path(path), 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def native_path(path):
    """Use extended Windows paths for file I/O; keep ordinary paths in reports."""
    value=os.path.abspath(os.fspath(path))
    if os.name=='nt' and not value.startswith('\\\\?\\'):
        return '\\\\?\\UNC\\'+value[2:] if value.startswith('\\\\') else '\\\\?\\'+value
    return value


def display_path(path):
    """Keep extended prefixes internal; logs and relative paths use normal names."""
    value=os.fspath(path)
    if value.startswith('\\\\?\\UNC\\'): value='\\\\'+value[8:]
    elif value.startswith('\\\\?\\'): value=value[4:]
    return Path(value)


def inside(path, parent):
    return path == parent or parent in path.parents


def validate_roots(root, archive):
    root, archive = root.resolve(), archive.resolve()
    if not root.is_dir():
        raise ValueError('The source folder does not exist.')
    if inside(root, archive) or inside(archive, root):
        raise ValueError('Choose separate source and archive folders; neither may be inside the other.')
    if archive.exists() and not archive.is_dir():
        raise ValueError('The archive path is not a folder.')
    return root, archive


def validate_review_root(root,archive,review):
    paths=[Path(root).resolve(),Path(archive).resolve(),Path(review).resolve()]
    for i,first in enumerate(paths):
        for second in paths[i+1:]:
            if inside(first,second) or inside(second,first):
                raise ValueError('Starting, originals archive and #REVIEW must be separate folders; none may contain another.')


def documents(root, audit=None):
    if audit is None: audit={}
    audit.update(foldersScanned=0,wordDocuments=0,otherFiles=0,excludedLinks=0,lockFiles=0,errors=[],folders=[])
    def failed(error):
        audit['errors'].append(dict(path=str(display_path(error.filename)) if error.filename else '(unknown)',error=str(error)))
    # Do not follow directory junctions/symlinks or process Word lock files.
    for folder, dirs, files in os.walk(native_path(root), followlinks=False,onerror=failed):
        folder=display_path(folder)
        audit['foldersScanned']+=1
        allowed=[]
        for name in sorted(dirs):
            try:
                if is_link(folder/name): audit['excludedLinks']+=1
                else: allowed.append(name)
            except OSError as error: failed(error)
        dirs[:]=allowed
        row=dict(folder=str(folder),wordDocuments=0,otherFiles=0,otherExtensions={})
        for name in sorted(files):
            path = Path(folder) / name
            if name.startswith('~$'):
                audit['lockFiles']+=1; continue
            # Non-inputs (including long .eml paths) need no file inspection.
            if path.suffix.lower() not in EXTENSIONS:
                audit['otherFiles']+=1; row['otherFiles']+=1
                ext=path.suffix.lower() or '(no extension)'
                row['otherExtensions'][ext]=row['otherExtensions'].get(ext,0)+1
                continue
            try:
                if is_link(path):
                    audit['excludedLinks']+=1; continue
            except OSError as error:
                failed(error); continue
            audit['wordDocuments']+=1; row['wordDocuments']+=1
            yield path
        audit['folders'].append(row)


def is_link(path):
    info=os.lstat(native_path(path))
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info,'st_reparse_tag',0) & 0x20000000) or (not hasattr(info,'st_reparse_tag') and bool(getattr(info,'st_file_attributes',0) & 0x400))


def validate_link(annotation):
    action=annotation.get('/A')
    if action is None:return
    action=action.get_object()
    if '/Next' in action:raise ValueError('Unsupported chained PDF link action.')
    kind=str(action.get('/S',''))
    if kind=='/GoTo':return
    if kind=='/URI':
        uri=action.get('/URI','')
        if isinstance(uri,bytes):uri=uri.decode('utf-8','strict')
        uri=str(uri)
        if urlsplit(uri).scheme.lower() in ('http','https','mailto'):
            decoded=unquote(uri).replace('\\','/').lower()
            workspace=str(LINK_WORKSPACE or '').replace('\\','/').lower().rstrip('/')
            if workspace and workspace in decoded:raise ValueError('PDF link contains a conversion workspace path.')
            return
    raise ValueError('Unsupported or local-file PDF link action; enable hyperlink removal to omit it.')


def walk_pdf(value, clean=False, seen=None, audit=None):
    if seen is None:
        seen = set()
    if isinstance(value, IndirectObject):
        key = (id(value.pdf), value.idnum, value.generation)
        if key in seen:
            return
        seen.add(key)
        value = value.get_object()
    if isinstance(value, DictionaryObject):
        for key in list(value.keys()):
            if str(key) == '/Annots':
                annotations = value[key]
                if any(a.get_object().get('/Subtype') != '/Link' for a in annotations):
                    raise ValueError('Unexpected PDF annotation (not a hyperlink).')
                if REMOVE_LINKS:
                    if clean:
                        if audit is not None:
                            audit['hyperlinks_removed'] = audit.get('hyperlinks_removed',0)+len(annotations)
                        del value[key]
                    else:
                        raise ValueError('PDF retains hyperlinks.')
                    continue
                for annotation in annotations:validate_link(annotation.get_object())
                walk_pdf(value.raw_get(key), clean, seen, audit)
                continue
            if str(key) in FORBIDDEN:
                raise ValueError('Unexpected PDF feature: ' + str(key))
            if str(key) in PRIVATE and STRIP_METADATA:
                if clean:
                    if audit is not None:
                        label={'/Metadata':'xmp_streams_removed','/PieceInfo':'private_app_dictionaries_removed','/LastModified':'object_dates_removed'}[str(key)]
                        audit[label]=audit.get(label,0)+1
                    del value[key]
                else:
                    raise ValueError('PDF retains private data: ' + str(key))
            else:
                walk_pdf(value.raw_get(key), clean, seen, audit)
    elif isinstance(value, ArrayObject):
        for item in value:
            walk_pdf(item, clean, seen, audit)


def bitmap_equal(left,right):
    """Compare visible pixels without copying complete bitmaps through Pillow."""
    if (left.width,left.height,left.format)!=(right.width,right.height,right.format):return False
    try:
        import numpy as np
    except ImportError:
        return left.to_pil().tobytes()==right.to_pil().tobytes()
    # to_numpy excludes padding and exposes a view into each live PDFium bitmap.
    return bool(np.array_equal(left.to_numpy(),right.to_numpy()))


def clean_and_check(raw, output, timings=None, audit=None):
    global LINK_WORKSPACE
    LINK_WORKSPACE=str(Path(raw).resolve().parent)
    if timings is None:
        timings = {}
    started = time.monotonic()
    reader = PdfReader(raw, strict=True)
    if reader.is_encrypted or not reader.pages:
        raise ValueError('PDF is encrypted or contains no pages.')
    # This pipeline accepts simple printer output only. Unexpected interactive
    # features are rejected, rather than silently changing the document.
    if audit is not None:
        audit['pdf_property_fields_removed'] = len(reader.metadata or {}) if STRIP_METADATA else 0
        standard={'/Title','/Author','/Subject','/Keywords','/Creator','/Producer','/CreationDate','/ModDate','/Trapped'}
        names=set(reader.metadata or {})
        known=sorted(names & standard);unknown=len(names-standard)
        if unknown:known.append(f'{unknown} non-standard field(s)')
        audit['pdf_property_names_removed'] = '; '.join(known) if STRIP_METADATA else ''
    walk_pdf(reader.trailer['/Root'], clean=True, audit=audit)
    writer = PdfWriter(clone_from=reader)
    if STRIP_METADATA:
        writer.metadata = None
    walk_pdf(writer.root_object, clean=True, audit=audit)
    with open(output, 'xb') as stream:
        writer.write(stream)
    writer.close()
    checked = PdfReader(output, strict=True)
    if STRIP_METADATA and (checked.metadata or checked.xmp_metadata):
        raise ValueError('PDF metadata was not removed.')
    walk_pdf(checked.trailer['/Root'])
    timings['cleanup_and_structure_checks'] = time.monotonic()-started
    started = time.monotonic()
    # Ensure cleanup preserves every page's rendered appearance and text.
    with pdfium.PdfDocument(str(raw)) as before, pdfium.PdfDocument(str(output)) as after:
        if len(before) != len(after):
            raise ValueError('Page count changed during cleanup.')
        for i in range(len(before)):
            a, b = before[i], after[i]
            try:
                if a.get_size() != b.get_size():
                    raise ValueError(f'Page {i+1} size changed.')
                ta, tb = a.get_textpage(), b.get_textpage()
                try:
                    if ta.get_text_range() != tb.get_text_range():
                        raise ValueError(f'Page {i+1} text changed during cleanup.')
                finally:
                    ta.close(); tb.close()
                ba, bb = a.render(scale=1.5), b.render(scale=1.5)
                try:
                    if not bitmap_equal(ba,bb):
                        raise ValueError(f'Page {i+1} appearance changed during cleanup.')
                finally:
                    ba.close(); bb.close()
            finally:
                a.close(); b.close()
    timings['text_and_visual_verification'] = time.monotonic()-started
    return len(checked.pages)


def compare_reference(reference, cleaned, audit=None):
    """Compare page text and pixels with the pre-preparation final view."""
    with pdfium.PdfDocument(str(reference)) as before, pdfium.PdfDocument(str(cleaned)) as after:
        if len(before)!=len(after):
            raise ValueError('Appearance changed from original final view: page count differs.')
        for i in range(len(before)):
            a,b=before[i],after[i]
            try:
                if a.get_size()!=b.get_size():
                    raise ValueError(f'Appearance changed from original final view: page {i+1} size differs.')
                original_text,current_text=a.get_textpage(),b.get_textpage()
                try:
                    if original_text.get_text_range()!=current_text.get_text_range():
                        raise ValueError(f'Text changed from original reference on page {i+1}; PDF withheld and original retained.')
                finally:original_text.close();current_text.close()
                ba,bb=a.render(scale=2),b.render(scale=2)
                try:
                    if not bitmap_equal(ba,bb):
                        if audit is not None:
                            left,right=ba.to_pil(),bb.to_pil()
                            from PIL import ImageChops
                            diff=ImageChops.difference(left.convert('RGB'),right.convert('RGB'))
                            audit['appearance_difference']=dict(page=i+1,dpi=144,pixel_dimensions=left.size,changed_bounds=diff.getbbox())
                        raise ValueError(f'Appearance changed from original final view on page {i+1}; PDF withheld and original retained.')
                finally:
                    ba.close(); bb.close()
            finally:
                a.close(); b.close()


def choose_folders():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    app = tk.Tk(); app.withdraw()
    root = filedialog.askdirectory(title='Choose ROOT folder containing Word documents')
    if not root:
        return None
    archive = filedialog.askdirectory(title='Choose ORIGINAL WORD DOCUMENTS archive folder (outside root)', mustexist=False)
    if not archive:
        return None
    root, archive = validate_roots(Path(root), Path(archive))
    count = sum(1 for _ in documents(root))
    if not messagebox.askokcancel('Convert Word documents',
        f'Found {count} Word documents.\n\nRoot: {root}\n\nOriginals archive: {archive}\n\n'
        'PDFs will be saved beside the originals. After each PDF passes checks, its unchanged original will be moved to the archive. '
        'Existing PDFs and archive files will be skipped.\n\n'
        'The final version of tracked changes will print, without markup. Close Word before continuing.\n\nStart?'):
        return None
    app.destroy()
    return root, archive


def main():
    global SETTINGS_FILE, STRIP_METADATA, PROGRESS_FILE, REVIEW_ROOT, QUEUE_NAME, DEBUG_OVERRIDE, ONLY_DOCUMENT
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path)
    parser.add_argument('--settings', type=Path)
    parser.add_argument('--progress', type=Path)
    parser.add_argument('--queue-name')
    parser.add_argument('--debug',action=argparse.BooleanOptionalAction,default=None)
    parser.add_argument('--preserve-root', action='store_true')
    parser.add_argument('--archive', type=Path)
    parser.add_argument('--preview', action='store_true', help='List planned paths without changing files.')
    parser.add_argument('--only',type=Path,help=argparse.SUPPRESS)
    parser.add_argument('--timeout', type=int, default=180, help='Maximum seconds per Word operation.')
    args = parser.parse_args()
    PROGRESS_FILE = args.progress
    QUEUE_NAME = args.queue_name
    DEBUG_OVERRIDE = args.debug
    if args.settings:
        SETTINGS_FILE = args.settings.resolve()
        STRIP_METADATA = json.loads(SETTINGS_FILE.read_text(encoding='utf-8-sig')).get('stripMetadata', True)
    if bool(args.root) != bool(args.archive):
        parser.error('Supply both --root and --archive, or neither for folder pickers.')
    folders = (args.root, args.archive) if args.root else choose_folders()
    if folders:
        root, originals = validate_roots(*folders)
        REVIEW_ROOT = originals.parent / '#REVIEW' / root.name
        if args.preserve_root:
            root, archive = validate_roots(*folders)
            if not root.name:
                raise ValueError('Select a named folder, not an entire drive.')
            folders = root, archive / root.name
            if root.name.casefold()=='#review':
                raise ValueError('#REVIEW is reserved. Choose a differently named starting folder.')
        ONLY_DOCUMENT=None
        if args.only:
            ONLY_DOCUMENT=args.only.resolve()
            if not inside(ONLY_DOCUMENT,root.resolve()) or ONLY_DOCUMENT.suffix.lower() not in EXTENSIONS:
                raise ValueError('The one-document conversion target must be a Word file inside Starting.')
        sys.modules['convert'] = sys.modules[__name__]
        from pipeline import batch as export_batch
        export_batch(*folders, timeout=args.timeout, preview=args.preview)


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        traceback.print_exc()
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(1)
