"""Freeze saved environment-dependent Word field values in TEMP DOCX/DOCM copies."""
from pathlib import Path
import re
import zipfile
import os
from fileops import native_path

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DYNAMIC={'DATE','TIME','PRINTDATE','USERNAME','USERINITIALS','USERADDRESS','CREATEDATE','SAVEDATE','FILESIZE'}
STORY=re.compile(r'^word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments|glossary/document)\.xml$')


def field_name(code):
    match=re.match(r'\s*([A-Za-z]+)\b',code or '')
    return match[1].upper() if match else ''


def freeze_saved_fields(path):
    """Set only field-lock/refresh flags; never rewrite field results or source files."""
    path=Path(path)
    if path.suffix.lower() not in ('.docx','.docm'):return {}
    from lxml import etree
    changed={};types=set();locked=0
    parser=etree.XMLParser(resolve_entities=False,no_network=True,remove_blank_text=False,strip_cdata=False)
    if not zipfile.is_zipfile(native_path(path)):
        return dict(fieldCacheInspection='Non-ZIP Word input; dynamic fields must already be locked')
    with zipfile.ZipFile(native_path(path)) as source:
        for entry in source.infolist():
            if not STORY.fullmatch(entry.filename):continue
            raw=source.read(entry)
            if b'fldSimple' not in raw and b'fldChar' not in raw:continue
            root=etree.fromstring(raw,parser);stack=[];part_changed=False
            def hold(element,name,has_cache):
                nonlocal locked,part_changed
                if name not in DYNAMIC:return
                if not has_cache:raise ValueError('A dynamic '+name+' field has no saved display value; original appearance cannot be guaranteed.')
                types.add(name)
                if element.get('{'+W+'}fldLock') not in ('true','1','on'):
                    element.set('{'+W+'}fldLock','true');locked+=1;part_changed=True
                if element.get('{'+W+'}dirty') not in (None,'false','0','off'):
                    element.set('{'+W+'}dirty','false');part_changed=True
            for element in root.iter():
                if element.tag=='{'+W+'}fldSimple':
                    name=field_name(element.get('{'+W+'}instr',''))
                    hold(element,name,any(True for tag in ('t','delText') for _ in element.iter('{'+W+'}'+tag)))
                elif element.tag=='{'+W+'}fldChar':
                    kind=element.get('{'+W+'}fldCharType')
                    if kind=='begin':stack.append(dict(begin=element,code=[],result=False,cached=False))
                    elif kind=='separate' and stack:stack[-1]['result']=True
                    elif kind=='end' and stack:
                        field=stack.pop();hold(field['begin'],field_name(''.join(field['code'])),field['cached'])
                elif element.tag in ('{'+W+'}instrText','{'+W+'}delInstrText') and stack and not stack[-1]['result']:stack[-1]['code'].append(element.text or '')
                elif element.tag in ('{'+W+'}t','{'+W+'}delText') and stack and stack[-1]['result']:stack[-1]['cached']=True
            if part_changed:
                changed[entry.filename]=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=root.getroottree().docinfo.standalone)
        signed=any(name.lower().startswith('_xmlsignatures/') for name in source.namelist())
        if changed and '[Content_Types].xml' in source.namelist():
            content_types=etree.fromstring(source.read('[Content_Types].xml'),parser)
            signed=signed or any('openxmlformats-package.digital-signature-' in node.get('ContentType','').lower() for node in content_types)
        if changed and signed:
            raise ValueError('Signed Word package contains dynamic fields; cannot freeze their display without invalidating the signature.')
        if changed:
            temporary=path.with_name(path.name+'.fields.tmp')
            try:
                with zipfile.ZipFile(native_path(temporary),'x') as target:
                    for entry in source.infolist():target.writestr(entry,changed.get(entry.filename,source.read(entry)))
                # Close the reader before replacing the TEMP copy on Windows.
            except BaseException:
                if os.path.exists(native_path(temporary)):os.unlink(native_path(temporary))
                raise
    if changed:
        import stat
        import shutil
        try:shutil.copystat(native_path(path),native_path(temporary))
        except OSError:pass # Field values are frozen; TEMP filesystem timestamps are nonessential.
        mode=os.stat(native_path(path)).st_mode
        if os.name=='nt' and not mode & stat.S_IWRITE:os.chmod(native_path(path),mode | stat.S_IWRITE)
        os.replace(native_path(temporary),native_path(path))
    return dict(dynamicFieldsLocked=locked,dynamicFieldTypes='; '.join(sorted(types)))
