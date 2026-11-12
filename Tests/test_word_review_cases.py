"""Real Word checks for the read-only review fixes. Uses generated TEMP fixtures."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from PIL import Image
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import convert
import pipeline
from fieldcache import freeze_saved_fields
from pypdf import PdfReader


def revision_doc(path,visible):
    doc=Document();doc.add_paragraph('Body text remains.')
    ins=OxmlElement('w:ins');ins.set(qn('w:id'),'12');ins.set(qn('w:author'),'Review Test')
    run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='HEADER_INSERTION_MARKER';run.append(text);ins.append(run)
    doc.sections[0].header.paragraphs[0]._p.append(ins)
    view=OxmlElement('w:revisionView');view.set(qn('w:markup'),'1' if visible else '0');doc.settings.element.append(view)
    doc.save(path)


def process_alive(pid):
    import ctypes
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    kernel.OpenProcess.argtypes=[ctypes.c_uint32,ctypes.c_bool,ctypes.c_uint32];kernel.OpenProcess.restype=ctypes.c_void_p
    kernel.WaitForSingleObject.argtypes=[ctypes.c_void_p,ctypes.c_uint32];kernel.WaitForSingleObject.restype=ctypes.c_uint32
    kernel.CloseHandle.argtypes=[ctypes.c_void_p]
    handle=kernel.OpenProcess(0x100000,False,pid)
    if not handle:return False
    try:return kernel.WaitForSingleObject(handle,0)==258
    finally:kernel.CloseHandle(handle)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--orphan-child',type=Path);args=parser.parse_args()
    if args.orphan_child:
        folder=args.orphan_child
        session=pipeline.WordSession(folder,folder/'settings.json',40)
        print(session.ready['pid'],flush=True)
        os._exit(0)
    folder=Path(tempfile.mkdtemp(prefix='word-pdf-word-review-'));source=folder/'Source';source.mkdir();workspace=folder/'Workspace';workspace.mkdir()
    settings=workspace/'settings.json';settings.write_text(json.dumps(pipeline.DEFAULTS))
    revision_doc(source/'HeaderVisible.docx',True);revision_doc(source/'HeaderHidden.docx',False)
    doc=Document();doc.add_paragraph('Body text remains.')
    field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),' FILENAME \\p ');run=OxmlElement('w:r');text=OxmlElement('w:t');text.text=str(source/'FullPath.docx');run.append(text);field.append(run);doc.sections[0].footer.paragraphs[0]._p.append(field);doc.save(source/'FullPath.docx')
    doc=Document();paragraph=doc.add_paragraph('Body text remains.')
    props=paragraph._p.get_or_add_pPr();runs=OxmlElement('w:rPr');runs.append(OxmlElement('w:vanish'));props.append(runs);doc.save(source/'HiddenTerminal.docx')
    doc=Document();doc.add_paragraph('Protected body text remains.')
    protection=OxmlElement('w:documentProtection');protection.set(qn('w:edit'),'readOnly');protection.set(qn('w:enforcement'),'1');doc.settings.element.append(protection);doc.save(source/'ProtectedClean.docx')
    doc=Document();doc.add_paragraph('Linked text source.');doc.save(source/'linked.docx')
    doc=Document();field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),' INCLUDETEXT "linked.docx" ')
    run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='CACHED_LINKED_MARKER';run.append(text);field.append(run);doc.add_paragraph()._p.append(field);doc.save(source/'CachedLinkedText.docx')
    Image.new('RGB',(40,40),(255,0,0)).save(source/'diagram.png')
    doc=Document();doc.add_paragraph('Linked image.');doc.add_picture(str(source/'diagram.png'));picture=source/'UncachedImage.docx';doc.save(picture)
    with zipfile.ZipFile(picture) as package:contents={name:package.read(name) for name in package.namelist()}
    xml=etree.fromstring(contents['word/document.xml']);blip=xml.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip');relationship=blip.get(qn('r:embed'));del blip.attrib[qn('r:embed')];blip.set(qn('r:link'),relationship);contents['word/document.xml']=etree.tostring(xml)
    xml=etree.fromstring(contents['word/_rels/document.xml.rels'])
    for link in xml:
        if link.get('Id')==relationship:link.set('Target','diagram.png');link.set('TargetMode','External')
    contents['word/_rels/document.xml.rels']=etree.tostring(xml)
    with zipfile.ZipFile(picture,'w',zipfile.ZIP_DEFLATED) as package:
        for name,data in contents.items():package.writestr(name,data)
    doc=Document();doc.add_paragraph('Saved dynamic values:')
    for name,value in (('DATE','1999-12-31'),('TIME','01:02:03'),('PRINTDATE','1998-01-02'),('USERNAME','Original Author'),('USERINITIALS','OA'),('USERADDRESS','Original Address'),('CREATEDATE','1981-02-03'),('SAVEDATE','1982-03-04'),('FILESIZE','42')):
        field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),' '+name+' ');run=OxmlElement('w:r');text=OxmlElement('w:t');text.text=value;run.append(text);field.append(run);doc.add_paragraph()._p.append(field)
    doc.save(source/'DynamicValues.docx')
    hashes={p.name:convert.digest(p) for p in source.iterdir()}
    session=pipeline.WordSession(workspace,settings,40);word_pid=session.ready['pid']
    try:
        for name in ('HeaderVisible','HeaderHidden','FullPath','HiddenTerminal','ProtectedClean','UncachedImage','CachedLinkedText','DynamicValues'):
            original=source/(name+'.docx');copy=workspace/original.name;shutil.copyfile(original,copy)
            frozen=freeze_saved_fields(copy)
            pdf=workspace/(name+'.pdf')
            try:result=session.export(copy,pdf,lambda:None,original.name)
            except pipeline.WordDocumentError as error:
                if name not in ('FullPath','HiddenTerminal','UncachedImage'):raise
                expected={'FullPath':'Full-path FILENAME','HiddenTerminal':'hidden structural character','UncachedImage':'Uncached linked content'}[name]
                assert expected in str(error),(name,str(error))
                assert not pdf.exists(),name
                assert process_alive(word_pid),'Ordinary document error killed Word'
                print('PASS:',name,'withheld promptly; reusable Word remains open.',flush=True)
                continue
            assert name not in ('FullPath','UncachedImage'),(name,result)
            audit=result['audit']
            text='\n'.join(page.extract_text() for page in PdfReader(pdf).pages)
            if name=='FullPath':
                assert ''.join(str(original).split()) in ''.join(text.split()),text
                assert ''.join(str(workspace).split()) not in ''.join(text.split()),text
            if name=='DynamicValues':
                for saved in ('1999-12-31','01:02:03','1998-01-02','Original Author','OA','Original Address','1981-02-03','1982-03-04','42'):assert saved in text,(saved,text)
            if name=='HeaderVisible':assert audit['revisionsFound']>=1 and audit['visibleRevisionsPreserved'],audit
            if name=='HeaderHidden':assert audit['revisionsAccepted']>=1 and not audit['visibleRevisionsPreserved'],audit
            clean=workspace/(name+'.clean.pdf');convert.clean_and_check(pdf,clean);convert.compare_reference(Path(str(pdf)+'.baseline.pdf'),clean)
            print('PASS:',name,'prepared, verified and counted correctly.',flush=True)
    finally:session.close()
    assert not (workspace/'options.json').exists(),'Normal shutdown did not restore Word options'
    for name,expected in hashes.items():assert convert.digest(source/name)==expected,name
    # Exercise forced shutdown and recovery without opening any document.
    forced=folder/'Forced';forced.mkdir();(forced/'settings.json').write_text(json.dumps(pipeline.DEFAULTS))
    session=pipeline.WordSession(forced,forced/'settings.json',40)
    assert (forced/'options.json').exists()
    session.close(force=True)
    assert not (forced/'options.json').exists(),'Forced shutdown option recovery did not finish'
    print('PASS: forced shutdown restored saved Word options.',flush=True)
    # A dead Python parent must not leave an idle hidden Word instance behind.
    orphan=folder/'ParentExit';orphan.mkdir();(orphan/'settings.json').write_text(json.dumps(pipeline.DEFAULTS))
    result=subprocess.run([sys.executable,__file__,'--orphan-child',str(orphan)],capture_output=True,text=True,timeout=60,check=True)
    pid=int(result.stdout.strip().splitlines()[-1]);deadline=time.monotonic()+20
    while process_alive(pid) and time.monotonic()<deadline:time.sleep(.1)
    assert not process_alive(pid),'Idle Word survived its Python parent'
    assert not (orphan/'options.json').exists(),'Parent exit did not restore options'
    print('PASS: parent exit shut down idle Word and restored options.',flush=True)
    print('Fixture folder:',folder,flush=True)


if __name__=='__main__':main()
