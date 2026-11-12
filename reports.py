"""Excel reports built from durable CSV-format journals after processing finishes.

Uses only Python's standard library. Every text cell is literal, never a formula.
"""
import csv
import os
from pathlib import Path
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from fileops import copy_verified

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NUMERIC_HEADERS={'Eligible Word files','Other files','Source custom properties found','Comments removed','Revisions accepted','Hidden ranges removed','PDF property fields removed','XMP streams removed','Hyperlinks removed','Private app dictionaries removed','Object dates removed'}
MAX_ROWS=1048576


def native(path):
    value=os.path.abspath(path)
    if os.name=='nt' and not value.startswith('\\\\?\\'):
        return '\\\\?\\UNC\\'+value[2:] if value.startswith('\\\\') else '\\\\?\\'+value
    return value


def text_xml(value):
    value=re.sub(r'_x[0-9a-fA-F]{4}_',lambda m:'_x005F_'+m[0][1:],str(value))
    value=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]',lambda m:f'_x{ord(m[0]):04X}_',value)
    return escape(value)


def column(number):
    result=''
    while number:
        number,remainder=divmod(number-1,26); result=chr(65+remainder)+result
    return result


def report_rows(path):
    """Read current Excel reports, recovery journals, or historical CSV reports."""
    path=Path(path)
    csv.field_size_limit(16*1024*1024)
    if path.suffix.lower()!='.xlsx':
        with open(native(path),newline='',encoding='utf-8-sig') as stream:
            yield from csv.DictReader(stream)
        return
    with zipfile.ZipFile(native(path)) as package:
        shared=[]
        if 'xl/sharedStrings.xml' in package.namelist():
            shared=[''.join(node.itertext()) for node in ET.fromstring(package.read('xl/sharedStrings.xml'))]
        sheets=sorted((n for n in package.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)),key=lambda n:int(re.search(r'(\d+)\.xml',n)[1]))
        for sheet in sheets:
            headers=None
            with package.open(sheet) as stream:
                for event,row in ET.iterparse(stream,events=('end',)):
                    if row.tag!=f'{{{NS}}}row': continue
                    values=[]
                    for cell in row:
                        col=re.match(r'[A-Z]+',cell.attrib['r'])[0]; index=0
                        for letter in col:index=index*26+ord(letter)-64
                        while len(values)<index:values.append('')
                        kind=cell.get('t')
                        if kind=='inlineStr':value=''.join(t.text or '' for t in cell.iter(f'{{{NS}}}t'))
                        else:
                            value=cell.findtext(f'{{{NS}}}v','')
                            if kind=='s':value=shared[int(value)]
                        values[index-1]=re.sub(r'_x([0-9a-fA-F]{4})_',lambda m:chr(int(m[1],16)),value)
                    if headers is None:headers=values
                    else:yield dict(zip(headers,values+['']*max(0,len(headers)-len(values))))
                    row.clear()


def build_excel(journal,output,kind='Report'):
    """Publish a complete, CRC-checked workbook; retain the journal."""
    journal,output=Path(journal),Path(output)
    csv.field_size_limit(16*1024*1024)
    os.makedirs(native(output.parent),exist_ok=True)
    handle,temp=tempfile.mkstemp(prefix='word-pdf-report-',suffix='.xlsx')
    os.close(handle)
    sheets=[]; row_count=0; strings=[]; string_indices={}
    try:
        with open(native(journal),newline='',encoding='utf-8-sig') as source,zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as book:
            reader=csv.reader(source); headers=next(reader,[])
            if not headers:raise ValueError('Report journal has no header row.')
            if len(headers)>16384:raise ValueError('Report exceeds Excel column limit.')
            last=column(len(headers)); sheet=None; sheet_row=0
            def write_row(values,number,header=False):
                if len(values)>len(headers):raise ValueError('Report row has more values than headings.')
                style=0
                if header:style=1
                elif any('FAILED' in str(v).upper() or 'NOT RESTORED' in str(v).upper() or 'SOURCE RETAINED' in str(v).upper() for h,v in zip(headers,values) if h in ('Status','Outcome')):style=2
                elif any(str(v) for h,v in zip(headers,values) if h in ('Reasons for review','Review reasons','Scan error')):style=3
                sheet.write(f'<row r="{number}" ht="{32 if header else 42}" customHeight="1">'.encode())
                for i,value in enumerate(values):
                    value=str(value)
                    if len(value.encode('utf-16-le'))//2>32767:
                        locator=f' [Truncated; full text in recovery folder {journal.parent.name}, report {kind}, data row {row_count+1}, column {column(i+1)}.]'
                        budget=32767-len(locator.encode('utf-16-le'))//2
                        value=value.encode('utf-16-le')[:max(0,budget)*2].decode('utf-16-le','ignore')+locator
                    address=f'{column(i+1)}{number}'
                    numeric=not header and headers[i] in NUMERIC_HEADERS and re.fullmatch(r'-?\d{1,15}',value)
                    if numeric:
                        xml=f'<c r="{address}" s="{style}"><v>{value}</v></c>'
                    else:
                        if value not in string_indices:
                            string_indices[value]=len(strings);strings.append(value)
                        xml=f'<c r="{address}" s="{style}" t="s"><v>{string_indices[value]}</v></c>' 
                    sheet.write(xml.encode('utf-8'))
                sheet.write(b'</row>')
            def start_sheet():
                nonlocal sheet,sheet_row
                number=len(sheets)+1; name=(kind.replace('-',' ').title()[:25] or 'Report')+(f' {number}' if number>1 else '')
                name=re.sub(r'[\\/*?:\[\]]',' ',name)[:31]
                sheets.append(name)
                sheet=book.open(f'xl/worksheets/sheet{number}.xml','w');sheet_row=1
                sheet.write((f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="42"/><cols>').encode())
                for i,h in enumerate(headers,1):
                    width=60 if h in ('Source','Document','File','PDF','Archive','Folder','Review copy') else 48 if any(w in h.lower() for w in ('reason','detail','notes','error','timing','types')) else min(36,max(16,len(h)+2))
                    sheet.write(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>'.encode())
                sheet.write(b'</cols><sheetData>');write_row(headers,1,True)
            def finish_sheet():
                sheet.write(f'</sheetData><autoFilter ref="A1:{last}{sheet_row}"/></worksheet>'.encode());sheet.close()
            try:
                start_sheet()
                for values in reader:
                    if sheet_row>=MAX_ROWS:finish_sheet();start_sheet()
                    sheet_row+=1;write_row(values,sheet_row);row_count+=1
                finish_sheet();sheet=None
            finally:
                if sheet is not None:sheet.close()
            with book.open('xl/sharedStrings.xml','w') as shared:
                shared.write(f'<sst xmlns="{NS}" uniqueCount="{len(strings)}">'.encode())
                for value in strings:shared.write(f'<si><t xml:space="preserve">{text_xml(value)}</t></si>'.encode('utf-8'))
                shared.write(b'</sst>')
            book.writestr('[Content_Types].xml','<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'+''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(sheets)+1))+'</Types>')
            book.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            book.writestr('xl/workbook.xml',f'<workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i,name in enumerate(sheets,1))+'</sheets></workbook>')
            book.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))+f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId{len(sheets)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
            book.writestr('xl/styles.xml',f'<styleSheet xmlns="{NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF215AC8"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4">'+''.join(f'<xf numFmtId="0" fontId="{1 if i==1 else 0}" fillId="{[0,2,3,4][i]}" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' for i in range(4))+'</cellXfs>'+'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
        with zipfile.ZipFile(temp) as book:
            if book.testzip() is not None:raise ValueError('Excel report ZIP verification failed.')
            for name in book.namelist():
                with book.open(name) as stream:
                    parents=[]
                    for event,element in ET.iterparse(stream,events=('start','end')):
                        if event=='start':parents.append(element)
                        else:
                            parents.pop()
                            if parents:parents[-1].remove(element)
                            element.clear()
        # Publish only after the complete workbook has been verified.
        # Report outputs are logs, not Word/PDF conversion temporary files.
        copy_verified(temp,output)
        return row_count
    finally:
        os.unlink(temp)
