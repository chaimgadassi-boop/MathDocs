"""Create .docx with native editable OMML equations and right-to-left paragraphs."""
from __future__ import annotations
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .schema import validate_document
from .omml import make_equation

def _rtl(paragraph):
    pPr=paragraph._p.get_or_add_pPr()
    flag=OxmlElement("w:bidi"); flag.set(qn("w:val"),"1")
    pPr.append(flag)
    paragraph.alignment=WD_ALIGN_PARAGRAPH.RIGHT

def _font(run,*,bold=False,size=11,color=None):
    run.font.name="Arial"
    run.font.size=Pt(size)
    run.bold=bold
    if color:
        run.font.color.rgb=RGBColor(*color)
    rPr=run._r.get_or_add_rPr()
    fonts=rPr.rFonts
    if fonts is not None:
        fonts.set(qn("w:eastAsia"),"Arial")
        fonts.set(qn("w:cs"),"Arial")
    rtl=OxmlElement("w:rtl"); rtl.set(qn("w:val"),"1")
    rPr.append(rtl)

def _paragraph(doc,text,*,bold=False,color=None):
    p=doc.add_paragraph()
    _rtl(p); p.paragraph_format.space_after=Pt(7)
    r=p.add_run(text); _font(r,bold=bold,color=color)
    return p

def build_word(data:dict, out_path:Path)->Path:
    validate_document(data)
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(2.3)
    sec.bottom_margin=Cm(2.2)
    sec.left_margin=Cm(2.4)
    sec.right_margin=Cm(2.4)
    styles=doc.styles
    normal=styles["Normal"]; normal.font.name="Arial"; normal.font.size=Pt(11)
    normal.paragraph_format.space_after=Pt(6)
    doc.core_properties.title=data["title"]
    p=doc.add_paragraph()
    _rtl(p)
    p.paragraph_format.space_after=Pt(20)
    _font(p.add_run(data["title"]),bold=True,size=20,color=(22,49,74))
    _paragraph(doc,"MathDocs • מסמך מתמטי עם נוסחאות Word מקוריות לעריכה",color=(89,103,119))
    for si,section in enumerate(data["sections"],1):
        p=doc.add_paragraph()
        _rtl(p)
        p.paragraph_format.space_before=Pt(12)
        p.paragraph_format.space_after=Pt(10)
        _font(p.add_run(f"{si}. {section['heading']}"),bold=True,size=14,color=(22,70,110))
        for block in section["blocks"]:
            kind=block["kind"]
            if kind == "paragraph":
                _paragraph(doc,block["text"])
            elif kind=="note":
                _paragraph(doc,"חשוב: "+block["text"],bold=True,color=(116,78,24))
            elif kind=="equation":
                p=doc.add_paragraph()
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before=Pt(8)
                p.paragraph_format.space_after=Pt(12)
                p._p.append(make_equation(block["expr"]))
            elif kind=="table":
                rows=block["rows"]
                table=doc.add_table(rows=len(rows),cols=len(rows[0]))
                table.style="Light Shading Accent 1"
                table.alignment=WD_TABLE_ALIGNMENT.CENTER
                # A right-to-left table keeps first column on the right.
                tblPr=table._tbl.tblPr
                bidi=OxmlElement("w:bidiVisual"); bidi.set(qn("w:val"),"1")
                tblPr.append(bidi)
                for ri,row in enumerate(rows):
                    for ci,cell in enumerate(row):
                        c=table.cell(ri,ci)
                        c.text=""
                        p=c.paragraphs[0]; _rtl(p)
                        _font(p.add_run(cell),bold=(ri==0),size=10)
                doc.add_paragraph()
    out_path=Path(out_path)
    out_path.parent.mkdir(parents=True,exist_ok=True)
    if out_path.exists():
        raise FileExistsError(f"Output already exists; choose a new file name: {out_path}")
    doc.save(out_path)
    return out_path
