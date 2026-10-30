import json
import zipfile
from pathlib import Path
from lxml import etree
from mathdocs.word import build_word
from mathdocs.schema import validate_document

NS={"m":"http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def test_native_editable_math_and_rtl(tmp_path):
    p=Path(__file__).resolve().parents[1]/"examples"/"lesson_limits.json"
    data=validate_document(json.loads(p.read_text(encoding="utf-8")))
    docpath=build_word(data,tmp_path/"test.docx")
    with zipfile.ZipFile(docpath) as z:
        xml=etree.fromstring(z.read("word/document.xml"))
        assert len(xml.xpath("//m:oMathPara",namespaces=NS)) >= 4
        assert len(xml.xpath("//m:f",namespaces=NS)) >= 3
        assert len(xml.xpath("//m:m/m:mr",namespaces=NS))==3
        assert len(xml.xpath("//m:m/m:mr/m:e",namespaces=NS))==9
        assert len(xml.xpath("//w:bidi",namespaces=NS))>=5
        assert not any(p.startswith("word/media/") for p in z.namelist()), "Equations must not be images"

def test_no_api_dependency_for_word(tmp_path,monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    p=Path(__file__).resolve().parents[1]/"examples"/"lesson_limits.json"
    data=json.loads(p.read_text(encoding="utf-8"))
    assert build_word(data,tmp_path/"offline.docx").exists()


def test_output_is_not_overwritten(tmp_path):
    import pytest
    original=tmp_path/"existing.docx"
    original.write_bytes(b"KEEP")
    p=Path(__file__).resolve().parents[1]/"examples"/"lesson_limits.json"
    data=json.loads(p.read_text(encoding="utf-8"))
    with pytest.raises(FileExistsError):
        build_word(data,original)
    assert original.read_bytes()==b"KEEP"
