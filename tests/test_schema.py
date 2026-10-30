from mathdocs.schema import validate_document, DocumentValidationError

def test_sample_is_valid():
    import json
    from pathlib import Path
    p=Path(__file__).resolve().parents[1]/"examples"/"lesson_limits.json"
    assert validate_document(json.loads(p.read_text(encoding="utf-8")))["version"]=="1"

def test_non_rectangular_matrix_rejected():
    d={"version":"1","title":"t","sections":[{"heading":"m","blocks":[{"kind":"equation","expr":{
      "type":"matrix","rows":[[{"type":"text","value":"1"}],[{"type":"text","value":"2"},{"type":"text","value":"3"}]]
    }}]}]}
    try:
        validate_document(d)
    except DocumentValidationError:
        pass
    else:
        raise AssertionError("Expected nonrectangular matrix validation error")
