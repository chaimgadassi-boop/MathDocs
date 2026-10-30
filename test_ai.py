import pytest
from mathdocs.ai_optional import request_summary

def test_ai_does_not_run_by_default():
    with pytest.raises(RuntimeError, match="disabled"):
        request_summary("secret material")
