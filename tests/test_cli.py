from mathdocs.cli import main


def test_demo_outside_project(tmp_path, monkeypatch, capsys):
    """The installed demo must not depend on the repository or current directory."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    output = tmp_path / "demo.docx"
    assert main(["demo", "--output", str(output)]) == 0
    assert output.is_file()
    assert "Created:" in capsys.readouterr().out
