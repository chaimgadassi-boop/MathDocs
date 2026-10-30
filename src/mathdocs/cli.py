"""MathDocs CLI: offline by default; no API key needed."""
from __future__ import annotations
import argparse
from importlib.resources import as_file, files
from pathlib import Path
from .schema import read_document
from .word import build_word

def main(argv=None)->int:
    parser=argparse.ArgumentParser(description="MathDocs Local — editable Word math")
    cmd=parser.add_subparsers(dest="command",required=True)
    build=cmd.add_parser("build",help="Build DOCX from structured JSON")
    build.add_argument("input",type=Path)
    build.add_argument("--output",type=Path,default=Path("output/mathdocs.docx"))
    demo=cmd.add_parser("demo",help="Create editable Word demo without AI")
    demo.add_argument("--output",type=Path,default=Path("output/mathdocs_demo.docx"))
    excel=cmd.add_parser("excel-demo",help="Create Excel COM demo (Windows + Excel only)")
    excel.add_argument("--output",type=Path,default=Path("output/mathdocs_demo.xlsx"))
    args=parser.parse_args(argv)
    if args.command=="demo":
        sample=files("mathdocs").joinpath("data/lesson_limits.json")
        with as_file(sample) as sample_path:
            out=build_word(read_document(sample_path),args.output)
    elif args.command=="build":
        out=build_word(read_document(args.input),args.output)
    else:
        from .excel_windows import create_interactive_excel
        out=create_interactive_excel(args.output)
    print(f"Created: {out}")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
