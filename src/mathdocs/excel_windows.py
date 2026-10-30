"""Optional Windows-only Excel COM integration. No network or AI calls."""
from __future__ import annotations
from pathlib import Path
import os
import sys

def create_interactive_excel(path: Path) -> Path:
    """Needs Microsoft Excel desktop and pywin32. Does not run headlessly on a server."""
    if sys.platform != "win32":
        raise RuntimeError("Excel COM requires Windows with Microsoft Excel desktop.")
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("Install the optional Windows dependency: pip install -e '.[windows]'") from exc
    path=Path(path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        raise FileExistsError(f"Output already exists: {path}")
    excel=win32.DispatchEx("Excel.Application")
    excel.Visible=False
    excel.DisplayAlerts=False
    wb=None
    try:
        wb=excel.Workbooks.Add()
        matrix=wb.Worksheets(1); matrix.Name="Matrix"
        matrix.Range("A1").Value="Editable 3x3 Matrix"
        for r,row in enumerate([[1,2,3],[0,4,5],[1,0,6]],3):
            for c,v in enumerate(row,2):
                matrix.Cells(r,c).Value=v
        matrix.Range("F3").Value="Determinant"
        matrix.Range("G3").Formula="=MDETERM(B3:D5)"
        matrix.Range("F4").Value="Trace"
        matrix.Range("G4").Formula="=B3+C4+D5"
        matrix.Range("B3:D5").Interior.Color=0xF3E5CF
        matrix.Columns("A:G").AutoFit()
        graph=wb.Worksheets.Add(After=matrix); graph.Name="Graph"
        graph.Range("A1").Value="x"
        graph.Range("B1").Value="f(x)=(2x²+5x)/(4x²-3)"
        for r,x in enumerate(range(2,32),2):
            graph.Cells(r,1).Value=x
            graph.Cells(r,2).Formula=f"=(2*A{r}^2+5*A{r})/(4*A{r}^2-3)"
        chart_obj=graph.ChartObjects().Add(260,55,550,330)
        chart=chart_obj.Chart
        chart.ChartType=74 # xlXYScatterLines
        series=chart.SeriesCollection().NewSeries()
        series.XValues=graph.Range("A2:A31")
        series.Values=graph.Range("B2:B31")
        series.Name="f(x)"
        chart.HasTitle=True
        chart.ChartTitle.Text="Interactive function graph"
        graph.Columns("A:B").AutoFit()
        # 51: xlOpenXMLWorkbook
        wb.SaveAs(str(path),FileFormat=51)
        return path
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
