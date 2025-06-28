import tempfile
import win32com.client as win32

def break_external_links(input_path: str) -> str:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    tmp_path = tmp.name

    wb = excel.Workbooks.Open(Filename=input_path, UpdateLinks=0, ReadOnly=True)
    links = wb.LinkSources(Type=1)
    if links:
        for link in links:
            wb.BreakLink(Name=link, Type=1)
    wb.SaveAs(tmp_path, FileFormat=51)  # 明确保存为 .xlsx
    wb.Close()
    excel.Quit()
    return tmp_path
