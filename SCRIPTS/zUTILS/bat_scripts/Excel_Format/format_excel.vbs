
'#####################################################################
if WScript.Arguments.Count < 1 Then
    WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file>"
    Wscript.Quit
End If

xlFile = objFSO.objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
Dim xlApp = CreateObject("Excel.Application")
Dim xlwb = xlApp.Workbooks.Open(src_file)

'Set cell formatting here.

xlwb.Close False
xlApp.Quit

