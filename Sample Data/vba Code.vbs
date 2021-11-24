Function Main(fileName As String, shtName As String) As String
    On Error GoTo errorOccured
    Dim wbk As Workbook
    Set wbk = Workbooks.Open(fileName)
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(shtName).Delete
    Application.DisplayAlerts = True
    Main = "Successful!"
    Exit Function
errorOccured:
    Main = "Failed!"
End Function

