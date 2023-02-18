Attribute VB_Name = "Daily_Report_Update"
Sub Daily_QTY_update()

Dim wss, wsd As Worksheet
Dim i As Integer
Dim wbs, wbd As Workbook
Dim reportDate As Date
Dim bulunan As Range
Dim dateRange As Range
Dim sourceCol, destCol As Integer

Dim sourceFileName, sourceWorksheetName, destinationFileName, destinationWorkseetName, sourcePath, destinationPath As String

Application.DisplayAlerts = False


sourceFileName = ThisWorkbook.Worksheets("Daily Report Update").Range("B1").Value
sourcePath = ThisWorkbook.Worksheets("Daily Report Update").Range("B2").Value & sourceFileName

Set wbs = Workbooks.Open(sourcePath, True)

destinationFileName = ThisWorkbook.Worksheets("Daily Report Update").Range("B3").Value
destinationPath = ThisWorkbook.Worksheets("Daily Report Update").Range("B4").Value & destinationFileName

Set wbd = Workbooks.Open(destinationPath, True)

Set wss = wbs.Worksheets("report")

Set wsd = wbd.Worksheets("Executed QTY")

Set dateRange = wsd.Range("J1:UY1")


reportDate = wss.Range("O1")

Set bulunan = dateRange.Find(reportDate)

destCol = bulunan.Column


If wsd.FilterMode Then

wsd.ShowAllData

End If

If wss.FilterMode Then

wss.ShowAllData

End If

If wbd.Worksheets("Report").FilterMode Then

    wbd.Worksheets("Report").ShowAllData
End If


i = 5

For i = 5 To 546


If wss.Cells(i, 16) <> "" Or wss.Cells(i, 16) <> 0 Then

    wsd.Cells(i, destCol) = wss.Cells(i, 16)
    wbd.Worksheets("Report").Cells(i, 16) = wss.Cells(i, 16)
    wbd.Worksheets("Report").Cells(i, 17) = wss.Cells(i, 17)
    
End If

Next i


wbd.Worksheets("Report").Range("A3:S546").AutoFilter Field:=12, Criteria1:=">0", _
    Operator:=xlOr, Criteria2:=""


wbs.Close (False)
wbd.Close (True)




Application.DisplayAlerts = True






End Sub
