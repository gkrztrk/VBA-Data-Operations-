Attribute VB_Name = "Copy_Materials"
Sub Material_List_Combination()

fastSettings (False)
loopSheets
fastSettings (True)

End Sub
Sub loopSheets()
Dim destLastRow As Long
Dim m3HesapFirstRow As Long
Dim i As Long
Dim wss As Worksheet
Dim wsd As Worksheet
Dim wbd As Workbook
Dim wbs As Workbook
Dim destPath As String
Dim sourcePath As String
Dim sourceWbPath As String
Dim sourceWbName As String
Dim rngDlt As Range
Dim a As Integer

'Application.DisplayAlerts = False
destPath = ThisWorkbook.Worksheets("Data").Cells(1, 2) & "\" & ThisWorkbook.Worksheets("Data").Cells(2, 2)

Set wbd = Workbooks.Open(destPath)
Set wsd = wbd.Worksheets("All")

a = 4
sourceWbPath = ThisWorkbook.Worksheets("Data").Cells(3, 2)

Do While ThisWorkbook.Worksheets("Data").Cells(a, 2) <> ""
    sourceWbName = ThisWorkbook.Worksheets("Data").Cells(a, 2)
    Set wbs = Workbooks.Open(sourceWbPath & "\" & sourceWbName)
    
    For Each wss In wbs.Worksheets
    
    
        wss.Activate
        destLastRow = wsd.Cells(Rows.Count, 3).End(xlUp).Row
        Call copySheets(wsd, destLastRow)
        
    Next
    
    wbs.Close
    
    a = a + 1
Loop

'//////////////MASAR LIST COPY//////////////////////////////////////////////////

    sourceWbName = ThisWorkbook.Worksheets("Data").Cells(1, 5)
    If sourceWbName <> "" Then
        Set wbs = Workbooks.Open(sourceWbPath & "\" & sourceWbName)
        Set wss = wbs.Worksheets(ThisWorkbook.Worksheets("Data").Cells(1, 6).Value)
    
        wss.Activate
            destLastRow = wsd.Cells(Rows.Count, 3).End(xlUp).Row
            Call copySheets(wsd, destLastRow)
    
            wbs.Close
    End If
    
        
'//////////////////////////////////////////////////////////////////////////////

destLastRow = wsd.Cells(Rows.Count, 3).End(xlUp).Row
Set rngDlt = wsd.Range("C" & destLastRow & ":C" & Rows.Count)
On Error Resume Next
rngDlt.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
'wsd.UsedRange.RemoveDuplicates Columns:=Array(2, 3), Header:=xlYes
m3HesapFirstRow = wsd.Cells(Rows.Count, 14).End(xlUp).Row

For i = m3HesapFirstRow To destLastRow
    
    Call Nesma_SC_Other(i, wsd)
    Call m3_hesap(i, wbd)
    Call amount(i, wbd)
    
    
Next i

wbd.Close (True)

End Sub

Sub copySheets(destSheet As Worksheet, destLastRow As Long)

Dim rowNo As Long

rowNo = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row

ActiveSheet.Range("A2:L" & rowNo).Copy
destSheet.Cells(destLastRow, 1).PasteSpecial xlValues
destSheet.Range("B1:B" & destSheet.Cells(Rows.Count, 3).End(xlUp).Row).NumberFormat = "dd.mm.yy"
destSheet.UsedRange.RemoveDuplicates Columns:=Array(2, 3, 4, 5, 6, 7, 8, 9), Header:=xlYes
End Sub

Sub Nesma_SC_Other(rowNo As Long, ws As Worksheet)

    If IsInArray2(ws.Cells(rowNo, 8), Array("ZONE-5 CRUSHER", "AL GHARBI FAYHA")) Then
    
        ws.Cells(rowNo, 15) = "AL GHARBI"
    ElseIf IsInArray2(ws.Cells(rowNo, 8), Array("MAKKAH CRUSHER", "FAYHA CRUSHER", "MASAR")) Then
    
        ws.Cells(rowNo, 15) = "Nesma"
    Else
        
        ws.Cells(rowNo, 15) = "OTHER"
    End If
    

End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
Public Function IsInArray2(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray2 = True
            Exit Function
        End If
    Next i
    IsInArray2 = False

End Function

Sub fastSettings(OnOff As Boolean)

Application.DisplayAlerts = OnOff
Application.ScreenUpdating = OnOff
Application.EnableEvents = OnOff

If OnOff = True Then

    Application.Calculation = xlCalculationAutomatic
    
Else

    Application.Calculation = xlCalculationManual
    
End If



End Sub
