Attribute VB_Name = "Edit_Files"
Sub Edit_Materials()
    Call AL_Gharbi_Zone_5_Edit
    

End Sub

Sub AL_Gharbi_Zone_5_Edit()

Dim wb As Workbook
Dim ws As Worksheet
Dim path As String
Dim lastRow As Long
Dim wblist As Variant

wblist = ThisWorkbook.Worksheets("Data").Range("B4:B20")
path = ThisWorkbook.Worksheets("Data").Cells(3, 2)

If IsInArray("AL GHARBI ZONE 5(NEW).xlsb", wblist) Then
            
        Set wb = Workbooks.Open(path & "\AL GHARBI ZONE 5(NEW).xlsb")
        Set ws = wb.Worksheets("sheet1")
        lastRow = ws.Cells(Rows.Count, 3).End(xlUp).Row
        
        Cells.FormatConditions.Delete
        
         ws.Range("G2:G" & lastRow) = "AL GHARBI ZONE-5"
         ws.Range("H2:H" & lastRow) = "ZONE-5 CRUSHER"
         ws.Range("I2:I" & lastRow) = "MASAR"
         
         ws.Range("K2:K" & lastRow) = ""
         ws.Range("F2:F" & lastRow) = "0-100 MM"
         ws.Range("L2:L" & lastRow) = 29000
         wb.Close (True)
End If


End Sub
