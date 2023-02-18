Attribute VB_Name = "Our_EQ_List"


Sub MBabarList()


Dim wss_EW, wsPanel_EW As Worksheet
Dim wsd_EW As Worksheet
Dim wbs_EW, wbd_EW, wbPanel_EW As Workbook
Dim i As Integer
Dim j As Long
Dim ldr As Long
Dim lsr As Integer
Dim wb As Workbook
Dim counter As Integer
Dim copyrnge As Range
Dim sLastRow As Range

Application.IgnoreRemoteRequests = True
Application.ScreenUpdating = False
Application.DisplayAlerts = False


Set wbPanel_EW = ThisWorkbook
Set wsPanel_EW = wbPanel_EW.Worksheets("Our EQ Timesheets")

counter = 0

'**************************************IS OPEN CONTROL***********************************************

For Each wb In Workbooks

    If wb.Name = wsPanel_EW.Cells(2, 2).Value Then
    
        Set wbs_EW = wb
        counter = counter + 1
    
    ElseIf wb.Name = wsPanel_EW.Cells(4, 2).Value Then
    
        Set wbd_EW = wb
        counter = counter + 10
        
    End If
Next

If counter = 1 Then

    Set wbd_EW = Workbooks.Open(wsPanel_EW.Cells(3, 2) & "\" & wsPanel_EW.Cells(4, 2), True)
    
ElseIf counter = 10 Then
    
    Set wbs_EW = Workbooks.Open(wsPanel_EW.Cells(1, 2) & "\" & wsPanel_EW.Cells(2, 2), True)

ElseIf counter = 11 Then
    
    
Else
    Set wbs_EW = Workbooks.Open(wsPanel_EW.Cells(1, 2) & "\" & wsPanel_EW.Cells(2, 2), True)
    Set wbd_EW = Workbooks.Open(wsPanel_EW.Cells(3, 2) & "\" & wsPanel_EW.Cells(4, 2), True)
    
End If

'**************************************COPY DATA******************************************


Set wsd_EW = wbd_EW.Worksheets("Makina_Saat")


    For Each wss_EW In wbs_EW.Worksheets

            ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row + 1
            
            lsr = wss_EW.Cells(Rows.Count, 8).End(xlUp).Row
            wss_EW.Activate
            Set sLastRow = wss_EW.Range("H" & lsr)
            
            sLastRow.Select
            
            If Selection = "" And Selection.Row > 2 Then lsr = Selection.End(xlUp).Row
  
            'lsr = wss_EW.Cells(Rows.Count, 8).End(xlUp).Row
            
            
            If lsr < 2 Then lsr = 2
            
            wss_EW.Range("B2:M" & lsr).Copy
            wsd_EW.Range("A" & ldr).PasteSpecial xlPasteValuesAndNumberFormats
            
       
    Next
    
    ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row
    'wsd_EW.ListObjects("Table_1").Resize wsd_EW.Range("A1:Q" & ldr)

   Call copyFormatAndFormulas(wsd_EW, ldr)

    wsd_EW.Range("Q2:Q" & ldr).Copy
    wsd_EW.Range("G2:G" & ldr).PasteSpecial xlPasteValues
    
    
 '**********************************FINAL CORRECTIONS******************************************
    
    
    Dim dupRange As Range

    wsd_EW.Activate
    
    
    Call ZoneDuzelt(wsd_EW, ldr)
    
    wsd_EW.Range("G2:G" & ldr).TextToColumns
    Call correctionDictionary(wsd_EW, wbd_EW.Worksheets("Correction Dictionary"), ldr)
    
    wsd_EW.Range("A2:A" & ldr).NumberFormat = "dd.mm.yy"
    
    ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row
    
    For j = 2 To ldr
    
        'wsd_EW.Cells(j, 7) = Trim(wsd_EW.Cells(j, 7).Value)
    
        If wsd_EW.Cells(j, 7) = "" Then
            
            On Error Resume Next
            wsd_EW.Rows(j).Delete
            
            
            j = j - 1
            
            If j = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row Then Exit For
            
        End If
        
    Next j
    
    
    ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row
    
    Set dupRange = wsd_EW.Range("A1:Q" & ldr)
    dupRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17), Header:=xlYes
    
    ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row
    
    
    
    Call makina_duzelt(wsd_EW, wbd_EW.Worksheets("Makina_List"), ldr)
    
    Call copyFormatAndFormulas(wsd_EW, ldr)
    Set dupRange = wsd_EW.Range("A1:Q" & ldr)
    wsd_EW.Range("G2:G" & ldr).TextToColumns
    dupRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17), Header:=xlYes
    
    ldr = wsd_EW.Cells(Rows.Count, 7).End(xlUp).Row
    
    Call copyFormatAndFormulas(wsd_EW, ldr)
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.IgnoreRemoteRequests = False

    wbd_EW.Close (True)
    wbs_EW.Close



End Sub

Sub ZoneDuzelt(ws As Worksheet, lr As Long)

Dim arr As Variant
Dim i As Long
Dim zone As String


arr = ws.Range("B2:B" & lr).Value

For i = LBound(arr) To UBound(arr)

    zone = arr(i, 1)
    
    Select Case zone
    
    Case 0
    
        arr(i, 1) = "Zone-0"
    
    Case 1
    
        arr(i, 1) = "Zone-1"
    
    Case 2
    
        arr(i, 1) = "Zone-2"
    
    Case 3
    
        arr(i, 1) = "Zone-3"
    
    Case 4
    
        arr(i, 1) = "Zone-4"
    
    Case 5
    
        arr(i, 1) = "Zone-5"
    
    Case 6
    
        arr(i, 1) = "Zone-6"
    
    Case 7
    
        arr(i, 1) = "Zone-7"
    
    Case 5
    
        arr(i, 1) = "Zone-5"
    
    Case "5 C"
    
        arr(i, 1) = "Zone-5C"
    
    Case "C 5"
    
        arr(i, 1) = "Zone-5C"
    
    Case "S C"
    
        arr(i, 1) = "Zone-5C"
        
    Case "5 C "
    
        arr(i, 1) = "Zone-5C"
        
    Case "5C"
    
        arr(i, 1) = "Zone-5C"
        
    Case " 5C"
    
        arr(i, 1) = "Zone-5C"
        
    Case "5C "
    
        arr(i, 1) = "Zone-5C"
        
        
    Case ""
    
        
        arr(i, 1) = ws.Range("B2:B" & Int((100 * Rnd) + 2)).Value
        
    Case "-"
    
        
        arr(i, 1) = ws.Range("B2:B" & Int((100 * Rnd) + 2)).Value
    

    
    End Select
    
Next i
    

ws.Range("B2:B" & lr) = arr



End Sub

Sub makina_duzelt(ws As Worksheet, ws2 As Worksheet, lr As Long)

    Dim arr As Variant
    Dim arr2 As Variant
    Dim dict As Dictionary
    
    Set dict = New Dictionary
    
    arr = ws.Range("G2:G" & lr).Value
    lr2 = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        '/////equipment types/////
        
        For i = 2 To lr2
            
            dict(ws2.Cells(i, 1).Value) = ws2.Cells(i, 3).Value
            
        Next i
        
        
        For i = 1 To lr
            On Error Resume Next
            arr(i, 1) = dict.Item(arr(i, 1))
            
        Next i
        
        ws.Range("E2:E" & lr).Value = arr
        
        
End Sub

Sub correctionDictionary(ws As Worksheet, wsdict As Worksheet, lr As Long)

    Dim lrdict As Long
    Dim arr As Variant
    Dim dict As New Dictionary
    lrdict = wsdict.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lrdict
    
        dict(wsdict.Cells(i, 1).Value) = wsdict.Cells(i, 2).Value
        
    Next i
    
    arr = ws.Range("G60000:G" & lr).Value
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        If dict.Exists(arr(i, 1)) Then
            arr(i, 1) = dict.Item(arr(i, 1))
        
        End If
        
    Next i
    
    ws.Range("G60000:G" & lr).Value = arr
    
End Sub

Sub renkss()
'ActiveSheet.UsedRange 'Refresh UsedRange
  'LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
MsgBox ActiveSheet.Tab.Color
'MsgBox ActiveSheet.Range("A1").CurrentRegion.Rows.Count
MsgBox ActiveSheet.Cells(2, 8).End(xlDown).Row

End Sub

Sub copyFormatAndFormulas(wsd_EW As Worksheet, ldr As Long)

    '///////////////////FORMULAS///////////////////
    
    
    wsd_EW.Range("D2").Copy
    
    wsd_EW.Range("D3:D" & ldr).PasteSpecial xlPasteFormulas
    
    wsd_EW.Range("N2").Copy
    
    wsd_EW.Range("N3:N" & ldr).PasteSpecial xlPasteFormulas
    
    wsd_EW.Range("O2").Copy
    
    wsd_EW.Range("O3:O" & ldr).PasteSpecial xlPasteFormulas
    
    wsd_EW.Range("P2").Copy
    
    wsd_EW.Range("P3:P" & ldr).PasteSpecial xlPasteFormulas
    
    wsd_EW.Range("Q2").Copy
    
    wsd_EW.Range("Q3:Q" & ldr).PasteSpecial xlPasteFormulas
    
    '/////////////////////FORMATS////////////////////////
    
    wsd_EW.Range("A2:Q2").Copy
    
    wsd_EW.Range("A3:Q" & ldr).PasteSpecial xlPasteFormats
    
    
End Sub
