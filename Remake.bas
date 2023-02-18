Attribute VB_Name = "Remake"
Sub rmk_amnt()

Dim dest_row As Long
Dim ws As Worksheet
Dim i As Long

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Set wb_panel = ActiveWorkbook

Dest_Path = wb_panel.Worksheets("Data").Cells(1, 2).Value

         
    Set wbd = Workbooks.Open(FileName:=Dest_Path & "\" & wb_panel.Worksheets("Data").Cells(2, 2), UpdateLinks:=0)
        
        
Set ws = wbd.Worksheets("All")
        
lr = ws.Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To lr
    
        Call amount(i, wbd)
        
    Next i

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


Sub rmk_m3()

Dim dest_row As Long
Dim ws As Worksheet
Dim i As Long

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set wb_panel = ActiveWorkbook

Dest_Path = wb_panel.Worksheets("Data").Cells(1, 2).Value

         
    Set wbd = Workbooks.Open(FileName:=Dest_Path & "\" & wb_panel.Worksheets("Data").Cells(2, 2), UpdateLinks:=0)
        
        
Set ws = wbd.Worksheets("All")
        
lr = ws.Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To lr
    
        Call m3_hesap(i, wbd)
        
    Next i
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
