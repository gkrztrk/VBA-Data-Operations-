Attribute VB_Name = "Unpivot_Nesma_Timesheet"
Option Explicit


Public hata_row As Integer

Sub LoopAllExcelFilesInFolder_Nesma()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim wb_data As Workbook
Dim ws_data As Worksheet
Dim wb_data_path As String
Dim wb_data_file As String
Dim ws_data_name As String


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False


wb_data_path = ThisWorkbook.Worksheets("UnPivot PMV EQ Timesheets").Range("B3").Value
wb_data_file = ThisWorkbook.Worksheets("UnPivot PMV EQ Timesheets").Range("B4").Value
ws_data_name = ThisWorkbook.Worksheets("UnPivot PMV EQ Timesheets").Range("B5").Value
Set wb_data = Workbooks.Open(wb_data_path & "\" & wb_data_file)
Set ws_data = wb_data.Worksheets(ws_data_name)

hata_row = 2


'Optimize Macro Speed
  

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
    On Error Resume Next
      Set wb = Workbooks.Open(FileName:=myPath & myFile, UpdateLinks:=0)
        wb.Activate
        
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Change First Worksheet's Background Fill Blue
      Call UnpivotWorkbook_Nesma(wb_data, ws_data)
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"
    
ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


Sub UnpivotWorkbook_Nesma(wbData As Workbook, wsData As Worksheet)

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "DATA" Then
        
            ws.Activate
        
            'If ws.Tab.ColorIndex <> 4 And ws.Visible Then
                Call UnpivotWorksheets_Nesma(wbData, wsData)
                
            'End If
        End If
    
    Next ws
    'MsgBox "ISLEM TAMAMLANDI"
End Sub
Sub UnpivotWorksheets_Nesma(wbData As Workbook, ws As Worksheet)
Static sn As Integer
Dim bulunan As Range
Dim t As Integer
'Dim ws As Worksheet
Dim ss As Worksheet
sn = 0
Dim lr, lc, dr, dc, date_row, date_column, i, j As Integer

'COLUMNSSSSS
Dim clnfind As Range
Dim eq_type, shift, sr_no, plt_no, iqm, prjct_code, zone, drvr_name As Integer




Set bulunan = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("TAG NO.")
Set ss = ActiveSheet

If bulunan Is Nothing Then

    'Set bulunan = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("PLATE")
        'Range(Cells(1, bulunan.Column + 1), Cells(1, bulunan.Column + 1)).EntireColumn.Insert
        'Cells(bulunan.Row, bulunan.Column + 1).Value = "SERIAL NO."
    MsgBox "SERIAL NO. Column couldnt find in " & ActiveWorkbook.Name
    ThisWorkbook.Worksheets("UnPivot PMV EQ Timesheets").Cells(hata_row, 3) = ActiveWorkbook.Name
    hata_row = hata_row + 1

End If
Set bulunan = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("TAG NO.")
lr = bulunan.Row + 2
lc = bulunan.Column
date_row = lr - 2
date_column = lc + 3


'MsgBox lr & "   " & lc

Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("TAG NO.")

If Not clnfind Is Nothing Then

    sr_no = clnfind.Column
    
End If

Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("DESCRIPTION")

If Not clnfind Is Nothing Then

    eq_type = clnfind.Column
    
End If

'Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("IQAMA")
'
'If Not clnfind Is Nothing Then
'
'    iqm = clnfind.Column
'
'End If



    prjct_code = sr_no + 1
    


    zone = sr_no + 2
    


'Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("DRIVER")
'
'If Not clnfind Is Nothing Then
'
'    drvr_name = clnfind.Column
'
'End If
'

'Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("PLATE")
'
'If Not clnfind Is Nothing Then
'
'    plt_no = clnfind.Column
'
'End If
    
'Set clnfind = ActiveSheet.Range(Cells(1, 1), Cells(50, 50)).Find("SHIFT")
'
'If Not clnfind Is Nothing Then
'
'    shift = clnfind.Column
'
'    date_column = shift + 1
'
'
'End If
t = 21

For j = date_column To (ss.Cells(date_row, Columns.Count).End(xlToLeft).Column) - 5
    date_column = j
    lr = bulunan.Row + 2
    
    For i = lr To ss.Cells(Rows.Count, 1).End(xlUp).Row
    
        lr = i
        
        dr = ws.Cells(Rows.Count, 2).End(xlUp).Row + 1
        
        
        'ws.Cells(dr, 1) = dr - 1
        ws.Cells(dr, 2) = t + ThisWorkbook.Worksheets("UnPivot PMV EQ Timesheets").Range("B6") - 1
        
        If eq_type <> 0 Then
        
            If ss.Cells(lr, eq_type).MergeCells = True Then
            
               ss.Cells(lr, eq_type).UnMerge
               ss.Cells(lr, eq_type).Copy ss.Cells(lr + 1, eq_type)
               
            End If
        
                
        ws.Cells(dr, 3) = ss.Cells(lr, eq_type).Value
        End If
        
        '------------------------------------------------------
        If plt_no <> 0 Then
        
            If ss.Cells(lr, plt_no).MergeCells = True Then
            
               ss.Cells(lr, plt_no).UnMerge
               ss.Cells(lr, plt_no).Copy ss.Cells(lr + 1, plt_no)
               
            End If
        
        
        ws.Cells(dr, 4) = ss.Cells(lr, plt_no).Value
       End If
       
        '------------------------------------------------------
        
        
            If ss.Cells(lr, sr_no).MergeCells = True Then
            
               ss.Cells(lr, sr_no).UnMerge
               ss.Cells(lr, sr_no).Copy ss.Cells(lr + 1, sr_no)
               
            End If
        
        
        ws.Cells(dr, 5) = ss.Cells(lr, sr_no).Value
        
        '-------------------------------------------------------
        If drvr_name <> 0 Then
        
        ws.Cells(dr, 6) = ss.Cells(lr, drvr_name).Value
        
        End If
        
        '-------------------------------------------------------
        If iqm <> 0 Then
        
        ws.Cells(dr, 7) = ss.Cells(lr, iqm).Value
        
        End If
        
        '-------------------------------------------------------
        If prjct_code <> 0 Then
        
        ws.Cells(dr, 8) = ss.Cells(lr, prjct_code).Value
        
        End If
        
        '-------------------------------------------------------
        If zone <> 0 Then
        
        ws.Cells(dr, 9) = ss.Cells(lr, zone).Value
        
        End If
        
        '-------------------------------------------------------
        If shift <> 0 Then
        
        ws.Cells(dr, 10) = ss.Cells(lr, shift).Value
        
        End If
        
        '-------------------------------------------------------
        
        ws.Cells(dr, 11) = ss.Cells(lr, date_column).Value
        
        '-------------------------------------------------------
        
        
        ws.Cells(dr, 12) = "NESMA"
        
        ws.Cells(dr, 13) = ss.Cells(3, 1).Value
        
        
        

    Next i
        
   t = t + 1
    
Next j



    ss.Tab.ColorIndex = 4
    




End Sub



