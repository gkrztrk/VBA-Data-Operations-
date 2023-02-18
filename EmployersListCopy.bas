Attribute VB_Name = "EmployersListCopy"
Sub OpenAllWorkbooks()
'Step 1:Declare your variables
    Dim MyFiles As String
    Dim wss As Worksheet
    Dim wsd As Worksheet
    Dim wbs As Workbook
    Dim wbd As Workbook
    Dim bulunan As Range
    Dim lr As Long
'Step 2: Specify a target folder/directory, you may change it.
    Dim sPath As String
    Dim dPath As String
    
    
    Application.ScreenUpdating = False
Application.DisplayAlerts = False
    
    sPath = sh_Employers.Cells(3, 2).Value
    MyFiles = Dir(sPath & "\*.xlsx")
    dPath = sh_Employers.Cells(1, 2).Value & "\" & sh_Employers.Cells(2, 2).Value
    Set wbd = Workbooks.Open(dPath)
    Set wsd = wbd.Worksheets("PERSONNEL LIST 21")
    
    Do While MyFiles <> ""
'Step 3: Open Workbooks one by one
        Set wbs = Workbooks.Open(sPath & "\" & MyFiles)
        Set wss = wbs.Worksheets(5)
        
        On Error Resume Next
        wss.AutoFilter.ShowAllData
        
        lr = wsd.Cells(Rows.Count, 1).End(xlUp).Row + 1
        lrs = wss.Cells(Rows.Count, 1).End(xlUp).Row
        
        Set bulunan = ThisWorkbook.Worksheets("DLmail").Range("A:A").Find(wbs.Name)
        
        If Not bulunan Is Nothing Then
        
            r = bulunan.Row
            
            ThisWorkbook.Worksheets("DLmail").Cells(r, 2).Copy
            
            wss.Range("Q2:Q" & lrs).PasteSpecial xlPasteValues
            
            
        End If
        
        
        wss.Range("A2:Q" & lrs).Copy
        
        wsd.Range("A" & lr).PasteSpecial xlPasteValues

    
        wbs.Close
    
    'Step 4: Next File in the folder/Directory
        MyFiles = Dir
    Loop
    
    
    
    wbd.Close (True)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub


Sub addWorkingHours()

Dim wbd As Workbook
Dim wsd As Worksheet
Dim arr As Variant
Dim arr2 As Variant


dPath = sh_Employers.Cells(1, 2).Value & "\" & sh_Employers.Cells(2, 2).Value
    Set wbd = Workbooks.Open(dPath)
    Set wsd = wbd.Worksheets("PERSONNEL LIST 21")
    
    lr = wsd.Cells(Rows.Count, 1).End(xlUp).Row
    
   arr = wsd.Range("O2:O" & lr).Value
   
   For i = LBound(arr, 1) To UBound(arr, 1)
   
        If arr(i, 1) = "PRESENT" Or arr(i, 1) = "PRESENT-E" Or arr(i, 1) = "Present" Then
        
            arr(i, 1) = 10
        Else
        
            arr(i, 1) = 0
        End If
        
    Next i
    
   
    wsd.Range("R2:R" & lr).Value = arr
    


End Sub
