Attribute VB_Name = "AmntAndM3"
Option Explicit


Sub amount(a As Long, d As Workbook)
Dim t As Integer
Dim dt As Date

dt = d.Worksheets("All").Cells(a, 2)
t = Month(dt)

    If d.Worksheets("All").Cells(a, 15).Value = "Nesma" Then
        
        If Year(dt) = 2021 Then
        
            d.Worksheets("All").Cells(a, 16) = Int((d.Worksheets("Crusher Rates").Cells(t, 2).Value) * d.Worksheets("All").Cells(a, 12) / 1000)
            
        Else
            
            d.Worksheets("All").Cells(a, 16) = Int((d.Worksheets("Crusher Rates").Cells(t, 5).Value) * d.Worksheets("All").Cells(a, 12) / 1000)
        
        End If
        
        d.Worksheets("All").Cells(a, 17) = Int(8.75 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    ElseIf d.Worksheets("All").Cells(a, 8).Value = "AL GHARBI FAYHA" Then
        
        d.Worksheets("All").Cells(a, 18) = Int(7.5 * d.Worksheets("All").Cells(a, 12) / 1000)
    
        'd.Worksheets("All").Cells(a, 16) = Int(7.5 * d.Worksheets("All").Cells(a, 12) / 1000) crusher cost yok
        
        d.Worksheets("All").Cells(a, 17) = Int(8.75 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
   ElseIf d.Worksheets("All").Cells(a, 8).Value = "ZONE-5 CRUSHER" Then
        
        d.Worksheets("All").Cells(a, 18) = Int(7.5 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        'd.Worksheets("All").Cells(a, 16) = Int(7.5 * d.Worksheets("All").Cells(a, 12) / 1000) crusher cost yok
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "SAND" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(12 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "DRY SAND" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(18 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
     ElseIf d.Worksheets("All").Cells(a, 6).Value = "0-5 MM" And d.Worksheets("All").Cells(a, 15).Value <> "Nesma" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(18 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
     ElseIf d.Worksheets("All").Cells(a, 6).Value = "0-50 MM" And d.Worksheets("All").Cells(a, 15).Value <> "Nesma" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(22 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
    
      ElseIf d.Worksheets("All").Cells(a, 6).Value = "10-40 MM" And d.Worksheets("All").Cells(a, 15).Value <> "Nesma" And d.Worksheets("All").Cells(a, 7).Value = "AL-JUSOOR" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(26 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "10-40 MM" And d.Worksheets("All").Cells(a, 15).Value <> "Nesma" And d.Worksheets("All").Cells(a, 7).Value <> "AL-JUSOOR" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(24 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "3/8 AGG" And d.Worksheets("All").Cells(a, 15).Value <> "Nesma" Then
    
        d.Worksheets("All").Cells(a, 18) = Int(26 * d.Worksheets("All").Cells(a, 12) / 1000)
        
        d.Worksheets("All").Cells(a, 19) = d.Worksheets("All").Cells(a, 16).Value + d.Worksheets("All").Cells(a, 17).Value + d.Worksheets("All").Cells(a, 18).Value
        
    Else
    
        d.Worksheets("All").Cells(a, 17) = "Err"
        
    End If


End Sub

Sub m3_hesap(a As Long, d As Workbook)



    If d.Worksheets("All").Cells(a, 6).Value = "0-100 MM" Then
    
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 2.2) / 1000
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "0-50 MM" Or d.Worksheets("All").Cells(a, 6).Value = "0-40 MM" Or d.Worksheets("All").Cells(a, 6).Value = "0-70 MM" Then
        
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 2.2) / 1000
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "0-5 MM" Then
        
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 1.7) / 1000
    
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "10-40 MM" Then
        
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 1.8) / 1000
    
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "SAND" Or d.Worksheets("All").Cells(a, 6).Value = "DRY SAND" Then
        
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 1.7) / 1000
        
    ElseIf d.Worksheets("All").Cells(a, 6).Value = "3/8 AGG" Then
        
        d.Worksheets("All").Cells(a, 14) = Int(d.Worksheets("All").Cells(a, 12).Value / 1.6) / 1000
        
    Else
    
        d.Worksheets("All").Cells(a, 14) = 0
    
    End If
    


End Sub

