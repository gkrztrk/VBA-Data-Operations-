Attribute VB_Name = "Auto_Trigger_Macro"


Sub open_n_call_schl()

Dim wb As Workbook

Set wb = Workbooks.Open("J:\My Drive\Gkr\Reports\Roads & Paving from Agreed Dates Schedule.xlsm")

Application.Run "'Roads & Paving from Agreed Dates Schedule.xlsm'!Update_and_send_mail"

wb.Close True

End Sub
