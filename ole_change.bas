Attribute VB_Name = "ole_change"
Sub olechange()


If Application.IgnoreRemoteRequests = True Then

Application.IgnoreRemoteRequests = False

MsgBox "OLE KAPANDI"

Else

Application.IgnoreRemoteRequests = True

MsgBox "OLE ACILDI"

End If

ThisWorkbook.Save
Application.Workbooks.Open (ThisWorkbook.FullName)

End Sub
