Attribute VB_Name = "DownloadAttachments"
Sub dlAttachments()

Dim olApp As Outlook.Application
Dim objNS As Outlook.Namespace
Set olApp = Outlook.Application
Dim olFolder As Outlook.MAPIFolder
Set objNS = olApp.GetNamespace("MAPI")
Dim Msg As Outlook.MailItem
Dim att As Outlook.Attachment
Dim attlist As Scripting.Dictionary

    Set attlist = New Dictionary

  ' default local Inbox
  Set olFolder = objNS.GetDefaultFolder(olFolderInbox).Folders("Timekeeper")
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Folders("Timekeeper").Items

 If TypeName(Items) = "MailItem" Then
    Set Msg = Items
 End If

a = 2
    For Each Item In olFolder.Items
        
        If TypeName(Item) = "MailItem" Then
            
            Set Msg = Item

       
        
        For Each att In Msg.Attachments
        
        
            If InStr(att.FileName, "Zone Wise") > 0 Then
                oPath = "J:\My Drive\Gkr\Data Source\employers\" & att.FileName
                att.SaveAsFile oPath
                
                ThisWorkbook.Worksheets("DLmail").Cells(a, 1) = att.FileName
                ThisWorkbook.Worksheets("DLmail").Cells(a, 2) = Msg.ReceivedTime - 1
                
                a = a + 1
                
                
            End If
        Next
        
        End If
    Next

End Sub

