Attribute VB_Name = "logCode"
Sub SendToGoogle()
 
'This Macro Requires Reference to "Microsoft XML, v6.0" (VBA Editor > Tools > References, find &amp; select from list)
 
Dim URL_First As String        'Assign the first part of URL to send the data
Dim URL_Last As String         'Assign the last part of URL where we will update the information
Dim Form_URL As String         'To store the Form URL after merging Beginning and End URL
 
Dim HeaderName As String       'Variable to store the header type i.e. Content-Type
Dim SendID As String           'To store the information required to send a particular information to Google form
 
'Variables to store user inputs from Excel UserForm
Dim UsrName As String
Dim CompName As String
Dim FlName As String

 
'Assign User inputs to variables
 
UsrName = Environ("username")
CompName = Environ("computername")
FlName = ThisWorkbook.Name
 
'Variable to store what we need to send to server
 
Dim TicketInfo As MSXML2.ServerXMLHTTP60 'XML variable to send the information to server
 
'Content-Type is actually a header type which tells the client what the content type of the returned content actually is. Google recognizes this header type
 
HeaderName = "Content-Type"
 
'SendID  required to send a particular information to Google Form
SendID = "application/x-www-form-urlencoded; charset=utf-8"
 
'In actual link, we need to replace viewform? with formResponse?ifq&amp;
'need to find the “name” attributes for the text boxes and the value for them
'add at the end &amp;submit=Submit and use it, it must post all the data you specified in one step.
 
'formRespose is used to get the response from Google Form after submitting the details
'Submit - it is a command to submit the filled form
 
URL_First = "https://docs.google.com/forms/d/e/.........."
 
URL_Last = "&entry.2006625714=" & UsrName & "&entry.812953752=" & CompName & "&entry.1687321223=" & FlName & "&submit=Submit"
 
'Creating the Final URL
Form_URL = URL_First & URL_Last
 
Set TicketInfo = New ServerXMLHTTP60 'Setting the reference of new server xmlhttp 60
 
TicketInfo.Open "POST", Form_URL, False ' Posting the entire link
 
TicketInfo.setRequestHeader HeaderName, SendID 'Specifies the name of an HTTP header.
 
TicketInfo.Send 'Send all the information over google
 
'StatusText is provide the status of data submission. It will show OK if data will be successfully submitted
 
If TicketInfo.statusText = "OK" Then 'Check for successful send
 
  'Call Reset 'Call Reset procedure to reset form Excel Form after submitting the data
  'MsgBox "Thank you for submitting data!"
 
Else
  MsgBox "Please check your internet connection &amp; required details"
End If
 
End Sub
