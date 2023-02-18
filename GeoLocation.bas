Attribute VB_Name = "GeoLocation"
Sub test()

Dim IP As String
Dim IPcity As String
Dim IPcountry As String

Dim http As Object
Dim xmlDoc As MSXML2.DOMDocument60
Dim strURL As String

' requires reference to Microsoft XML 6.0

    IP = GetIPAddress
    strURL = "https://ipapi.co/" & IP & "/xml/"
        
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", strURL, False
    
    http.Send
    Set xmlDoc = New MSXML2.DOMDocument60
    
    xmlDoc.LoadXML http.responseText
    
    'Debug.Print http.responseText
    IPcity = xmlDoc.SelectSingleNode("//root/city").Text
    IPcountry = xmlDoc.SelectSingleNode("//root/country").Text
    
    Debug.Print IPcity
    Debug.Print IPcountry


End Sub



Function GetIPAddress()
    Const strComputer As String = "."   ' Computer name. Dot means local computer
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' Connect to the WMI service
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress
        End If
    Next

    GetIPAddress = strIPAddress
End Function
