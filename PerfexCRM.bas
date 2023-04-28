Attribute VB_Name = "PerfexCRM"
Public Const perfexcrm_url As String = "https://myperfexcrmurl.com"
Public Const authtoken As String = "AbcABC..."
Public Const default_priority As Integer = 2
Public Const default_departmentID As Integer = 5


Sub PerfexCRM_OpenTicket()
    Dim priority As Integer
    Dim otlMailItem As MailItem
    Set otlMailItem = ActiveExplorer.Selection.Item(1)

    priority = default_priority

    With frmPerfexCRM
        .txtSubject.Value = Trim(otlMailItem.subject)
        .txtName.Value = Trim(otlMailItem.Sender)
        .txtEmail.Value = otlMailItem.SenderEmailAddress
        .txtPriority.Value = priority
        .txtCC.Value = Trim(otlMailItem.cc)
        .txtMessage.Value = Trim(otlMailItem.Body)
    End With
    
    frmPerfexCRM.Show
End Sub

Function PerfexCRM_getContactInfo(email As String) As Variant
    Dim response As String
    Dim r(2) As Integer
    On Error GoTo errLines

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", perfexcrm_url & "/api/contacts/search/" & email, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "authtoken", authtoken
    objHTTP.Send ""
    response = Trim(StrConv(objHTTP.responseBody, vbUnicode))
    response = Right(response, Len(response) - 1)
    response = Left(response, Len(response) - 1)
    
    'https://github.com/VBA-tools/VBA-JSON
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(response)
    r(0) = Json("id")
    r(1) = Json("userid")
    PerfexCRM_getContactInfo = r
    
    Exit Function
    
errLines:
    MsgBox ("Cannot get contact ID. Response:" & response)
    PerfexCRM_getContactInfo = 0
End Function

Function PerfexCRM_OpenTicketPost(subject As String, name As String, email As String, priority As Integer, message As String, cc As String)
    Dim data As String
    Dim contactinfo As Variant
    Dim r As String
    
    contactinfo = PerfexCRM_getContactInfo(email)
    If IsNumeric(contactinfo) Then Exit Function
    
    data = "subject=" & subject
    data = data & "&department=" & CStr(default_departmentID)
    data = data & "&contactid=" & CStr(contactinfo(0))
    data = data & "&userid=" & CStr(contactinfo(1))
    data = data & "&email=" & email
    data = data & "&priority=" & CStr(priority)
    data = data & "&message=" & message
    
    On Error GoTo errLines

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "POST", perfexcrm_url & "/api/tickets", False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objHTTP.setRequestHeader "authtoken", authtoken
    
    objHTTP.Send data
    r = StrConv(objHTTP.responseBody, vbUnicode)
    MsgBox r
    PerfexCRM_OpenTicketPost = True
    Exit Function
    
errLines:
    MsgBox ("Cannot post the data. Check your connection or the PerfexCRM form URL. Response:" & r)
    PerfexCRM_OpenTicketPost = False
End Function
