Sub getJSON()

  Dim taskId As String
  Dim jiraRequest As Object
  Dim Json As Object
  Set jiraRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
  jiraRequest.Open "POST", "https://domain-jira.ru/login.jsp?os_username=yourlogin&os_password=yourpass"
  jiraRequest.Send

  Dim i As Long
  For i = 3 To Application.WorksheetFunction.CountA(Range("A3:A300")) + 2
  Cells(i, 3) = i
    taskId = Range("A" & i)
    jiraRequest.Open "GET", "https://domain-jira.ru/rest/api/2/issue/" & taskId
    jiraRequest.Send
    
    Set Json = JsonConverter.ParseJson(jiraRequest.ResponseText)
    Range("B" & i) = Json("fields")("status")("name")
    Range("C" & i) = Json("fields")("priority")("name")
    Range("D" & i) = Json("fields")("created")
    Range("E" & i) = Json("fields")("summary")
  Next
  MsgBox "Done!"
End Sub