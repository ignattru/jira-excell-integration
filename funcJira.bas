Function getTaskStatus(ByVal taskURL As String) As String
  ' Basic authorization
  Dim jiraRequest As Object
  Set jiraRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
  jiraRequest.Open "POST", "https://domain-jira.ru/login.jsp?os_username=jiralogin&os_password=jirapass"
  jiraRequest.Send

  ' GET request in rest API Jira
  jiraRequest.Open "GET", taskURL
  jiraRequest.Send

  ' Parse Json response
  Dim Json As Object
  Set Json = JsonConverter.ParseJson(jiraRequest.ResponseText)

  ' return required attribute
  getTaskStatus = Json("fields")("status")("name")
End Function