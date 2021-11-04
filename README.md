# jira-excell-integration

Fastest excell - jiraAPI integration with macros or import bas file on VBA

**Requirements**
1. Open MS Excell 
2. Alt+f11
3. Tools -> References -> check "Microsoft Scripting Runtime" -> ok
4. Click on left pad by "Modules" and import funcJira.bas
5. Download json parse module "JsonConverter" on "https://github.com/VBA-tools/VBA-JSON"

**funcJira.bas:**
    
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


Function on excell cell: "=getTaskStatus("https://domain-jira.ru/rest/api/2/issue/" & A2 & ".json")"

**Macro event by onclick button:**.
   
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
      MsgBox "Готово!"
    End Sub
      
