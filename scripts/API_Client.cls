VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "API_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mURL As String
Public mAPI_key As String
Public mJobRoute As String
Public mResultRoute As String

Public Sub Initialize(url As String, key As String)
    Me.mJobRoute = "api/jobs"
    Me.mResultRoute = "api/results"
    Me.mURL = url
    Me.mAPI_key = key
End Sub

Public Function call_api_home_credit_model(data As String, resultsSheet As String) As Boolean
    ' setup the destination and response URLs
    Dim DestURL As String
    DestURL = Me.mURL & Me.mJobRoute
    Dim ResURL As String
    ResURL = Me.mURL & Me.mResultRoute
    
    ' create json body of nested dictionaries
    ' inner most dict first
    Dim data_string As Dictionary
    Set data_string = New Dictionary
    data_string.Add "input", data
    
    ' source name (excel-data)
    Dim source_name As Dictionary
    Set source_name = New Dictionary
    source_name.Add "excel-data", data_string
    
    ' entire input contents
    Dim input_contents As Dictionary
    Set input_contents = New Dictionary
    input_contents.Add "type", "text"
    input_contents.Add "sources", source_name
    
    
    ' model contents
    Dim model_contents As Dictionary
    Set model_contents = New Dictionary
    model_contents.Add "identifier", "jt3jyqtufg"
    model_contents.Add "version", "1.0.0"
    
    ' combin model and input into a single json
    Dim final_dict As Dictionary
    Set final_dict = New Dictionary
    final_dict.Add "model", model_contents
    final_dict.Add "input", input_contents

    body = JsonConverter.ConvertToJson(ByVal final_dict)
    
    'submit the request
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "POST", DestURL, False
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "Accept", "application/json"
    MyRequest.setRequestHeader "Authorization", "ApiKey " & Me.mAPI_key
    
    MyRequest.Send body
    
    'get the current job  ID
    Dim jobResponseJSON As Object
    Set jobResponseJSON = JsonConverter.ParseJson(MyRequest.ResponseText)
    
    jobID = jobResponseJSON("jobIdentifier")

    ' get the status of the Job
    Dim statusURL As String
    statusURL = DestURL & "/" & jobID
    
    ' wait until job is finished
     Do While check_status(statusURL)
        waitTill = Now() + TimeValue("00:00:05")

        While Now() < waitTill
            DoEvents
        Wend
    Loop
    
    'get result
    result_raw = get_result(ResURL & "/" & jobID)
    
    Dim resultSheet As Worksheet
    Set resultSheet = Sheets(resultsSheet)
    
    Dim resultResponseJSON As Object
    Set resultResponseJSON = JsonConverter.ParseJson(UTF8_Decode(result_raw))
    
    Dim key, pred
    Dim c As Range
    Dim rowCounter As Integer
    rowCounter = 1

    Set c = resultSheet.Range("A7")

    'loop over timesheets
    For Each key In resultResponseJSON("results")("excel-data")("results.json")("data")("result")("predictions")
        c.Cells(rowCounter, 1).Value = key("row")
        c.Cells(rowCounter, 2).Value = key("pred")
        rowCounter = rowCounter + 1
    Next key
    
End Function

Public Function UTF8_Decode(ByVal sStr As String)
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            iChar3 = Asc(Mid(sStr, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function


Public Function check_status(url As String) As Boolean

    Set MyStatus = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyStatus.Open "GET", url, False
    MyStatus.setRequestHeader "Content-type", "application/json"
    MyStatus.setRequestHeader "Accept", "application/json"
    MyStatus.setRequestHeader "Authorization", "ApiKey " & Me.mAPI_key
    
    MyStatus.Send
    response2 = MyStatus.ResponseText
    
    statusLocation = InStr(response2, "status") + 9
    statusReport = Mid(response2, statusLocation, 9)
    'MsgBox statusReport
    
    rValue = True
    
    If StrComp(statusReport, "COMPLETED") = 0 Then
        rValue = False
    End If
    
    check_status = rValue
    
End Function

Public Function get_result(url As String) As String

    Set MyResult = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyResult.Open "GET", url, False
    MyResult.setRequestHeader "Content-type", "application/json"
    MyResult.setRequestHeader "Accept", "application/json"
    MyResult.setRequestHeader "Authorization", "ApiKey " & Me.mAPI_key
    
    MyResult.Send
    response = MyResult.ResponseText
    
    
    get_result = response
    
End Function

Public Function list_models() As String

    Dim Destination As String
    Destination = "http://" & Me.mURL & "api/models"
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Destination, False
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "Accept", "application/json"
    MyRequest.setRequestHeader "Authorization", "ApiKey " & Me.mAPI_key
    
    
    MyRequest.Send
    sHTML = MyRequest.ResponseText
    
    'And we get this response
    MsgBox sHTML
    
    ActiveDocument.Content.InsertAfter Text:=sHTML

End Function


Public Function test() As String

    Dim Destination As String
    Destination = "https://demo.modzy.engineering/api/jobs/"
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Destination, False
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "Accept", "application/json"
    MyRequest.setRequestHeader "Authorization", "ApiKey " & Me.mAPI_key
    
    
    MyRequest.Send
    sHTML = MyRequest.ResponseText
    
    'And we get this response
    MsgBox sHTML

End Function

Private Sub Class_Initialize()

End Sub






