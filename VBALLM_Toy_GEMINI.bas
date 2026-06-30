Option Explicit

Sub CallGemini()

    Dim apiKey As String
    Dim http As Object
    Dim url As String
    Dim prompt As String
    Dim body As String
    Dim json As Object
    Dim answer As String

    apiKey = InputBox("Gemini API Key를 입력하세요")
    If apiKey = "" Then Exit Sub

    prompt = Range("A1").Value
    If prompt = "" Then
        Range("B1").Value = "A1 셀에 질문을 입력하세요."
        Exit Sub
    End If

    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" & apiKey

    body = "{""contents"":[{""parts"":[{""text"":""" & EscapeJson(prompt) & """}]}]}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send body

    If http.Status <> 200 Then
        Range("B1").Value = "Error: " & http.Status & vbCrLf & http.ResponseText
        Exit Sub
    End If

    Set json = JsonConverter.ParseJson(http.ResponseText)

    answer = json("candidates")(1)("content")("parts")(1)("text")

    Range("B1").Value = answer

End Sub

Function EscapeJson(ByVal s As String) As String

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")

    EscapeJson = s

End Function
