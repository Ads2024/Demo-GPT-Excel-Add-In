Option Explicit

Function WorksheetExists(worksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
    On Error GoTo 0
End Function

Function ParseResponse(ByVal response As String) As String
    On Error Resume Next
    Dim jsonResponse As Object
    Set jsonResponse = JsonConverter.ParseJson(response)
    ParseResponse = jsonResponse("choices")(1)("message")("content")
    On Error GoTo 0
End Function

Function CleanJSONString(inputStr As String) As String
    On Error Resume Next
    CleanJSONString = Replace(inputStr, vbCrLf, "")
    CleanJSONString = Replace(CleanJSONString, vbCr, "")
    CleanJSONString = Replace(CleanJSONString, vbLf, "")
    CleanJSONString = Replace(CleanJSONString, """", "'")
    On Error GoTo 0
End Function

Function ReplaceBackslash(text As Variant) As String
    On Error Resume Next
    Dim i As Integer
    Dim newText As String
    newText = ""
    For i = 1 To Len(text)
        If Mid(text, i, 2) = "\" & Chr(34) Then
            newText = newText & Chr(34)
            i = i + 1
        Else
            newText = newText & Mid(text, i, 1)
        End If
    Next i
    ReplaceBackslash = newText
    On Error GoTo 0
End Function