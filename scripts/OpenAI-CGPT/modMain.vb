Option Explicit

Const OUTPUT_WORKSHEET As String = "Result"

Sub OpenAI_Completion(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    If API_KEY = "" Then
        MsgBox "Please set a valid API key. You can set it using the 'Set API Key' button.", vbCritical, "No API Key Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If MODEL = "" Then
        MODEL = "gpt-3.5-turbo"
    End If

    Dim prompt As String
    Dim cell As Range
    Dim selectedRange As Range
    Set selectedRange = Selection

    For Each cell In selectedRange
        prompt = prompt & cell.Value & " "
    Next cell

    If Trim(prompt) <> "" Then
        prompt = CleanJSONString(prompt)
    Else
        MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If Not WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets.Add(After:=Sheets(Sheets.Count)).Name = OUTPUT_WORKSHEET
    End If

    Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents
    Application.StatusBar = "Processing OpenAI request..."

    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    Dim requestBody As String
    requestBody = "{" & _
        """model"": """ & MODEL & """," & _
        """messages"": [{""role"": ""user"", ""content"": """ & prompt & """}]," & _
        """max_tokens"": 1024," & _
        """temperature"": 0.5" & _
        "}"

    With httpRequest
        .Open "POST", "https://api.openai.com/v1/chat/completions", False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & API_KEY
        .send (requestBody)
    End With

    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText

        Dim completion As String
        completion = ParseResponse(response)

        Dim lines As Variant
        lines = Split(completion, "\n")

        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            Worksheets(OUTPUT_WORKSHEET).Cells(i + 1, 1).Value = ReplaceBackslash(lines(i))
        Next i

        Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
        MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"

        With Worksheets(OUTPUT_WORKSHEET)
            .Activate
            .Range("A1").Select
            .Tab.Color = RGB(169, 208, 142)
        End With

    Else
        MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub