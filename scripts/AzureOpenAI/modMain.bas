Option Explicit

Const OUTPUT_WORKSHEET As String = "Result"

' Global variable for the task pane
Public taskPane As Office.CustomTaskPane

Public Sub OpenAI_Completion(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    If API_KEY = "" Then
        MsgBox "Please set a valid API key. You can set it using the 'Set API Key' button.", vbCritical, "No API Key Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If AZURE_OPENAI_ENDPOINT = "" Then
        MsgBox "Please set a valid Azure OpenAI Endpoint. You can set it using the 'Set Azure OpenAI Endpoint' button.", vbCritical, "No Azure OpenAI Endpoint Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If API_VERSION = "" Then
        MsgBox "Please set a valid API version. You can set it using the 'Set API Version' button.", vbCritical, "No API Version Found"
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
        .Open "POST", AZURE_OPENAI_ENDPOINT & "/openai/deployments/" & MODEL & "/chat/completions?api-version=" & API_VERSION, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-key", API_KEY
        .Send (requestBody)
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

Public Sub ShowTaskPane(control As IRibbonControl)
    ' Check if the task pane is already created
    If taskPane Is Nothing Then
        ' Create the task pane with the UserForm
        Set taskPane = Application.CreateCustomTaskPane("frmTaskPane", "ChatGPT Task Pane")
        ' Show the task pane
        taskPane.Visible = True
    Else
        ' Toggle visibility
        taskPane.Visible = Not taskPane.Visible
    End If
End Sub







