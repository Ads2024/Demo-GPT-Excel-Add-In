' This subroutine is triggered when the OpenAI_Completion button is clicked on the Excel ribbon.
' It retrieves the selected range of cells and sends the cell values as a prompt to the OpenAI API for completion.
' The completed text is then written to a new worksheet named "Result".
' If the API key is not set or the selected cell range is empty, appropriate error messages are displayed.
' The completed text is split into lines and written to the "Result" worksheet.
' The "Result" worksheet is formatted and a success message is displayed.
' If the API request fails, an error message is displayed.


' Created on 31/05/2024
' Modified on 4/06/2024


Option Explicit

Const OUTPUT_WORKSHEET As String = "Result"
Const MAX_PROMPT_LENGTH As Long = 4096 ' Maximum length for the prompt

Sub OpenAI_Completion(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Check if API key is set
    If API_KEY = "" Then
        MsgBox "Please set a valid API key. You can set it using the 'Set API Key' button.", vbCritical, "No API Key Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Set default model if not specified
    If MODEL = "" Then
        MODEL = "gpt-3.5-turbo"
    End If
    
    ' Retrieve selected range of cells
    Dim prompt As String
    Dim cell As Range
    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Concatenate cell values as prompt
    For Each cell In selectedRange
        prompt = prompt & cell.Value & " "
    Next cell

    ' Check if prompt is not empty
    If Trim(prompt) <> "" Then
        prompt = CleanJSONString(prompt)
    Else
        MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Create "Result" worksheet if it doesn't exist
    If Not WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets.Add(After:=Sheets(Sheets.Count)).Name = OUTPUT_WORKSHEET
    End If

    ' Send request to OpenAI API
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

    ' Process API response
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText

        Dim completion As String
        completion = ParseResponse(response)

        Dim lines As Variant
        lines = Split(completion, vbLf)

        Dim i As Long
        Dim nextRow As Long
        nextRow = Worksheets(OUTPUT_WORKSHEET).Cells(Worksheets(OUTPUT_WORKSHEET).Rows.Count, 1).End(xlUp).Row + 1

        ' Write completed text to "Result" worksheet
        For i = LBound(lines) To UBound(lines)
            Worksheets(OUTPUT_WORKSHEET).Cells(nextRow, 1).Value = ReplaceBackslash(lines(i))
            nextRow = nextRow + 1
        Next i

        ' Format "Result" worksheet
        Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
        MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"

        ' Activate "Result" worksheet and highlight cell A1
        With Worksheets(OUTPUT_WORKSHEET)
            .Activate
            .Range("A1").Select
            .Tab.Color = RGB(2, 19, 158)
        End With

    Else
        ' Display error message if API request fails
        MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
    End If

    ' Clean up and restore application settings
    Application.StatusBar = False
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Display error message if an error occurs
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub
