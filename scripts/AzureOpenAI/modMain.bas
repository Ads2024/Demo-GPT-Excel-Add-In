' This module contains the main code for the Azure OpenAI Excel Add-In.
' Created on 31/05/2024
' Modified on 4/06/2024
Option Explicit

Const OUTPUT_WORKSHEET As String = "Result"

' Global variable for the task pane
Public taskPane As Office.CustomTaskPane

' This subroutine is triggered when the "OpenAI_Completion" button is clicked on the Ribbon.
' It performs the OpenAI completion request using the selected range of cells as input.
Public Sub OpenAI_Completion(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Check if the API key is set
    If API_KEY = "" Then
        MsgBox "Please set a valid API key. You can set it using the 'Set API Key' button.", vbCritical, "No API Key Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Check if the Azure OpenAI Endpoint is set
    If AZURE_OPENAI_ENDPOINT = "" Then
        MsgBox "Please set a valid Azure OpenAI Endpoint. You can set it using the 'Set Azure OpenAI Endpoint' button.", vbCritical, "No Azure OpenAI Endpoint Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Check if the API version is set
    If API_VERSION = "" Then
        MsgBox "Please set a valid API version. You can set it using the 'Set API Version' button.", vbCritical, "No API Version Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Set the default model if not specified
    If MODEL = "" Then
        MODEL = "gpt-3.5-turbo"
    End If

    Dim prompt As String
    Dim cell As Range
    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Concatenate the values of the selected cells as the prompt
    For Each cell In selectedRange
        prompt = prompt & cell.Value & " "
    Next cell

    ' Check if the prompt is not empty
    If Trim(prompt) <> "" Then
        prompt = CleanJSONString(prompt)
    Else
        MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Create the output worksheet if it doesn't exist
    If Not WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets.Add(After:=Sheets(Sheets.Count)).Name = OUTPUT_WORKSHEET
    End If

    Application.StatusBar = "Processing OpenAI request..."

    ' Create an HTTP request object
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Create the request body in JSON format
    Dim requestBody As String
    requestBody = "{" & _
        """model"": """ & MODEL & """," & _
        """messages"": [{""role"": ""user"", ""content"": """ & prompt & """}]," & _
        """max_tokens"": 1024," & _
        """temperature"": 0.5" & _
        "}"

    ' Send the HTTP POST request to the Azure OpenAI Endpoint
    With httpRequest
        .Open "POST", AZURE_OPENAI_ENDPOINT & "/openai/deployments/" & MODEL & "/chat/completions?api-version=" & API_VERSION, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-key", API_KEY
        .Send (requestBody)
    End With

    ' Check if the request was successful
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText

        ' Parse the response to extract the completion
        Dim completion As String
        completion = ParseResponse(response)

        ' Split the completion into lines
        Dim lines As Variant
        lines = Split(completion, vbLf)

        Dim i As Long
        Dim nextRow As Long
        nextRow = Worksheets(OUTPUT_WORKSHEET).Cells(Worksheets(OUTPUT_WORKSHEET).Rows.Count, 1).End(xlUp).Row + 1

        ' Write each line of the completion to the output worksheet
        For i = LBound(lines) To UBound(lines)
            Worksheets(OUTPUT_WORKSHEET).Cells(nextRow, 1).Value = ReplaceBackslash(lines(i))
            nextRow = nextRow + 1
        Next i

        ' Auto-fit the columns in the output worksheet
        Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
        MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"

        ' Activate the output worksheet and highlight the first cell
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

