Option Explicit

Sub ClearOutputSheet(control As IRibbonControl)
    On Error Resume Next
    If WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents
        MsgBox "Output worksheet cleared.", vbInformation, "Clear Output"
    Else
        MsgBox "Output worksheet does not exist.", vbExclamation, "Clear Output"
    End If
End Sub

Sub ShowHelp(control As IRibbonControl)
    MsgBox "This add-in allows you to interact with OpenAI's ChatGPT model. Use 'Ask AI' to get responses based on your input. Use 'Clear Output' to clear the results sheet. Set your API Key and model using the settings buttons.", vbInformation, "Help"
End Sub

Sub SetAPIKey(control As IRibbonControl)
    API_KEY = InputBox("Please enter your OpenAI API Key:", "Set API Key")
    If API_KEY = "" Then
        MsgBox "API Key not set. Please set it using the 'Set API Key' button.", vbExclamation, "Set API Key"
    Else
        MsgBox "API Key set successfully.", vbInformation, "Set API Key"
    End If
End Sub

Sub SetModel(control As IRibbonControl)
    MODEL = InputBox("Please enter the OpenAI model to use (e.g., gpt-3.5-turbo):", "Set Model", "gpt-3.5-turbo")
    If MODEL = "" Then
        MsgBox "Model not set. Defaulting to 'gpt-3.5-turbo'.", vbExclamation, "Set Model"
        MODEL = "gpt-3.5-turbo"
    Else
        MsgBox "Model set to " & MODEL & ".", vbInformation, "Set Model"
    End If
End Sub