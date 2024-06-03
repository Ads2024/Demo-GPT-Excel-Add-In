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
    MsgBox "This add-in allows you to interact with OpenAI's ChatGPT model. Use 'Ask AI' to get responses based on your input. Use 'Clear Output' to clear the results sheet. Set your API Key, endpoint, and model using the settings buttons.", vbInformation, "Help"
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

Sub SetAzureEndpoint(control As IRibbonControl)
    AZURE_OPENAI_ENDPOINT = InputBox("Please enter your Azure OpenAI Endpoint:", "Set Azure OpenAI Endpoint")
    If AZURE_OPENAI_ENDPOINT = "" Then
        MsgBox "Azure OpenAI Endpoint not set. Please set it using the 'Set Azure OpenAI Endpoint' button.", vbExclamation, "Set Azure OpenAI Endpoint"
    Else
        MsgBox "Azure OpenAI Endpoint set successfully.", vbInformation, "Set Azure OpenAI Endpoint"
    End If
End Sub

Sub SetAPIVersion(control As IRibbonControl)
    API_VERSION = InputBox("Please enter the API version:", "Set API Version", "2023-05-15")
    If API_VERSION = "" Then
        MsgBox "API Version not set. Defaulting to '2023-05-15'.", vbExclamation, "Set API Version"
        API_VERSION = "2023-05-15"
    Else
        MsgBox "API Version set to " & API_VERSION & ".", vbInformation, "Set API Version"
    End If
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


