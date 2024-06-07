' This Module contains the UI-related code for the Azure OpenAI Excel Add-In.
' Created on 31/05/2024
' Modified on 7/06/2024

Option Explicit

' This constant represents the name of the output worksheet where the results will be displayed.
Const OUTPUT_WORKSHEET As String = "Result"

' This subroutine clears the contents of the output worksheet.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub ClearOutputSheet(control As IRibbonControl)
    On Error Resume Next
    If WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents
        MsgBox "Output worksheet cleared.", vbInformation, "Clear Output"
    Else
        MsgBox "Output worksheet does not exist.", vbExclamation, "Clear Output"
    End If
End Sub

' This subroutine displays a help message to the user.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub ShowHelp(control As IRibbonControl)
    MsgBox "This add-in allows you to interact with AzureOpenAI. Use 'Ask AI' to get responses based on your input. Use 'Clear Output' to clear the results sheet. Set your API Key, endpoint, and model using the settings buttons.", vbInformation, "Help"
End Sub

' This subroutine sets the API Key for the AzureOpenAI service.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub SetAPIKey(control As IRibbonControl)
    ' Prompt the user to enter the API Key
    API_KEY = InputBox("Please enter your OpenAI API Key:", "Set API Key")
    If API_KEY = "" Then
        MsgBox "API Key not set. Please set it using the 'Set API Key' button.", vbExclamation, "Set API Key"
    Else
        MsgBox "API Key set successfully.", vbInformation, "Set API Key"
    End If
End Sub

' This subroutine sets the AzureOpenAI model to be used.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub SetModel(control As IRibbonControl)
    ' Prompt the user to enter the model name
    MODEL = InputBox("Please enter the your AzureOpenAI model Deployment name to use (e.g., gpt-3.5-turbo):", "Set Model", "gpt-3.5-turbo")
    If MODEL = "" Then
        MsgBox "Model not set. Defaulting to 'gpt-3.5-turbo'.", vbExclamation, "Set Model"
        MODEL = "gpt-3.5-turbo"
    Else
        MsgBox "Model set to " & MODEL & ".", vbInformation, "Set Model"
    End If
End Sub

' This subroutine sets the Azure OpenAI Endpoint.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub SetAzureEndpoint(control As IRibbonControl)
    ' Prompt the user to enter the Azure OpenAI Endpoint
    AZURE_OPENAI_ENDPOINT = InputBox("Please enter your Azure OpenAI Endpoint:", "Set Azure OpenAI Endpoint")
    If AZURE_OPENAI_ENDPOINT = "" Then
        MsgBox "Azure OpenAI Endpoint not set. Please set it using the 'Set Azure OpenAI Endpoint' button.", vbExclamation, "Set Azure OpenAI Endpoint"
    Else
        MsgBox "Azure OpenAI Endpoint set successfully.", vbInformation, "Set Azure OpenAI Endpoint"
    End If
End Sub

' This subroutine sets the API Version.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
Sub SetAPIVersion(control As IRibbonControl)
    ' Prompt the user to enter the API Version
    API_VERSION = InputBox("Please enter the API version:", "Set API Version", "2024-02-15-preview")
    If API_VERSION = "" Then
        MsgBox "API Version not set. Defaulting to '2024-02-15-preview'.", vbExclamation, "Set API Version"
        API_VERSION = "2024-02-15-preview"
    Else
        MsgBox "API Version set to " & API_VERSION & ".", vbInformation, "Set API Version"
    End If
End Sub

' This subroutine shows or hides the task pane.
' It takes a control parameter of type IRibbonControl, which is not used in this subroutine.
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


