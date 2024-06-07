'Description: This script contains utility functions used by the main script to interact with the OpenAI API and process the AI response.
'Created on: 2024-05-31
'Modified on: 2024-06-04
Option Explicit
Const OUTPUT_WORKSHEET As String = "Result"
' Name: ClearOutputSheet
' Description: Clears the output worksheet if it exists.
' Parameters:
'   - control: The IRibbonControl object that triggered the event.
Sub ClearOutputSheet(control As IRibbonControl)
    On Error Resume Next
    If WorksheetExists(OUTPUT_WORKSHEET) Then
        Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents
        MsgBox "Output worksheet cleared.", vbInformation, "Clear Output"
    Else
        MsgBox "Output worksheet does not exist.", vbExclamation, "Clear Output"
    End If
End Sub

' Name: ShowHelp
' Description: Displays a help message with information about the add-in.
' Parameters:
'   - control: The IRibbonControl object that triggered the event.
Sub ShowHelp(control As IRibbonControl)
    MsgBox "This add-in allows you to interact with OpenAI's ChatGPT model. Use 'Ask AI' to get responses based on your input. Use 'Clear Output' to clear the results sheet. Set your API Key and model using the settings buttons.", vbInformation, "Help"
End Sub

' Name: SetAPIKey
' Description: Prompts the user to enter their OpenAI API Key and sets it.
' Parameters:
'   - control: The IRibbonControl object that triggered the event.
Sub SetAPIKey(control As IRibbonControl)
    API_KEY = InputBox("Please enter your OpenAI API Key:", "Set API Key")
    If API_KEY = "" Then
        MsgBox "API Key not set. Please set it using the 'Set API Key' button.", vbExclamation, "Set API Key"
    Else
        MsgBox "API Key set successfully.", vbInformation, "Set API Key"
    End If
End Sub

' Name: SetModel
' Description: Prompts the user to enter the OpenAI model to use and sets it.
' Parameters:
'   - control: The IRibbonControl object that triggered the event.
Sub SetModel(control As IRibbonControl)
    MODEL = InputBox("Please enter the OpenAI model to use (e.g., gpt-3.5-turbo):", "Set Model", "gpt-3.5-turbo")
    If MODEL = "" Then
        MsgBox "Model not set. Defaulting to 'gpt-3.5-turbo'.", vbExclamation, "Set Model"
        MODEL = "gpt-3.5-turbo"
    Else
        MsgBox "Model set to " & MODEL & ".", vbInformation, "Set Model"
    End If
End Sub