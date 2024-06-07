' Description: This script contains utility functions used by the main script to interact with the OpenAI API and process the AI response.
' Created on: 2024-05-31
' Modified on: 2024-06-04

Option Explicit

' Function to check if a worksheet exists in the workbook
' Returns True if the worksheet exists, False otherwise
Function WorksheetExists(worksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
    On Error GoTo 0
End Function

' Function to parse the response from the AI model
' Returns the content of the message in the response
Function ParseResponse(ByVal response As String) As String
    On Error Resume Next
    Dim jsonResponse As Object
    Set jsonResponse = JsonConverter.ParseJson(response)
    ParseResponse = jsonResponse("choices")(1)("message")("content")
    On Error GoTo 0
End Function

' Function to clean a JSON string by removing line breaks and replacing double quotes with single quotes
' Returns the cleaned JSON string
Function CleanJSONString(inputStr As String) As String
    On Error Resume Next
    CleanJSONString = Replace(inputStr, vbCrLf, "")
    CleanJSONString = Replace(CleanJSONString, vbCr, "")
    CleanJSONString = Replace(CleanJSONString, vbLf, "")
    CleanJSONString = Replace(CleanJSONString, """", "'")
    On Error GoTo 0
End Function

' Function to replace backslashes followed by double quotes with just double quotes
' Returns the string with replaced backslashes
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

' Utility function to get AI response using the OpenAI API
' Returns the response from the AI model
Function GetAIResponse(query As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim responseBody As String
    
    ' Create an XMLHTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Define the URL and request body
    url = "https://api.openai.com/v1/engines/" & MODEL & "/completions"
    requestBody = "{""prompt"":""" & query & """,""max_tokens"":100}"
    
    ' Open an HTTP POST request
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & API_KEY
    
    ' Send the request
    http.Send requestBody
    
    ' Get the response
    responseBody = http.responseText
    
    ' Extract the AI response
    GetAIResponse = responseBody
End Function

