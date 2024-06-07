' Description: This script contains utility functions used by the main script to interact with the OpenAI API and process the AI response.
' Created on: 2024-05-31
' Modified on: 2024-06-04

Option Explicit
'---------------------------------------------------------------------------
' Function: WorksheetExists
'
' Description: Checks if a worksheet with the specified name exists in the workbook.
'
' Parameters:
'   - worksheetName: The name of the worksheet to check.
'
' Returns:
'   - Boolean: True if the worksheet exists, False otherwise.
'---------------------------------------------------------------------------
Function WorksheetExists(worksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------
' Function: ParseResponse
'
' Description: Parses the JSON response from the OpenAI API and returns the content of the message.
'
' Parameters:
'   - response: The JSON response string.
'
' Returns:
'   - String: The content of the message in the JSON response.
'---------------------------------------------------------------------------
Function ParseResponse(ByVal response As String) As String
    On Error Resume Next
    Dim jsonResponse As Object
    Set jsonResponse = JsonConverter.ParseJson(response)
    ParseResponse = jsonResponse("choices")(1)("message")("content")
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------
' Function: CleanJSONString
'
' Description: Cleans a JSON string by removing line breaks and replacing double quotes with single quotes.
'
' Parameters:
'   - inputStr: The input JSON string.
'
' Returns:
'   - String: The cleaned JSON string.
'---------------------------------------------------------------------------
Function CleanJSONString(inputStr As String) As String
    On Error Resume Next
    CleanJSONString = Replace(inputStr, vbCrLf, "")
    CleanJSONString = Replace(CleanJSONString, vbCr, "")
    CleanJSONString = Replace(CleanJSONString, vbLf, "")
    CleanJSONString = Replace(CleanJSONString, """", "'")
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------
' Function: ReplaceBackslash
'
' Description: Replaces backslashes followed by double quotes with just double quotes in a text string.
'
' Parameters:
'   - text: The input text string.
'
' Returns:
'   - String: The text string with backslashes followed by double quotes replaced with just double quotes.
'---------------------------------------------------------------------------
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

'---------------------------------------------------------------------------
' Function: GetAIResponse
'
' Description: Sends a query to the OpenAI API and returns the response.
'
' Parameters:
'   - query: The query to send to the API.
'
' Returns:
'   - String: The response from the OpenAI API.
'---------------------------------------------------------------------------
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
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & API_KEY
    
    ' Send the request
    http.send requestBody
    
    ' Get the response
    responseBody = http.responseText
    
    ' Extract the AI response
    GetAIResponse = responseBody
End Function