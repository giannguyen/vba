Public Const HEAD = "{" & vbCrLf & String(4, " ") & """ver"": ""0.0.1""," & vbCrLf & String(4, " ") & """errors"": [" & vbCrLf


Public Const TAIL = String(4, " ") & "]" & vbCrLf & "}"

Public Const BLANK_MSG = ""

Public Const SIZE_MSG = ""

Public Const HALF_SIZE_MSG = ""

Public Const DATE_MSG = ""


Public Function CreateBody(errorCode As Integer, variableName As String, errorMessage As String, minLength As Integer, maxLength As Integer) As String
    Dim result As String
    result = ""
    errorMessage = Replace(errorMessage, "{minLength}", minLength)
    errorMessage = Replace(errorMessage, "{maxLength}", maxLength)
    result = result & String(8, " ") & "{" & vbCrLf & String(12, " ") & """code"": " & errorCode & "," & vbCrLf & String(12, " ") & """field"": """ & variableName & """," & _
                String(12, " ") & """message"": """ & errorMessage & """" & vbCrLf & String(8, "") & "}"
    CreateBody = result
End Function


