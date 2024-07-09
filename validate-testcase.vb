Enum CharType
    Kanji = 1
    Hiragana = 2
    katakana = 3
    Romanji = 4
End Enum

Function GenerateRandomString(allowedChars As String, minLength As Integer, maxLength As Integer) As String
    Dim strLength As Integer
    Dim result As String
    Dim i As Integer
    Dim randIndex As Integer
    strLength = Int((maxLength - minLength + 1) * Rnd + minLength)

    result = ""
    For i = 1 To strLength
        randIndex = Int((Len(allowedChars) * Rnd) + 1)
        result = result & Mid(allowedChars, randIndex, 1)
    Next i

    GenerateRandomString = result
End Function


Function GenerateString(charType As CharType, minLength As Integer, maxLength As Integer)
    Select Case charType
        Case CharType.Kanji
            GenerateString = GenerateRandomString("漢字の例", minLength, maxLength)
        Case CharType.Hiragana
            GenerateString = GenerateRandomString("ひらがなの例", minLength, maxLength)
        Case CharType.katakana
            GenerateString = GenerateRandomString("カタカナ", minLength, maxLength)
        Case CharType.Romanji
            GenerateString = GenerateRandomString("abcdEDFG", minLength, maxLength)
    End Select

End Function

Function CreateRequestDataPost(dict As Scripting.Dictionary)
    Dim result As String
    result = "{" & vbCrLf

    Dim key As Variant
    Dim value As Variant
    For Each key In dict.Keys
        value = dict(key)
        If VarType(dict(key)) = vbString Then
            result = result & "  """ & key & """:""" & dict(key) & """" & vbCrLf
        Else
            result = result & "  """ & key & """:" & dict(key) & vbCrLf
        End If
    Next key
    result = result & "}"

    CreateRequestDataPost = result
End Function

Function CreateRequestDataGet(dict As Scripting.Dictionary)
    Dim result As String
    result = ""

    Dim key As Variant
    For Each key In dict.Keys
        result = resul & key & ":" & dict(key) & vbCrLf
    Next key

    CreateRequestDataGet = result
End Function


Sub GenerateTestcase()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim coll As New Scripting.Dictionary
    coll.Add "Age" 30
    coll.Add "City" "Da Nang"


    Dim i As Integer
    For i = 1 To 5
        coll.Add Item:=GenerateString(i, 10, 20), Key:="Name"
        ws.Cells(i, 1).value = CreateRequestDataPost(coll)
    Next i
End Sub