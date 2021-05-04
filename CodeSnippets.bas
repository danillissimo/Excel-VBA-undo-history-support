Attribute VB_Name = "CodeSnippets"
Option Explicit

Private ignoreToDos As Boolean

Public Sub inc(ByRef value As Variant, Optional delta As Variant = 1)
    value = value + delta
End Sub

Public Sub dec(ByRef value As Variant, Optional delta As Variant = 1)
    value = value - delta
End Sub

Public Function getChar(str As String, position As Integer) As String
    getChar = Mid(str, position, 1)
End Function

Public Sub concat(ByRef str As String, value As Variant)
    str = str & value
End Sub

Public Sub shift(ByRef Range As Range, numLines As Integer, Optional numColumns As Integer = 0)
    Set Range = Range.Offset(numLines, numColumns)
End Sub

Public Sub todo(tip As String, Optional source As String = "")
    If ignoreToDos Then
        Exit Sub
    End If

    Dim answer As Integer
    answer = MsgBox(tip, vbCritical + vbAbortRetryIgnore, "Not implemented")
    If answer = vbIgnore Then
        ignoreToDos = True
    End If
    Debug.Assert answer = vbRetry Or answer = vbIgnore
End Sub
