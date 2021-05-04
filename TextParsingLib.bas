Attribute VB_Name = "TextParsingLib"
Option Compare Text
Option Explicit

'lib343.TextParsingLib

Public Enum CursorMovementDirection
    cmdForward = 1
    cmdBackward = -1
End Enum

'Creates a temporary new cursor that matches the original
'Used in conjunction with the functions of this library to prevent changes to the cursor position
'Example: getNextWord(text, lockCursor(cursor))
Public Function lockCursor(cursor As Integer) As Integer
    lockCursor = cursor
End Function

'Skips the group of spaces the cursor is currently pointing to
'And returns all text, starting from the character under the cursor, up to the next space, or to the end of the line
'Positions the cursor at a space after a word, or at the edge of a line
'
'For example, getNextWord(text, cursor, cmdForward):
'1-
'address: somemail@email.com
'        ^
'      cursor
'Returns "somemail@email.com"
'
'2-
'address: somemail@email.com
'            ^
'          cursor
'Returns "email@email.com"
Public Function getNextWord(text As String, cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As String
    skipSpacesIfAny text, cursor, direction
    getNextWord = getTextBeforeFirstSpace(text, cursor, direction)
End Function

'Skips the group of spaces the cursor is currently pointing to
'Returns all text, starting from the character under the cursor, to the next separator, or to the end of the line
'Positions the cursor on a separator after a word, or on the edge of a line
'
'For example, getNextWordPart(text, cursor, cmdForward):
'1-
'address: somemail@email.com
'        ^
'      cursor
'Returns "somemail"
'
'2-
'address: somemail@email.com
'            ^
'          cursor
'Returns "email"
Public Function getNextWordPart(text As String, cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As String
    skipSpacesIfAny text, cursor, direction
    getNextWordPart = getTextBeforeFirstSpecialSymbol(text, cursor, direction)
End Function

'==================================================================================================

Public Function cursorIsWithinText(text As String, cursor As Integer) As Boolean
    cursorIsWithinText = cursor >= 1 And cursor <= Len(text)
End Function

Public Function cursorIsAboutToLeaveText(text As String, cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As Boolean
    Select Case direction
        Case cmdForward
            cursorIsAboutToLeaveText = cursor >= Len(text)
        Case cmdBackward
            cursorIsAboutToLeaveText = cursor <= 1
    End Select
    'cursorIsAboutToLeaveText = cursor <= 1 Or cursor >= Len(text)
End Function

Public Function isSpecialSymbol(char As String) As Boolean
    isSpecialSymbol = (Not isDigit(char)) And (Not isLetter(char))
End Function

Public Function isDigit(char As String) As Boolean
    Dim code As Integer
    code = Asc(char)
    isDigit = code > 47 And code < 58
End Function

'Made purely for the Russian symbol table, not according to the general standard
Public Function isLetter(char As String) As Boolean
    Dim code As Integer
    code = Asc(char)
    isLetter = (code > 64 And code < 91) Or (code > 96 And code < 123) Or (code > 191)
End Function
'==================================================================================================

'Positions the cursor behind the first group of spaces it finds
Public Sub skipSpaces(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward)
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        If getChar(text, cursor) <> " " Then
            inc cursor, direction
        Else
            Exit Do
        End If
    Loop
    
    skipSpacesIfAny text, cursor, direction
End Sub

'==================================================================================================

'Checks if the cursor is pointing to a group of spaces and places the cursor immediately after it or at the end of the line
Public Sub skipSpacesIfAny(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward)
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        If getChar(text, cursor) = " " Then
            inc cursor, direction
        Else
            Exit Do
        End If
    Loop
End Sub

'==================================================================================================

'Places the cursor immediately after a group of separators, or at the edge of a line
Public Sub skipSpecialSymbols(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward)
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        If Not isSpecialSymbol(getChar(text, cursor)) Then
            inc cursor, direction
        Else
            Exit Do
        End If
    Loop
    
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        If isSpecialSymbol(getChar(text, cursor)) Then
            inc cursor, direction
        Else
            Exit Do
        End If
    Loop
End Sub

'==================================================================================================

'Searches for the first separator starting at the next character after the cursor
'Returns false if hit the edge of the line, otherwise true
'Sets the cursor on the found separator
Public Function skipToNextSpecialSymbol(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As Boolean
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        inc cursor, direction
        If isSpecialSymbol(getChar(text, cursor)) Then
            skipToNextSpecialSymbol = True
            Exit Function
        End If
    Loop
    
    skipToNextSpecialSymbol = False
End Function

'==================================================================================================

'Searches for the first space starting from the character following the cursor
'Returns false if hit the edge of the line, otherwise true
'Sets the cursor on the found space
Public Function skipToNextSpace(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As Boolean
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        inc cursor, direction
        If getChar(text, cursor) = " " Then
            skipToNextSpace = True
            Exit Function
        End If
    Loop
    
    skipToNextSpace = False
End Function

'==================================================================================================

'Returns the remainder of the word up to the first separator in the given direction, including the character pointed to by the cursor
'Sets the cursor
'   1 behind the separator, if the cursor originally pointed to it
'   2 on the separator by word
'   3 at the edge of the line
Public Function getTextBeforeFirstSpecialSymbol(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As String
    Dim char As String * 1
    
    char = getChar(text, cursor)
    If isSpecialSymbol(char) Then
        getTextBeforeFirstSpecialSymbol = char
        If Not cursorIsAboutToLeaveText(text, cursor, direction) Then
            inc cursor, direction
        End If
        Exit Function
    End If
    
    
    
    Dim begining As Integer
    Dim corrector As Integer
    
    begining = cursor
    corrector = IIf(skipToNextSpecialSymbol(text, cursor, direction), 1, 0)

    If begining < cursor Then
        getTextBeforeFirstSpecialSymbol = Mid(text, begining, cursor - corrector - begining + 1)
    Else
        getTextBeforeFirstSpecialSymbol = Mid(text, cursor + corrector, begining - (cursor + corrector) + 1)
    End If
End Function

'==================================================================================================

'Returns the remainder of the word up to the first space in the given direction, including the character pointed to by the cursor
'Sets the cursor at a space behind a word, or at the edge of a line
Private Function getTextBeforeFirstSpace(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As String
    Dim begining As Integer
    Dim corrector As Integer
    
    begining = cursor
    corrector = IIf(skipToNextSpace(text, cursor, direction), 1, 0)

    If begining < cursor Then
        getTextBeforeFirstSpace = Mid(text, begining, cursor - corrector - begining + 1)
    Else
        getTextBeforeFirstSpace = Mid(text, cursor + corrector, begining - (cursor + corrector) + 1)
    End If
End Function

'==================================================================================================

Public Function containsDigits(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If isDigit(getChar(text, i)) Then
            containsDigits = True
            Exit Function
        End If
    Next i
    
    containsDigits = False
End Function

'Doesn't accept any separators in text
Public Function consistsOfDigitsOnly(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If Not isDigit(getChar(text, i)) Then
            consistsOfDigitsOnly = False
            Exit Function
        End If
    Next i
    
    consistsOfDigitsOnly = True
End Function

Public Function containsLetters(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If isLetter(getChar(text, i)) Then
            containsLetters = True
            Exit Function
        End If
    Next i
    
    containsLetters = False
End Function

'Doesn't accept any separators in text
Public Function consistsOfLettersOnly(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If Not isLetter(getChar(text, i)) Then
            consistsOfLettersOnly = False
            Exit Function
        End If
    Next i
    
    consistsOfLettersOnly = True
End Function

Public Function containsSpecialSymbols(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If isSpecialSymbol(getChar(text, i)) Then
            containsSpecialSymbols = True
            Exit Function
        End If
    Next i
    
    containsSpecialSymbols = False
End Function

Public Function consistsOfSpecialSymbolsOnly(text As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(text)
        If Not isSpecialSymbol(getChar(text, i)) Then
            consistsOfSpecialSymbolsOnly = False
            Exit Function
        End If
    Next i
    
    consistsOfSpecialSymbolsOnly = True
End Function

'Searches for the first separator except spaces, starting at the character behind the cursor
'Returns its value on success, otherwise an empty string
'Sets the cursor on a separator or at the end of a line
Public Function getNextSpecialSymbol(text As String, ByRef cursor As Integer, Optional direction As CursorMovementDirection = cmdForward) As String
    Dim symbol As String * 1
    
    Do While Not cursorIsAboutToLeaveText(text, cursor, direction)
        inc cursor, direction
        symbol = getChar(text, cursor)
        If isSpecialSymbol(symbol) And symbol <> " " Then
            getNextSpecialSymbol = symbol
            Exit Function
        End If
    Loop
    
    getNextSpecialSymbol = ""
End Function
