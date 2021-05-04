VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaffingBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim originalCalcMode As Long

Private Sub Worksheet_Activate()
    'If shouldIgnoreWorksheetSelectionEvents Then
    '    Exit Sub
    'End If
    
    Application.OnKey "+~", "onShiftEnter"
    'better safe than sorry
    WB.registerKeyEvents
    
    originalCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
End Sub

Private Sub Worksheet_Deactivate()
    'If shouldIgnoreWorksheetSelectionEvents Then
    '    Exit Sub
    'End If
    
    Application.OnKey "+~"
    Application.Calculation = originalCalcMode
End Sub

'==================================================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCells As Range
    Dim cell As Range
    Dim i As Integer
    Dim nextTransactionIndex As Integer
    Dim program As String

    If updateIsRunning Or transactionReplayIsRunning Then
        Exit Sub
    End If
    
    Set changedCells = Application.Intersect(StaffingBook.Range(RANGE_STAFFING_IDS), Target)
    If changedCells Is Nothing Then
        Exit Sub
    End If
    
    ReDim pendingTransactions(1 To changedCells.Count) As TransactionData
    i = 1
    For Each cell In changedCells
        Set pendingTransactions(i).userUpdatedCell = cell
        pendingTransactions(i).userValue = cell.value
        Set pendingTransactions(i).copyTargetRange = cell.Offset(0, 1)
        Set pendingTransactions(i).copySourceRange = getPersonalDataForUpload(cell.value)
        inc i
    Next cell
    
    nextTransactionIndex = currentTransactionIndex + 1
    For i = LBound(pendingTransactions) To UBound(pendingTransactions)
        'PROGRAM prepare for transaction index input
        'EMULATE input transaction index
        'PROGRAM prepare for new value input; keep in mind that new value can be "" which can only be emulated as {DEL}
        'EMULATE input new value
        'PROGRAM copy personal data
        'EMULATE paste
        'PROGRAM prepare for transaction index removal
        'EMULATE remove transaction index
        'PROGRAM finalize
        concat _
            program, _
            KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP & _
            nextTransactionIndex & "~" & _
            KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP & _
            "{DEL}" & pendingTransactions(i).userValue & "~" & _
            KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP & _
            "~" & _
            KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP & _
            "{DEL}" & _
            KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP
        inc nextTransactionIndex
    Next i
    
    processedTransactionIndex = LBound(pendingTransactions)
    Debug.Print program
    updateIsRunning = True
    originalKeyboardLayoutCode = GetKeyboardLayout().LoWord
    ActivateKeyboardLayout KEYBOARD_LAYOUT_CODE_ENGLISH, 0
    SendKeys KEY_UNDO & program, True
End Sub

Private Function getPersonalDataForUpload(ids As String)
    Dim cell As Range
    Dim result As Range
    Dim cursor As Integer
    Dim firstId As String
    
    ids = LCase(ids)
    
    cursor = 1
    firstId = getNextWord(ids, cursor)
    
    If cursorIsAboutToLeaveText(ids, cursor) Then
        checkSpecialIds ids
        Set getPersonalDataForUpload = searchForPersonalData(ids)
    Else
        Dim values() As String
        Dim ws As Worksheet
        ReDim values(1 To EMPLOYEE_NUM_DATA_COLUMNS) As String
        
        Set ws = createBufferForOwercrowdedPosition()
        With ws
            Set getPersonalDataForUpload = .Range(.Range("a1"), .Range("a1").Offset(0, EMPLOYEE_NUM_DATA_COLUMNS - 1))
        End With

        processEntryForOvecrowdedPosition firstId, values, True
        Do
            processEntryForOvecrowdedPosition getNextWord(ids, cursor), values, False
        Loop While Not cursorIsAboutToLeaveText(ids, cursor)

        Set cell = ws.Range("a1")
        For cursor = LBound(values) To UBound(values)
            cell.value = values(cursor)
            shift cell, 0, 1
        Next cursor
    End If
End Function

Private Sub processEntryForOvecrowdedPosition(id As String, ByRef resultContainer() As String, firstEntry As Boolean)
    Dim cell As Range
    Dim personalData As Range
    Dim i As Integer
    
    checkSpecialIds id
    
    Set personalData = searchForPersonalData(id)
    Set cell = personalData.Range("A1")
    
    If firstEntry Then
        For i = LBound(resultContainer) To UBound(resultContainer)
            resultContainer(i) = cell.value
            shift cell, 0, 1
        Next i
    Else
        For i = LBound(resultContainer) To UBound(resultContainer)
            concat resultContainer(i), Chr(10) & cell.value
            shift cell, 0, 1
        Next i
    End If
End Sub

Private Function searchForPersonalDataUnsafe(id As String)
    Dim cell As Range

    'Nuh-huh, UsedRange clears the undo history
    'Set id = Intersect(Employee.Range(RANGE_EMPLOYES_IDS).EntireColumn, Employee.UsedRange)
    
    For Each cell In Employee.Range(RANGE_EMPLOYES_IDS)
        If cell.value = id Then
            Set searchForPersonalDataUnsafe = Employee.Range( _
                cell.Offset(0, 1).Address, _
                cell.Offset(0, EMPLOYEE_NUM_DATA_COLUMNS).Address _
            )
            Exit Function
        ElseIf cell.value = "" Then
            Exit For
        End If
    Next cell
    
    Set searchForPersonalDataUnsafe = Nothing
End Function

Private Function searchForPersonalData(id As String)
    Set searchForPersonalData = searchForPersonalDataUnsafe(id)

    If searchForPersonalData Is Nothing Then
        Set searchForPersonalData = searchForPersonalDataUnsafe(ID_ERROR)
    End If

    If searchForPersonalData Is Nothing Then
        'TODO Some actions should be taken if found nothing
    End If
End Function

Private Sub checkSpecialIds(ByRef id As String)
    If id = SPECIAL_VALUE_VACANT Then
        id = ID_VACANT
    ElseIf id = SPECIAL_VALUE_EMPTY Then
        id = ID_EMPTY
    End If
End Sub

'Any programmatic action, taken too close to the data, being edited by the user, will immediately clear the action history
'Thus, a lot can be done with worksheets, unavailable to user
'A lot, but not all
'Changing the values of such a sheet is allowed
'Cell formatting is still limited
'This is the moment when constructor factory appears
'Factory is a template, containing all required formats
'When some values got to be created on the fly, a copy of this template is created and populated with created data
'And then it can be published to user
'The only problem is that once it is referenced by action history, it can't be programmatically altered without corrupting action history anymore
'So it needs to be kept until the file is closed or opened
'You can see the cleanup function in the WB module
Private Function createBufferForOwercrowdedPosition() As Worksheet
    'Worksheet.Copy works completly unpredictable and does not return the new sheet
    'That's the reason for the strange staff going on here
    Dim ws As Worksheet
    Dim name As String
    Dim cell As Range
    
    With WorksheetFunction
        name = _
            "C-tor{" & _
            .Dec2Hex(.RandBetween(0, 65535), 4) & _
            .Dec2Hex(.RandBetween(0, 65535), 4) & _
            .Dec2Hex(.RandBetween(0, 65535), 4) & _
            .Dec2Hex(.RandBetween(0, 65535), 4) & _
            "}"
    End With
    ConstructorFactory.Range(RANGE_CONSTRUCTOR_NEW_NAME).value = name
    ConstructorFactory.Visible = xlSheetHidden
    ConstructorFactory.Copy After:=ConstructorFactory
    ConstructorFactory.Visible = xlSheetVeryHidden
    ConstructorFactory.Range(RANGE_CONSTRUCTOR_NEW_NAME).value = ""
    
    For Each ws In WB.Sheets
        If ws.Range(RANGE_CONSTRUCTOR_NEW_NAME).value = name Then
            Exit For
        End If
    Next ws
    ws.Visible = xlSheetVeryHidden
    ws.name = name
    
    Set createBufferForOwercrowdedPosition = ws
End Function

