Attribute VB_Name = "HybridProgramming"
Option Explicit

Public Const KEY_REDO = "^y"
Public Const KEY_UNDO = "^z"

'Go on, press it, I dare you
Public Const KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP = "+^%{F15}"
Public Const KEY_EVENT_CONTROL_TRANSACTION_REPLAY = "+^%{F14}"

Public Const KEYBOARD_LAYOUT_CODE_ENGLISH = 1033
Public Const KEYBOARD_LAYOUT_CODE_RUSSIAN = 1049
Public Const KEYBOARD_LAYOUT_CODE_NEXT = 1
Public Const KEYBOARD_LAYOUT_CODE_PREVIOUS = 0

'Num times to send undo\redo from keyboard to replay the whole transaction
'Basic value for the rest of the constants
Public Const TRANSACTION_NUM_ACTIONS = 5
'If we find ourselfs in a transaction, then at least one action is already done
Public Const TRANSACTION_MAX_REPLAY_ACTIONS = TRANSACTION_NUM_ACTIONS - 1
'Repeat undo\redo MAX_REPLAY times - if didn't get out of transaction, move in the opposite direction
'The direction we came from
'All undesired actions will be suppressed once out of transaction
Public Const TRANSACTION_NUM_ACTIONS_PER_PROGRAM = TRANSACTION_MAX_REPLAY_ACTIONS * 2

Public Type DWORD
    LoWord As Integer
    HiWord As Integer
End Type

Public Type TransactionData
    userUpdatedCell As Range
    userValue As String
    copySourceRange As Range
    copyTargetRange As Range
End Type

Public Enum ReplayDirection
    rdBackward = -1
    rdForward = 1
End Enum

Public transactionReplayIsRunning As Boolean
Public transactionReplayStep As Integer
Public transactionReplayIsSuppressed As Boolean
Public transactionReplayDirection As ReplayDirection

Public updateIsRunning As Boolean
Public pendingTransactions() As TransactionData
Public processedTransactionIndex As Integer
Public transactionStep As Integer

Public currentTransactionIndex As Integer

Public originalKeyboardLayoutCode As Integer

Public sheetSelectedBeforeTransaction As Worksheet

Public Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, Optional Flag As Long = 0) As Long
Public Declare PtrSafe Function GetKeyboardLayout Lib "user32" (Optional ByVal idThread As Long = 0) As DWORD

'This key combination has to be suppressed somehow, as it breaks the whole thing
Private Sub onShiftEnter()
    'Next doesn't handle all cases:
    'If ActiveCell.Row > 1 Then
    '    ActiveCell.Offset(-1, 0).Activate
    'End If
End Sub

Private Sub log(stepNum As Integer, begining As Boolean)
    'Application.Wait Now() + TimeValue("0:0:01")
    'Debug.Print "step " & stepNum & IIf(begining, " begining", " ending")
End Sub

Private Sub hideTransactionIndexContainer()
    TransactionIndexContainer.Visible = xlSheetVeryHidden
End Sub

Private Sub performNextTransactionStep()
    Select Case transactionStep
        Case 0
            log 0, True
            TransactionIndexContainer.Visible = xlSheetVisible
            TransactionIndexContainer.Select
            TransactionIndexContainer.Range(RANGE_TRANSACTION_INDEX).Select
            'Next will be emulated input of next transaction index
            inc transactionStep
            log 0, False
        Case 1
            log 1, True
            pendingTransactions(processedTransactionIndex).userUpdatedCell.Parent.Select
            pendingTransactions(processedTransactionIndex).userUpdatedCell.Select
            'Next will be emulated input of pendingTransactions(processedTransactionIndex).userValue
            inc transactionStep
            log 1, False
        Case 2
            log 2, True
            pendingTransactions(processedTransactionIndex).copyTargetRange.Parent.Select
            pendingTransactions(processedTransactionIndex).copyTargetRange.Select
            pendingTransactions(processedTransactionIndex).copySourceRange.Copy
            'Next will be emulated paste
            inc transactionStep
            log 2, False
        Case 3
            log 3, True
            TransactionIndexContainer.Visible = xlSheetVisible
            TransactionIndexContainer.Select
            TransactionIndexContainer.Range(RANGE_TRANSACTION_INDEX).Select
            'next will be emulated deletion
            inc transactionStep
            log 3, False
        Case 4
            log 4, True
            'finalize previous step
            If processedTransactionIndex = UBound(pendingTransactions) Then
                ActivateKeyboardLayout originalKeyboardLayoutCode, 0
                pendingTransactions(processedTransactionIndex).userUpdatedCell.Parent.Select
                pendingTransactions(processedTransactionIndex).userUpdatedCell.Offset(1, 0).Select
                hideTransactionIndexContainer
                Debug.Print "All pending transactions are processed"
                updateIsRunning = False
            End If
            
            transactionStep = 0
            
            inc currentTransactionIndex
            inc processedTransactionIndex
            
            log 4, False
        Case Else
            MsgBox "processedTransactionIndex corrupted!", vbCritical
    End Select
End Sub

'This is used when user presses an un/redo and enters a transaction
'We don't know how much un/redos we need, so we issue maximum possible number, and supress them once
'desired state is reached
Private Sub controlTransactionReplay()
    'Application.Wait Now() + TimeValue("0:0:01")
    Debug.Print "proc " & transactionReplayStep
    
    If transactionReplayStep < TRANSACTION_NUM_ACTIONS_PER_PROGRAM Then
        inc transactionReplayStep
    Else
        hideTransactionIndexContainer
        If TransactionIndexContainer.Range(RANGE_TRANSACTION_INDEX).value <> "" Then
            MsgBox "Ќе удалось завершить воспроизведение транзакции ни в одном из направлений! «акройте файл без сохранени€ и откройте снова (все последние изменени€ будут утер€ны), либо запустите проверку целостности (данные могут быть поверждены, но могут быть и восстановлены). ƒальнейша€ работа без прин€ти€ каких либо мер насто€тельно не рекомендуетс€, т.к. может вызвать повреждени€ данных! ¬о избежание таких ошибок в будущем, не выполн€йте никаких действий пока программа зан€та (не нажимайте никаких клавиш или кнопок, не мен€йте активный системный €зык).", vbCritical + vbOKOnly, "ќшибка!"
        End If
        transactionReplayStep = 0
        transactionReplayIsSuppressed = False
        ActivateKeyboardLayout originalKeyboardLayoutCode
        Application.OnKey KEY_UNDO
        Application.OnKey KEY_REDO
        transactionReplayIsRunning = False
        sheetSelectedBeforeTransaction.Select
        Exit Sub
    End If
    
    If transactionReplayIsSuppressed Then
        Debug.Print "suppressed"
        Exit Sub
    End If
    
    If TransactionIndexContainer.Range(RANGE_TRANSACTION_INDEX).value = "" Then
        Debug.Print "suppressing"
        transactionReplayIsSuppressed = True
        'Assign empty handlers to un/redo keys, so they stop doing what expected
        Application.OnKey KEY_UNDO, ""
        Application.OnKey KEY_REDO, ""
    End If
End Sub

'Some debug helpers down there

Private Sub getKBLayout()
    Debug.Print GetKeyboardLayout().LoWord
End Sub

Private Sub resetBig()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    updateIsRunning = False
    transactionStep = 0
    
    'transactionReplayIsIssued = False
    transactionReplayStep = 0
    transactionReplayIsSuppressed = False
    transactionReplayIsRunning = False
    
    currentTransactionIndex = 0
    
    Application.OnKey KEY_UNDO
    Application.OnKey KEY_REDO
End Sub

Private Sub resetSmall()
    Application.EnableEvents = False
End Sub
