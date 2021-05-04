VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionIndexContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Begin transaction replay once an index apperas out of nowhere
Private Sub Worksheet_Change(ByVal Target As Range)
    If updateIsRunning Then
        Exit Sub
    End If
    
    If Target.Address(False, False) <> RANGE_TRANSACTION_INDEX Then
        Exit Sub
    End If
    
    If Not Target.value = "" Then
        If Not transactionReplayIsRunning Then
            transactionReplayIsRunning = True
            transactionReplayStep = 0
            transactionReplayIsSuppressed = False
            Set sheetSelectedBeforeTransaction = Application.ActiveSheet
            
            originalKeyboardLayoutCode = GetKeyboardLayout().LoWord
            ActivateKeyboardLayout KEYBOARD_LAYOUT_CODE_ENGLISH
            
            If transactionReplayDirection = rdForward Then
                transactionReplayDirection = IIf(Target.value >= currentTransactionIndex, rdForward, rdBackward)
            Else
                transactionReplayDirection = IIf(Target.value > currentTransactionIndex, rdForward, rdBackward)
            End If
            
            SendKeys _
                buildTransactionReplayProgram(transactionReplayDirection), _
                False
        End If
        currentTransactionIndex = Target.value
    Else
        Debug.Print "Index " & currentTransactionIndex & " erased"
        inc currentTransactionIndex, transactionReplayDirection
        Debug.Print "New index: " & currentTransactionIndex
    End If
End Sub

Private Function buildTransactionReplayProgram(direction As ReplayDirection) As String
    Dim cmdForward As String
    Dim cmdBackward As String
    Dim result As String
    Dim i As Integer
    
    If direction = rdForward Then
        cmdForward = KEY_REDO
        cmdBackward = KEY_UNDO
    Else
        cmdForward = KEY_UNDO
        cmdBackward = KEY_REDO
    End If
    
    For i = 1 To TRANSACTION_MAX_REPLAY_ACTIONS
        concat result, KEY_EVENT_CONTROL_TRANSACTION_REPLAY & cmdForward
    Next i
    For i = 1 To TRANSACTION_MAX_REPLAY_ACTIONS
        concat result, KEY_EVENT_CONTROL_TRANSACTION_REPLAY & cmdBackward
    Next i
    
    'Escape to remove focus from dummy key-event receiver if any
    buildTransactionReplayProgram = result & KEY_EVENT_CONTROL_TRANSACTION_REPLAY
    Debug.Print "Generated transaction replay program"
    Debug.Print buildTransactionReplayProgram
End Function
