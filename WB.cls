VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    todo _
    "1)Need to make sure, that closing the book returns the original autocalucalation mode" & _
    Chr(10) & _
    "2)ConstructorFactory got to be formatted like #ConflictSample# when the book is opened" & _
    Chr(10) & _
    "3)Sheets can't be copied or deleted in shared mode. Preparing an array of formatted templates is likely to be the solution. Keep in mind user history keeps up to 100 actions. It can be configured through registry, but it's unlikely to happen" & _
    Chr(10) & _
    "4)Got to make shure that system supports English on startup"
    
    cleanOldConstructionAreas
    registerKeyEvents
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    cleanOldConstructionAreas
End Sub

'Formatted data samples must not be modified to prevent corruption of user history
'Thus, they are kept until document is closed\reopened
Private Sub cleanOldConstructionAreas()
    Dim ws As Worksheet
    
    For Each ws In WB.Worksheets
        Application.DisplayAlerts = False
        'C-tor{->16 hex values<-}
        If ws.name Like "C-tor{[0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF][0123456789ABCDEF]}" Then
            ws.Visible = xlSheetHidden
            ws.Delete
        End If
        Application.DisplayAlerts = True
    Next ws
End Sub

Public Sub registerKeyEvents()
    Application.OnKey KEY_EVENT_PERFORM_NEXT_TRANSACTION_STEP, "performNextTransactionStep"
    Application.OnKey KEY_EVENT_CONTROL_TRANSACTION_REPLAY, "controlTransactionReplay"
End Sub
