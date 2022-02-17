VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Application_Quit()
    fMainFlowTimerStop
    'formLog.Hide
End Sub

Private Sub Application_Startup()
    fMainFlowTimerStart
    'formLog.Show
End Sub

'Public Sub fLogFormShowHideTrigger()
    'If formLog.Visible Then
    '    formLog.Hide
    'Else
    '    formLog.Show
    'End If
'End Sub

