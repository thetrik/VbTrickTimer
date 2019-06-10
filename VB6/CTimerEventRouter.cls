VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimerEventRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cTimer As CTrickTimer
Attribute m_cTimer.VB_VarHelpID = -1

Public Property Get Timer() As CTrickTimer
    Set Timer = m_cTimer
End Property

Public Property Set Timer( _
                    ByVal cValue As CTrickTimer)
    Set m_cTimer = cValue
End Property

Private Sub m_cTimer_Tick()
    frmTimers.Tick m_cTimer
End Sub
