VERSION 5.00
Begin VB.Form frmTimers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTrickTimer test"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetInterval 
      Caption         =   "Set interval..."
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemoveTimer 
      Caption         =   "Remove timer"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   780
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddTimer 
      Caption         =   "Add timer..."
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtLog 
      Height          =   1995
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2220
      Width           =   4995
   End
   Begin VB.ListBox lstTimers 
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // frmMain - form for CTrickTimer.cls testing
' // by The trick, 2019
' //

Option Explicit

' // The timer's event routers collection
Private m_cTimers   As Collection

' // Log to textbox
Public Sub Tick( _
           ByVal cTimer As CTrickTimer)
    
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = vbNewLine & "[0x" & Hex$(ObjPtr(cTimer)) & "] " & cTimer.Tag & " " & Format$(Timer, "0.000")
    
End Sub

' // Add a timer
Private Sub cmdAddTimer_Click()
    Dim sInterval   As String
    Dim lInterval   As Long
    Dim cTimer      As CTrickTimer
    Dim cRouter     As CTimerEventRouter
    
    sInterval = InputBox("Set interval (ms.)", , "1000")
    
    If Not IsNumeric(sInterval) Then Exit Sub
    
    lInterval = Val(sInterval)
    
    Set cTimer = New CTrickTimer
    
    ' // Set tag to display on log textbox
    cTimer.Tag = InputBox("Set tag")
    cTimer.Interval = lInterval
    
    Set cRouter = New CTimerEventRouter
    ' // Assign timer to handle event
    Set cRouter.Timer = cTimer
    
    m_cTimers.Add cRouter
    
    lstTimers.AddItem "[0x" & Hex$(ObjPtr(cTimer)) & "] " & cTimer.Tag & " - " & CStr(lInterval) & " ms."
    
End Sub

' // Remove timer
Private Sub cmdRemoveTimer_Click()
    
    If lstTimers.ListIndex = -1 Then Exit Sub
    
    m_cTimers.Remove (lstTimers.ListIndex + 1)
    lstTimers.RemoveItem lstTimers.ListIndex
    
End Sub

' // Update interval
Private Sub cmdSetInterval_Click()
    Dim sInterval   As String
    Dim lInterval   As Long
    Dim cTimer      As CTrickTimer
    Dim cRouter     As CTimerEventRouter
    
    If lstTimers.ListIndex = -1 Then Exit Sub
    
    Set cRouter = m_cTimers(lstTimers.ListIndex + 1)
    Set cTimer = cRouter.Timer
    
    sInterval = InputBox("Set interval (ms.)", , CStr(cTimer.Interval))
    
    If Not IsNumeric(sInterval) Then Exit Sub
    
    lInterval = Val(sInterval)
    
    cTimer.Interval = lInterval
    
    lstTimers.List(lstTimers.ListIndex) = "[0x" & Hex$(ObjPtr(cTimer)) & "] " & cTimer.Tag & " - " & CStr(lInterval) & " ms."
    
End Sub

Private Sub Form_Initialize()
    
    Set m_cTimers = New Collection
    
End Sub

Private Sub Form_Terminate()
    
    Set m_cTimers = Nothing
    
End Sub


