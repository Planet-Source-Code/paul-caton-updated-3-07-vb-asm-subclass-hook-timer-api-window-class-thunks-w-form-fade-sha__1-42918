VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bDone As Boolean
Private Timer As cTimer
Implements WinSubHook.iTimer

Private Sub Class_Initialize()
  Set Timer = New cTimer
  Call Timer.Start(100, Me)
  
  MsgBox "Timer events are appearing every 100 ms in the immediate window." & vbNewLine & vbNewLine & _
            "VB timers will only work with a Form/UserControl." & vbNewLine & _
            "VB timers stop when a MsgBox is active in the IDE.", vbInformation
  
  Set Timer = Nothing
End Sub

Private Sub iTimer_Fire(ByVal lElapsedMS As Long)
  Static nCount As Long
  
  nCount = nCount + 1
  Debug.Print "#" & nCount, "Total elapsed time: " & Format(lElapsedMS / 1000, "000.000")
End Sub
