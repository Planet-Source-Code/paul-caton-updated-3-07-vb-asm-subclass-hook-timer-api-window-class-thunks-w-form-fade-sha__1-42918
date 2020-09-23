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

Private nWnds  As Long        'Number of windows
Private Window As cWindow     'Window class instance

Implements WinSubHook.iWindow 'We're implementing the interface

Private Sub Class_Initialize()
  Dim hWnd As Long
  
  Set Window = New cWindow    'Create the window class instance
  With Window
    Set .Owner = Me           'Set the owner
    
    'Add the messages that we care about
    Call .AddMsg(WM_CLOSE)
    Call .AddMsg(WM_NCHITTEST)
    
    'Register the window class
    Call .WindowClassRegister("ApiWindowClass")
    
    'Create and show a window
    hWnd = .WindowCreate(0, WS_POPUPWINDOW Or WS_CAPTION Or WS_SYSMENU, , 100, 100, 200, 200, "API Window 1")
    Call WinSubHook.ShowWindow(hWnd, WinSubHook.SW_SHOWNOACTIVATE)
    
    'Create and show another window
    hWnd = .WindowCreate(WS_EX_CLIENTEDGE, WS_OVERLAPPEDWINDOW, , 150, 150, 200, 200, "API Window 2")
    Call WinSubHook.ShowWindow(hWnd, WinSubHook.SW_SHOWNOACTIVATE)
  End With
  
  'Loop until both windows are closed
  nWnds = 2
  Do While nWnds > 0
    DoEvents
  Loop
  
  Set Window = Nothing
End Sub

Private Sub iWindow_WndProc(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
  Const HTCAPTION As Long = 2
  Const HTCLIENT  As Long = 1
  
  Select Case uMsg
  Case WM_CLOSE
    'Destroy the window
    Call Window.WindowDestroy(hWnd)
    nWnds = nWnds - 1
    
  Case WM_NCHITTEST
    'Call the default window proc first, if it says the mouse is over the client area...
    If Window.CallDefWndProc(hWnd, uMsg, wParam, lParam) = HTCLIENT Then
      lReturn = HTCAPTION 'Lie that the mouse is over the caption. You can now click and drag on the client area
      bHandled = True     'Don't call DefWndProc hereafter
    End If
    
  End Select
End Sub
