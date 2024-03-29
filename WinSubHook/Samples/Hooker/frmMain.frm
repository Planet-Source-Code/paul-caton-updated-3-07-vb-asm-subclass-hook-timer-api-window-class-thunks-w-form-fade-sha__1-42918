VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hooker"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   717
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      Height          =   6735
      Left            =   7695
      ScaleHeight     =   6675
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      Begin VB.CheckBox chkAfter 
         Caption         =   "After"
         Height          =   195
         Left            =   2070
         TabIndex        =   6
         Top             =   2525
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "Before"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   2525
         Width           =   795
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3585
         Width           =   2685
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set Hook"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   3
         Top             =   2905
         Width           =   2685
      End
      Begin VB.ListBox lstHooks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "frmMain.frx":0000
         Left            =   180
         List            =   "frmMain.frx":0023
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   2685
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   7230
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   12753
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For XP manifests
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'API declarations
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Constants
Private Const SM_CXVSCROLL              As Long = 2
Private Const LVM_FIRST                 As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH        As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE            As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER  As Long = -2

Private Const HDR_WH_CALLWNDPROC        As String = "Process,hWnd,uMsg,wParam,lParam"
Private Const HDR_WH_CALLWNDPROCRET     As String = "Process,hWnd,uMsg,wParam,lParam,lResult"
Private Const HDR_WH_CBT                As String = "Code,wParam,lParam"
Private Const HDR_WH_FOREGROUNDIDLE     As String = "Action"
Private Const HDR_WH_GETMESSAGE         As String = "Process,hWnd,uMsg,wParam,lParam,X,Y,Time"
Private Const HDR_WH_KEYBOARD           As String = "Process,vKey,Repeat,Scan,Ext'd,Alt,Prev,Curr"
Private Const HDR_WH_MOUSE              As String = "Process,ID,X,Y,hWnd,Hit,Extra"
Private Const HDR_WH_MSGFILTER          As String = "Process,hWnd,uMsg,wParam,lParam,X,Y,Time"
Private Const HDR_WH_SHELL              As String = "Code,wParam,lParam"
    
Private ht As eHookType             'Hook type
Private hc As cHook                 'Hooker

Implements WinSubHook.iHook         'Hook interface

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Set hc = New cHook
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    lv.Move 0, 0, ScaleWidth - pic.Width, ScaleHeight
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  hc.UnHook
  Set hc = Nothing
End Sub

Private Sub iHook_After(lReturn As Long, ByVal nCode As WinSubHook.eHookCode, ByVal wParam As Long, ByVal lParam As Long)
  Dim cwp As WinSubHook.tCWPSTRUCT
  
  If chkAfter.Value = vbChecked Then
    Select Case ht
      Case eHookType.WH_CALLWNDPROC:    Call procWH_CALLWNDPROC("After", nCode, wParam, lParam)
      Case eHookType.WH_CALLWNDPROCRET: Call procWH_CALLWNDPROCRET("After", nCode, wParam, lParam)
      Case eHookType.WH_CBT:            Call procWH_CBT("After", nCode, wParam, lParam)
      Case eHookType.WH_FOREGROUNDIDLE: Call procWH_FOREGROUNDIDLE("After", nCode, wParam, lParam)
      Case eHookType.WH_GETMESSAGE:     Call procWH_GETMESSAGE("After", nCode, wParam, lParam)
      Case eHookType.WH_KEYBOARD:       Call procWH_KEYBOARD("After", nCode, wParam, lParam)
      Case eHookType.WH_MOUSE:          Call procWH_MOUSE("After", nCode, wParam, lParam)
      Case eHookType.WH_MSGFILTER:      Call procWH_MSGFILTER("After", nCode, wParam, lParam)
      Case eHookType.WH_SHELL:          Call procWH_SHELL("After", nCode, wParam, lParam)
    End Select
  End If
End Sub

Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, nCode As WinSubHook.eHookCode, wParam As Long, lParam As Long)
  If chkBefore.Value = vbChecked Then
    Select Case ht
      Case eHookType.WH_CALLWNDPROC:    Call procWH_CALLWNDPROC("Before", nCode, wParam, lParam)
      Case eHookType.WH_CALLWNDPROCRET: Call procWH_CALLWNDPROCRET("Before", nCode, wParam, lParam)
      Case eHookType.WH_CBT:            Call procWH_CBT("Before", nCode, wParam, lParam)
      Case eHookType.WH_FOREGROUNDIDLE: Call procWH_FOREGROUNDIDLE("Before", nCode, wParam, lParam)
      Case eHookType.WH_GETMESSAGE:     Call procWH_GETMESSAGE("Before", nCode, wParam, lParam)
      Case eHookType.WH_KEYBOARD:       Call procWH_KEYBOARD("Before", nCode, wParam, lParam)
      Case eHookType.WH_MOUSE:          Call procWH_MOUSE("Before", nCode, wParam, lParam)
      Case eHookType.WH_MSGFILTER:      Call procWH_MSGFILTER("Before", nCode, wParam, lParam)
      Case eHookType.WH_SHELL:          Call procWH_SHELL("Before", nCode, wParam, lParam)
    End Select
  End If
End Sub

Private Sub cmdSet_Click()
  Const CAP_SET   As String = "Set Hook"
  Const CAP_UNSET As String = "Unhook"
  Dim i           As Long
  Dim s           As String
  Dim sArray()    As String
  
  If cmdSet.Caption = CAP_UNSET Then
    Call hc.UnHook
    cmdSet.Caption = CAP_SET
    cmdSet.Enabled = False
    lstHooks.Enabled = True
    lstHooks.ListIndex = -1
    txtInfo = ""
    If ht = WH_MSGFILTER Then
      On Error Resume Next
      Call Unload(frmMenu)
    End If
    Exit Sub
  End If
  
  cmdSet.Caption = CAP_UNSET
  lstHooks.Enabled = False
  
  Call lv.ListItems.Clear
  
  With lv.ColumnHeaders
    Select Case ht
      Case eHookType.WH_CALLWNDPROC:    s = HDR_WH_CALLWNDPROC
      Case eHookType.WH_CALLWNDPROCRET: s = HDR_WH_CALLWNDPROCRET
      Case eHookType.WH_CBT:            s = HDR_WH_CBT
      Case eHookType.WH_FOREGROUNDIDLE: s = HDR_WH_FOREGROUNDIDLE
      
      Case eHookType.WH_KEYBOARD:       s = HDR_WH_KEYBOARD
      Case eHookType.WH_MOUSE:          s = HDR_WH_MOUSE
      Case eHookType.WH_SHELL:          s = HDR_WH_SHELL
      Case eHookType.WH_GETMESSAGE
        s = HDR_WH_GETMESSAGE
        MsgBox "To prevent recursion, callbacks from the list view aren't shown.", vbInformation
        
      Case eHookType.WH_MSGFILTER
        s = HDR_WH_MSGFILTER
        frmMenu.Show vbModeless, Me
        
    End Select
  
    If s = "" Then
      lv.Visible = False
    Else
      lv.Visible = True
      Call .Clear
      Call .Add(, , "When")
      sArray = Split(s, ",")
      For i = LBound(sArray) To UBound(sArray)
        Call .Add(, , sArray(i))
      Next i
    End If
  End With
  
  Call SizeColumns(lv)
  Call hc.Hook(ht, Me)
End Sub

Private Sub lstHooks_Click()
  If lstHooks.ListIndex = -1 Then
    Exit Sub
  End If
  
  ht = lstHooks.ItemData(lstHooks.ListIndex)
  cmdSet.Enabled = True
  
  Select Case lstHooks.ItemData(lstHooks.ListIndex)
    Case WinSubHook.WH_CALLWNDPROC:     txtInfo = "The system calls this function whenever the SendMessage function is called. Before passing the message to the destination window procedure, the system passes the message to the hook procedure. The hook procedure can examine the message; it cannot modify it."
    Case WinSubHook.WH_CALLWNDPROCRET:  txtInfo = "The system calls this function after the SendMessage function is called. The hook procedure can examine the message; it cannot modify it."
    Case WinSubHook.WH_CBT:             txtInfo = "The system calls this function before activating, creating, destroying, minimizing, maximizing, moving, or sizing a window; before completing a system command; before removing a mouse or keyboard event from the system message queue; before setting the keyboard focus; or before synchronizing with the system message queue. A computer-based training (CBT) application uses this hook procedure to receive useful notifications from the system."
    Case WinSubHook.WH_FOREGROUNDIDLE:  txtInfo = "The system calls this function whenever the foreground thread is about to become idle."
    Case WinSubHook.WH_GETMESSAGE:      txtInfo = "The system calls this function whenever the GetMessage function has retrieved a message from an application message queue. Before passing the retrieved message to the destination window procedure, the system passes the message to the hook procedure."
    Case WinSubHook.WH_KEYBOARD:        txtInfo = "The system calls this function whenever an application calls the GetMessage or PeekMessage function and there is a keyboard message (WM_KEYUP or WM_KEYDOWN) to be processed."
    Case WinSubHook.WH_MOUSE:           txtInfo = "The system calls this function whenever an application calls the GetMessage or PeekMessage function and there is a mouse message to be processed."
    Case WinSubHook.WH_MSGFILTER:       txtInfo = "The system calls this function after an input event occurs in a dialog box, message box, menu, or scroll bar, but before the message generated by the input event is processed. The hook procedure can monitor messages for a dialog box, message box, menu, or scroll bar created by the application."
    Case WinSubHook.WH_SHELL:           txtInfo = "The function receives notifications of shell events from the system."
  End Select
End Sub

Private Sub procWH_CALLWNDPROC(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim dat As WinSubHook.tCWPSTRUCT
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  dat = hc.xCWPSTRUCT(lParam)
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .SubItems(2) = fmt(dat.hwnd)
    .SubItems(3) = GetMsgName(dat.message)
    .SubItems(4) = fmt(dat.wParam)
    .SubItems(5) = fmt(dat.lParam)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_CALLWNDPROCRET(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim dat As WinSubHook.tCWPRETSTRUCT
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  dat = hc.xCWPRETSTRUCT(lParam)
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .SubItems(2) = fmt(dat.hwnd)
    .SubItems(3) = GetMsgName(dat.message)
    .SubItems(4) = fmt(dat.wParam)
    .SubItems(5) = fmt(dat.lParam)
    .SubItems(6) = fmt(dat.lResult)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_CBT(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHCBT(nCode)
    .SubItems(2) = fmt(wParam)
    .SubItems(3) = fmt(lParam)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_FOREGROUNDIDLE(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_GETMESSAGE(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim dat As WinSubHook.tMSG
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  dat = hc.xMSG(lParam)
  If dat.hwnd = lv.hwnd Then
    Exit Sub
  End If
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .SubItems(2) = fmt(dat.hwnd)
    .SubItems(3) = GetMsgName(dat.message)
    .SubItems(4) = fmt(dat.wParam)
    .SubItems(5) = fmt(dat.lParam)
    .SubItems(6) = dat.pt.x
    .SubItems(7) = dat.pt.y
    .SubItems(8) = fmt(dat.Time)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_KEYBOARD(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim itm As MSComctlLib.ListItem
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .SubItems(2) = fmt(wParam, 4)
    .SubItems(3) = fmt(lParam And &HFFFF&, 4)
    .SubItems(4) = fmt((lParam And &HFF0000) Mod &HFF, 2)
    .SubItems(5) = IIf(lParam And &H1000000, "Ext", "")
    .SubItems(6) = IIf(lParam And &H20000000, "Dn", "Up")
    .SubItems(7) = IIf(lParam And &H40000000, "Dn", "Up")
    .SubItems(8) = IIf(lParam And &H80000000, "Dn", "Up")
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
End Sub

Private Sub procWH_MOUSE(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim dat As tMOUSEHOOKSTRUCT
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  dat = hc.xMOUSEHOOKSTRUCT(lParam)
    
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHC(nCode)
    .SubItems(2) = GetMsgName(wParam)
    .SubItems(3) = dat.pt.x
    .SubItems(4) = dat.pt.y
    .SubItems(5) = fmt(dat.hwnd)
    .SubItems(6) = fmt(dat.wHitTestCode)
    .SubItems(7) = fmt(dat.dwExtraInfo)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_MSGFILTER(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim dat As WinSubHook.tMSG
  Dim itm As MSComctlLib.ListItem

  On Error GoTo Catch
  
  dat = hc.xMSG(lParam)
    
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeMSGF(nCode)
    .SubItems(2) = fmt(dat.hwnd)
    .SubItems(3) = GetMsgName(dat.message)
    .SubItems(4) = fmt(dat.wParam)
    .SubItems(5) = fmt(dat.lParam)
    .SubItems(6) = dat.pt.x
    .SubItems(7) = dat.pt.y
    .SubItems(8) = fmt(dat.Time)
    .EnsureVisible
  End With
  
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Sub procWH_SHELL(sWhen As String, nCode As eHookCode, wParam As Long, lParam As Long)
  Dim itm As MSComctlLib.ListItem
  
  On Error GoTo Catch
  
  Set itm = lv.ListItems.Add(, , sWhen)
  With itm
    .SubItems(1) = DecodeHSHELL(nCode)
    .SubItems(2) = fmt(wParam)
    .SubItems(3) = fmt(lParam)
    .EnsureVisible
  End With
    
  Call SizeColumns(lv)
  DoEvents
Catch:
End Sub

Private Function DecodeHC(nCode As eHookCode) As String
  Select Case nCode
    Case Is < 0:                                DecodeHC = "Pass"
    Case eHookCode.HC_ACTION:                   DecodeHC = "Action"
    Case eHookCode.HC_GETNEXT:                  DecodeHC = "Get Next"
    Case eHookCode.HC_NOREM:                    DecodeHC = "No Remove"
    Case eHookCode.HC_NOREMOVE:                 DecodeHC = "No Remove"
    Case eHookCode.HC_SKIP:                     DecodeHC = "Skip"
    Case eHookCode.HC_SYSMODALOFF:              DecodeHC = "SysModalOff"
    Case eHookCode.HC_SYSMODALON:               DecodeHC = "SysModalOn"
  End Select
End Function

Private Function DecodeHCBT(nCode As eHookCode) As String
  Select Case nCode
    Case eHookCode.HCBT_ACTIVATE:               DecodeHCBT = "Activate"
    Case eHookCode.HCBT_CLICKSKIPPED:           DecodeHCBT = "Click Skipped"
    Case eHookCode.HCBT_CREATEWND:              DecodeHCBT = "Create Wnd"
    Case eHookCode.HCBT_DESTROYWND:             DecodeHCBT = "Destroy Wnd"
    Case eHookCode.HCBT_KEYSKIPPED:             DecodeHCBT = "Key Skipped"
    Case eHookCode.HCBT_MINMAX:                 DecodeHCBT = "MinMax"
    Case eHookCode.HCBT_MOVESIZE:               DecodeHCBT = "Move Size"
    Case eHookCode.HCBT_QS:                     DecodeHCBT = "QS"
    Case eHookCode.HCBT_SETFOCUS:               DecodeHCBT = "SetFocus"
    Case eHookCode.HCBT_SYSCOMMAND:             DecodeHCBT = "SysCommand"
  End Select
End Function

Private Function DecodeHSHELL(nCode As eHookCode) As String
  Select Case nCode
    Case eHookCode.HSHELL_ACTIVATESHELLWINDOW:  DecodeHSHELL = "ActivShellWnd"
    Case eHookCode.HSHELL_GETMINRECT:           DecodeHSHELL = "GetMinRect"
    Case eHookCode.HSHELL_LANGUAGE:             DecodeHSHELL = "Language"
    Case eHookCode.HSHELL_REDRAW:               DecodeHSHELL = "Redraw"
    Case eHookCode.HSHELL_TASKMAN:              DecodeHSHELL = "Taskman"
    Case eHookCode.HSHELL_WINDOWACTIVATED:      DecodeHSHELL = "WndActivated"
    Case eHookCode.HSHELL_WINDOWCREATED:        DecodeHSHELL = "WndCreated"
    Case eHookCode.HSHELL_WINDOWDESTROYED:      DecodeHSHELL = "WndDestroyed"
  End Select
End Function

Private Function DecodeMSGF(nCode As eHookCode) As String
  Select Case nCode
    Case eHookCode.MSGF_DDEMGR:                 DecodeMSGF = "DdeMgr"
    Case eHookCode.MSGF_DIALOGBOX:              DecodeMSGF = "DlgBox"
    Case eHookCode.MSGF_MAX:                    DecodeMSGF = "Max"
    Case eHookCode.MSGF_MENU:                   DecodeMSGF = "Menu"
    Case eHookCode.MSGF_MESSAGEBOX:             DecodeMSGF = "MsgBox"
    Case eHookCode.MSGF_NEXTWINDOW:             DecodeMSGF = "NextWnd"
    Case eHookCode.MSGF_SCROLLBAR:              DecodeMSGF = "ScrlBar"
    Case eHookCode.MSGF_USER:                   DecodeMSGF = "User"
  End Select
End Function

Private Function DecodePM(nCode As eHookCode) As String
  Select Case nCode
    Case eHookCode.PM_NOREMOVE:                 DecodePM = "NoRemove"
    Case eHookCode.PM_NOYIELD:                  DecodePM = "NoYield"
    Case eHookCode.PM_REMOVE:                   DecodePM = "Remove"
  End Select
End Function

'Return the Value parameter converted to a hex string padded to 8 characters with a leading &H
Private Function fmt(Value As Long, Optional Resolution As Long = 8) As String
  Dim s As String
  
  s = Hex$(Value)
  fmt = String$(Resolution - Len(s), "0") & s
End Function

Private Sub SizeColumns(lv As MSComctlLib.ListView)
  Dim Col2adjust As Long
  
  Call LockWindowUpdate(lv.hwnd)
    For Col2adjust = 0 To lv.ColumnHeaders.Count - 1
      Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Col2adjust
  Call LockWindowUpdate(0)
End Sub

Private Function GetMsgName(uMsg As WinSubHook.eMsg)
  Select Case uMsg
  Case WinSubHook.WM_ACTIVATE:             GetMsgName = "WM_ACTIVATE"
  Case WinSubHook.WM_ACTIVATEAPP:          GetMsgName = "WM_ACTIVATEAPP"
  Case WinSubHook.WM_ASKCBFORMATNAME:      GetMsgName = "WM_ASKCBFORMATNAME"
  Case WinSubHook.WM_CANCELJOURNAL:        GetMsgName = "WM_CANCELJOURNAL"
  Case WinSubHook.WM_CANCELMODE:           GetMsgName = "WM_CANCELMODE"
  Case WinSubHook.WM_CAPTURECHANGED:       GetMsgName = "WM_CAPTURECHANGED"
  Case WinSubHook.WM_CHANGECBCHAIN:        GetMsgName = "WM_CHANGECBCHAIN"
  Case WinSubHook.WM_CHAR:                 GetMsgName = "WM_CHAR"
  Case WinSubHook.WM_CHARTOITEM:           GetMsgName = "WM_CHARTOITEM"
  Case WinSubHook.WM_CHILDACTIVATE:        GetMsgName = "WM_CHILDACTIVATE"
  Case WinSubHook.WM_CLEAR:                GetMsgName = "WM_CLEAR"
  Case WinSubHook.WM_CLOSE:                GetMsgName = "WM_CLOSE"
  Case WinSubHook.WM_COMMAND:              GetMsgName = "WM_COMMAND"
  Case WinSubHook.WM_COMPACTING:           GetMsgName = "WM_COMPACTING"
  Case WinSubHook.WM_COMPAREITEM:          GetMsgName = "WM_COMPAREITEM"
  Case WinSubHook.WM_COPY:                 GetMsgName = "WM_COPY"
  Case WinSubHook.WM_COPYDATA:             GetMsgName = "WM_COPYDATA"
  Case WinSubHook.WM_CREATE:               GetMsgName = "WM_CREATE"
  Case WinSubHook.WM_CTLCOLORBTN:          GetMsgName = "WM_CTLCOLORBTN"
  Case WinSubHook.WM_CTLCOLORDLG:          GetMsgName = "WM_CTLCOLORDLG"
  Case WinSubHook.WM_CTLCOLOREDIT:         GetMsgName = "WM_CTLCOLOREDIT"
  Case WinSubHook.WM_CTLCOLORLISTBOX:      GetMsgName = "WM_CTLCOLORLISTBOX"
  Case WinSubHook.WM_CTLCOLORMSGBOX:       GetMsgName = "WM_CTLCOLORMSGBOX"
  Case WinSubHook.WM_CTLCOLORSCROLLBAR:    GetMsgName = "WM_CTLCOLORSCROLLBAR"
  Case WinSubHook.WM_CTLCOLORSTATIC:       GetMsgName = "WM_CTLCOLORSTATIC"
  Case WinSubHook.WM_CUT:                  GetMsgName = "WM_CUT"
  Case WinSubHook.WM_DEADCHAR:             GetMsgName = "WM_DEADCHAR"
  Case WinSubHook.WM_DELETEITEM:           GetMsgName = "WM_DELETEITEM"
  Case WinSubHook.WM_DESTROY:              GetMsgName = "WM_DESTROY"
  Case WinSubHook.WM_DESTROYCLIPBOARD:     GetMsgName = "WM_DESTROYCLIPBOARD"
  Case WinSubHook.WM_DRAWCLIPBOARD:        GetMsgName = "WM_DRAWCLIPBOARD"
  Case WinSubHook.WM_DRAWITEM:             GetMsgName = "WM_DRAWITEM"
  Case WinSubHook.WM_DROPFILES:            GetMsgName = "WM_DROPFILES"
  Case WinSubHook.WM_ENABLE:               GetMsgName = "WM_ENABLE"
  Case WinSubHook.WM_ENDSESSION:           GetMsgName = "WM_ENDSESSION"
  Case WinSubHook.WM_ENTERIDLE:            GetMsgName = "WM_ENTERIDLE"
  Case WinSubHook.WM_ENTERMENULOOP:        GetMsgName = "WM_ENTERMENULOOP"
  Case WinSubHook.WM_ENTERSIZEMOVE:        GetMsgName = "WM_ENTERSIZEMOVE"
  Case WinSubHook.WM_ERASEBKGND:           GetMsgName = "WM_ERASEBKGND"
  Case WinSubHook.WM_EXITMENULOOP:         GetMsgName = "WM_EXITMENULOOP"
  Case WinSubHook.WM_EXITSIZEMOVE:         GetMsgName = "WM_EXITSIZEMOVE"
  Case WinSubHook.WM_FONTCHANGE:           GetMsgName = "WM_FONTCHANGE"
  Case WinSubHook.WM_GETDLGCODE:           GetMsgName = "WM_GETDLGCODE"
  Case WinSubHook.WM_GETFONT:              GetMsgName = "WM_GETFONT"
  Case WinSubHook.WM_GETHOTKEY:            GetMsgName = "WM_GETHOTKEY"
  Case WinSubHook.WM_GETMINMAXINFO:        GetMsgName = "WM_GETMINMAXINFO"
  Case WinSubHook.WM_GETTEXT:              GetMsgName = "WM_GETTEXT"
  Case WinSubHook.WM_GETTEXTLENGTH:        GetMsgName = "WM_GETTEXTLENGTH"
  Case WinSubHook.WM_HOTKEY:               GetMsgName = "WM_HOTKEY"
  Case WinSubHook.WM_HSCROLL:              GetMsgName = "WM_HSCROLL"
  Case WinSubHook.WM_HSCROLLCLIPBOARD:     GetMsgName = "WM_HSCROLLCLIPBOARD"
  Case WinSubHook.WM_ICONERASEBKGND:       GetMsgName = "WM_ICONERASEBKGND"
  Case WinSubHook.WM_IME_CHAR:             GetMsgName = "WM_IME_CHAR"
  Case WinSubHook.WM_IME_COMPOSITION:      GetMsgName = "WM_IME_COMPOSITION"
  Case WinSubHook.WM_IME_COMPOSITIONFULL:  GetMsgName = "WM_IME_COMPOSITIONFULL"
  Case WinSubHook.WM_IME_CONTROL:          GetMsgName = "WM_IME_CONTROL"
  Case WinSubHook.WM_IME_ENDCOMPOSITION:   GetMsgName = "WM_IME_ENDCOMPOSITION"
  Case WinSubHook.WM_IME_KEYDOWN:          GetMsgName = "WM_IME_KEYDOWN"
  Case WinSubHook.WM_IME_KEYLAST:          GetMsgName = "WM_IME_KEYLAST"
  Case WinSubHook.WM_IME_KEYUP:            GetMsgName = "WM_IME_KEYUP"
  Case WinSubHook.WM_IME_NOTIFY:           GetMsgName = "WM_IME_NOTIFY"
  Case WinSubHook.WM_IME_SELECT:           GetMsgName = "WM_IME_SELECT"
  Case WinSubHook.WM_IME_SETCONTEXT:       GetMsgName = "WM_IME_SETCONTEXT"
  Case WinSubHook.WM_IME_STARTCOMPOSITION: GetMsgName = "WM_IME_STARTCOMPOSITION"
  Case WinSubHook.WM_INITDIALOG:           GetMsgName = "WM_INITDIALOG"
  Case WinSubHook.WM_INITMENU:             GetMsgName = "WM_INITMENU"
  Case WinSubHook.WM_INITMENUPOPUP:        GetMsgName = "WM_INITMENUPOPUP"
  Case WinSubHook.WM_KEYDOWN:              GetMsgName = "WM_KEYDOWN"
  Case WinSubHook.WM_KEYFIRST:             GetMsgName = "WM_KEYFIRST"
  Case WinSubHook.WM_KEYLAST:              GetMsgName = "WM_KEYLAST"
  Case WinSubHook.WM_KEYUP:                GetMsgName = "WM_KEYUP"
  Case WinSubHook.WM_KILLFOCUS:            GetMsgName = "WM_KILLFOCUS"
  Case WinSubHook.WM_LBUTTONDBLCLK:        GetMsgName = "WM_LBUTTONDBLCLK"
  Case WinSubHook.WM_LBUTTONDOWN:          GetMsgName = "WM_LBUTTONDOWN"
  Case WinSubHook.WM_LBUTTONUP:            GetMsgName = "WM_LBUTTONUP"
  Case WinSubHook.WM_MBUTTONDBLCLK:        GetMsgName = "WM_MBUTTONDBLCLK"
  Case WinSubHook.WM_MBUTTONDOWN:          GetMsgName = "WM_MBUTTONDOWN"
  Case WinSubHook.WM_MBUTTONUP:            GetMsgName = "WM_MBUTTONUP"
  Case WinSubHook.WM_MDIACTIVATE:          GetMsgName = "WM_MDIACTIVATE"
  Case WinSubHook.WM_MDICASCADE:           GetMsgName = "WM_MDICASCADE"
  Case WinSubHook.WM_MDICREATE:            GetMsgName = "WM_MDICREATE"
  Case WinSubHook.WM_MDIDESTROY:           GetMsgName = "WM_MDIDESTROY"
  Case WinSubHook.WM_MDIGETACTIVE:         GetMsgName = "WM_MDIGETACTIVE"
  Case WinSubHook.WM_MDIICONARRANGE:       GetMsgName = "WM_MDIICONARRANGE"
  Case WinSubHook.WM_MDIMAXIMIZE:          GetMsgName = "WM_MDIMAXIMIZE"
  Case WinSubHook.WM_MDINEXT:              GetMsgName = "WM_MDINEXT"
  Case WinSubHook.WM_MDIREFRESHMENU:       GetMsgName = "WM_MDIREFRESHMENU"
  Case WinSubHook.WM_MDIRESTORE:           GetMsgName = "WM_MDIRESTORE"
  Case WinSubHook.WM_MDISETMENU:           GetMsgName = "WM_MDISETMENU"
  Case WinSubHook.WM_MDITILE:              GetMsgName = "WM_MDITILE"
  Case WinSubHook.WM_MEASUREITEM:          GetMsgName = "WM_MEASUREITEM"
  Case WinSubHook.WM_MENUCHAR:             GetMsgName = "WM_MENUCHAR"
  Case WinSubHook.WM_MENUSELECT:           GetMsgName = "WM_MENUSELECT"
  Case WinSubHook.WM_MOUSEACTIVATE:        GetMsgName = "WM_MOUSEACTIVATE"
  Case WinSubHook.WM_MOUSEMOVE:            GetMsgName = "WM_MOUSEMOVE"
  Case WinSubHook.WM_MOUSEWHEEL:           GetMsgName = "WM_MOUSEWHEEL"
  Case WinSubHook.WM_MOVE:                 GetMsgName = "WM_MOVE"
  Case WinSubHook.WM_MOVING:               GetMsgName = "WM_MOVING"
  Case WinSubHook.WM_NCACTIVATE:           GetMsgName = "WM_NCACTIVATE"
  Case WinSubHook.WM_NCCALCSIZE:           GetMsgName = "WM_NCCALCSIZE"
  Case WinSubHook.WM_NCCREATE:             GetMsgName = "WM_NCCREATE"
  Case WinSubHook.WM_NCDESTROY:            GetMsgName = "WM_NCDESTROY"
  Case WinSubHook.WM_NCHITTEST:            GetMsgName = "WM_NCHITTEST"
  Case WinSubHook.WM_NCLBUTTONDBLCLK:      GetMsgName = "WM_NCLBUTTONDBLCLK"
  Case WinSubHook.WM_NCLBUTTONDOWN:        GetMsgName = "WM_NCLBUTTONDOWN"
  Case WinSubHook.WM_NCLBUTTONUP:          GetMsgName = "WM_NCLBUTTONUP"
  Case WinSubHook.WM_NCMBUTTONDBLCLK:      GetMsgName = "WM_NCMBUTTONDBLCLK"
  Case WinSubHook.WM_NCMBUTTONDOWN:        GetMsgName = "WM_NCMBUTTONDOWN"
  Case WinSubHook.WM_NCMBUTTONUP:          GetMsgName = "WM_NCMBUTTONUP"
  Case WinSubHook.WM_NCMOUSEMOVE:          GetMsgName = "WM_NCMOUSEMOVE"
  Case WinSubHook.WM_NCPAINT:              GetMsgName = "WM_NCPAINT"
  Case WinSubHook.WM_NCRBUTTONDBLCLK:      GetMsgName = "WM_NCRBUTTONDBLCLK"
  Case WinSubHook.WM_NCRBUTTONDOWN:        GetMsgName = "WM_NCRBUTTONDOWN"
  Case WinSubHook.WM_NCRBUTTONUP:          GetMsgName = "WM_NCRBUTTONUP"
  Case WinSubHook.WM_NEXTDLGCTL:           GetMsgName = "WM_NEXTDLGCTL"
  Case WinSubHook.WM_NULL:                 GetMsgName = "WM_NULL"
  Case WinSubHook.WM_PAINT:                GetMsgName = "WM_PAINT"
  Case WinSubHook.WM_PAINTCLIPBOARD:       GetMsgName = "WM_PAINTCLIPBOARD"
  Case WinSubHook.WM_PAINTICON:            GetMsgName = "WM_PAINTICON"
  Case WinSubHook.WM_PALETTECHANGED:       GetMsgName = "WM_PALETTECHANGED"
  Case WinSubHook.WM_PALETTEISCHANGING:    GetMsgName = "WM_PALETTEISCHANGING"
  Case WinSubHook.WM_PARENTNOTIFY:         GetMsgName = "WM_PARENTNOTIFY"
  Case WinSubHook.WM_PASTE:                GetMsgName = "WM_PASTE"
  Case WinSubHook.WM_PENWINFIRST:          GetMsgName = "WM_PENWINFIRST"
  Case WinSubHook.WM_PENWINLAST:           GetMsgName = "WM_PENWINLAST"
  Case WinSubHook.WM_POWER:                GetMsgName = "WM_POWER"
  Case WinSubHook.WM_QUERYDRAGICON:        GetMsgName = "WM_QUERYDRAGICON"
  Case WinSubHook.WM_QUERYENDSESSION:      GetMsgName = "WM_QUERYENDSESSION"
  Case WinSubHook.WM_QUERYNEWPALETTE:      GetMsgName = "WM_QUERYNEWPALETTE"
  Case WinSubHook.WM_QUERYOPEN:            GetMsgName = "WM_QUERYOPEN"
  Case WinSubHook.WM_QUEUESYNC:            GetMsgName = "WM_QUEUESYNC"
  Case WinSubHook.WM_QUIT:                 GetMsgName = "WM_QUIT"
  Case WinSubHook.WM_RBUTTONDBLCLK:        GetMsgName = "WM_RBUTTONDBLCLK"
  Case WinSubHook.WM_RBUTTONDOWN:          GetMsgName = "WM_RBUTTONDOWN"
  Case WinSubHook.WM_RBUTTONUP:            GetMsgName = "WM_RBUTTONUP"
  Case WinSubHook.WM_RENDERALLFORMATS:     GetMsgName = "WM_RENDERALLFORMATS"
  Case WinSubHook.WM_RENDERFORMAT:         GetMsgName = "WM_RENDERFORMAT"
  Case WinSubHook.WM_SETCURSOR:            GetMsgName = "WM_SETCURSOR"
  Case WinSubHook.WM_SETFOCUS:             GetMsgName = "WM_SETFOCUS"
  Case WinSubHook.WM_SETFONT:              GetMsgName = "WM_SETFONT"
  Case WinSubHook.WM_SETHOTKEY:            GetMsgName = "WM_SETHOTKEY"
  Case WinSubHook.WM_SETREDRAW:            GetMsgName = "WM_SETREDRAW"
  Case WinSubHook.WM_SETTEXT:              GetMsgName = "WM_SETTEXT"
  Case WinSubHook.WM_SHOWWINDOW:           GetMsgName = "WM_SHOWWINDOW"
  Case WinSubHook.WM_SIZE:                 GetMsgName = "WM_SIZE"
  Case WinSubHook.WM_SIZING:               GetMsgName = "WM_SIZING"
  Case WinSubHook.WM_SIZECLIPBOARD:        GetMsgName = "WM_SIZECLIPBOARD"
  Case WinSubHook.WM_SPOOLERSTATUS:        GetMsgName = "WM_SPOOLERSTATUS"
  Case WinSubHook.WM_SYSCHAR:              GetMsgName = "WM_SYSCHAR"
  Case WinSubHook.WM_SYSCOLORCHANGE:       GetMsgName = "WM_SYSCOLORCHANGE"
  Case WinSubHook.WM_SYSCOMMAND:           GetMsgName = "WM_SYSCOMMAND"
  Case WinSubHook.WM_SYSDEADCHAR:          GetMsgName = "WM_SYSDEADCHAR"
  Case WinSubHook.WM_SYSKEYDOWN:           GetMsgName = "WM_SYSKEYDOWN"
  Case WinSubHook.WM_SYSKEYUP:             GetMsgName = "WM_SYSKEYUP"
  Case WinSubHook.WM_TIMECHANGE:           GetMsgName = "WM_TIMECHANGE"
  Case WinSubHook.WM_TIMER:                GetMsgName = "WM_TIMER"
  Case WinSubHook.WM_UNDO:                 GetMsgName = "WM_UNDO"
  Case WinSubHook.WM_USER:                 GetMsgName = "WM_USER"
  Case WinSubHook.WM_VKEYTOITEM:           GetMsgName = "WM_VKEYTOITEM"
  Case WinSubHook.WM_VSCROLL:              GetMsgName = "WM_VSCROLL"
  Case WinSubHook.WM_VSCROLLCLIPBOARD:     GetMsgName = "WM_VSCROLL"
  Case WinSubHook.WM_WINDOWPOSCHANGED:     GetMsgName = "WM_WINDOWPOSCHANGED"
  Case WinSubHook.WM_WINDOWPOSCHANGING:    GetMsgName = "WM_WINDOWPOSCHANGING"
  Case WinSubHook.WM_WININICHANGE:         GetMsgName = "WM_WININICHANGE"
  Case Else:                               GetMsgName = fmt(uMsg)
  End Select
End Function
