VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subclass"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Courier New"
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
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   717
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   7695
      ScaleHeight     =   6795
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      Begin VB.PictureBox picHolder2 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   495
         ScaleHeight     =   585
         ScaleWidth      =   1680
         TabIndex        =   21
         Top             =   3810
         Width           =   1680
         Begin VB.OptionButton optAfter 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optAfter 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.PictureBox picHold1 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   495
         ScaleHeight     =   600
         ScaleWidth      =   1770
         TabIndex        =   18
         Top             =   465
         Width           =   1770
         Begin VB.OptionButton optBefore 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optBefore 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   368
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkAfter 
         Caption         =   "After original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   16
         Top             =   3465
         Width           =   1935
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "Before original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   15
         Top             =   135
         Width           =   2040
      End
      Begin VB.Frame fraAfter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   180
         TabIndex        =   8
         Top             =   3450
         Width           =   2685
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_MOUSEWHEEL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   510
            TabIndex        =   14
            Top             =   2820
            Width           =   2025
         End
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_LBUTTONDBLCLK"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   510
            TabIndex        =   13
            Top             =   2460
            Width           =   2025
         End
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_LBUTTONUP"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   510
            TabIndex        =   12
            Top             =   2115
            Width           =   2025
         End
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_LBUTTONDOWN"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   510
            TabIndex        =   11
            Top             =   1740
            Width           =   2025
         End
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_MOUSEMOVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   510
            TabIndex        =   10
            Top             =   1410
            Width           =   2025
         End
         Begin VB.CheckBox chkAfterMsg 
            Caption         =   "WM_NCHITTEST"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   510
            TabIndex        =   9
            Top             =   1050
            Width           =   2025
         End
      End
      Begin VB.Frame fraBefore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   180
         TabIndex        =   1
         Top             =   105
         Width           =   2685
         Begin VB.CheckBox chk 
            Height          =   210
            Left            =   2310
            TabIndex        =   24
            Top             =   405
            Width           =   210
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_MOUSEWHEEL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   510
            TabIndex        =   7
            Top             =   2816
            Width           =   2025
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_LBUTTONDBLCLK"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   510
            TabIndex        =   6
            Top             =   2464
            Width           =   2025
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_LBUTTONUP"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   510
            TabIndex        =   5
            Top             =   2112
            Width           =   2025
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_LBUTTONDOWN"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   510
            TabIndex        =   4
            Top             =   1740
            Width           =   2025
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_MOUSEMOVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   510
            TabIndex        =   3
            Top             =   1408
            Width           =   2025
         End
         Begin VB.CheckBox chkBeforeMsg 
            Caption         =   "WM_NCHITTEST"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   510
            TabIndex        =   2
            Top             =   1056
            Width           =   2025
         End
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6840
      Left            =   0
      TabIndex        =   17
      Top             =   15
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   12065
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Messages"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "When"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "uMsg"
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "wParam"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "lParam"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "lReturn"
         Object.Width           =   1852
      EndProperty
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

Private sc As cSubclass           'Subclasser
Implements WinSubHook.iSubclass   'Subclasser interface

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Set sc = New cSubclass
  Call sc.Subclass(frmSubclassed.hWnd, Me)
  Call frmSubclassed.Show(vbModeless, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set sc = Nothing
End Sub

Private Sub chkAfter_Click()
  Dim i As Long
  
  If chkAfter = 0 Then
    optAfter(0).Enabled = False
    optAfter(0).Value = False
    optAfter(1).Enabled = False
    optAfter(1).Value = False
    
    For i = 0 To 5
      chkAfterMsg(i).Enabled = False
      chkAfterMsg(i).Value = vbUnchecked
    Next i
    
    Call sc.DelMsg(ALL_MESSAGES, MSG_AFTER)
  Else
    optAfter(0).Enabled = True
    optAfter(0).Value = False
    optAfter(1).Enabled = True
    optAfter(1).Value = False
  End If
End Sub

Private Sub chkAfterMsg_Click(Index As Integer)
  Dim uMsg As WinSubHook.eMsg
  
  Select Case Index
    Case 0: uMsg = WM_NCHITTEST
    Case 1: uMsg = WM_MOUSEMOVE
    Case 2: uMsg = WM_LBUTTONDOWN
    Case 3: uMsg = WM_LBUTTONUP
    Case 4: uMsg = WM_LBUTTONDBLCLK
    Case 5: uMsg = WM_MOUSEWHEEL
  End Select
    
  If chkAfterMsg(Index).Value = vbUnchecked Then
    Call sc.DelMsg(uMsg, MSG_AFTER)
  Else
    Call sc.AddMsg(uMsg, MSG_AFTER)
  End If
End Sub

Private Sub chkBefore_Click()
  Dim i As Long
  
  If chkBefore = 0 Then
    optBefore(0).Enabled = False
    optBefore(0).Value = False
    optBefore(1).Enabled = False
    optBefore(1).Value = False
    
    For i = 0 To 5
      chkBeforeMsg(i).Enabled = False
      chkBeforeMsg(i).Value = vbUnchecked
    Next i
    
    Call sc.DelMsg(ALL_MESSAGES, MSG_BEFORE)
  Else
    optBefore(0).Enabled = True
    optBefore(0).Value = False
    optBefore(1).Enabled = True
    optBefore(1).Value = False
  End If
End Sub

Private Sub chkBeforeMsg_Click(Index As Integer)
  Dim uMsg As WinSubHook.eMsg
  
  Select Case Index
    Case 0: uMsg = WinSubHook.eMsg.WM_NCHITTEST
    Case 1: uMsg = WinSubHook.eMsg.WM_MOUSEMOVE
    Case 2: uMsg = WinSubHook.eMsg.WM_LBUTTONDOWN
    Case 3: uMsg = WinSubHook.eMsg.WM_LBUTTONUP
    Case 4: uMsg = WinSubHook.eMsg.WM_LBUTTONDBLCLK
    Case 5: uMsg = WinSubHook.eMsg.WM_MOUSEWHEEL
  End Select
    
  If chkBeforeMsg(Index).Value = vbUnchecked Then
    Call sc.DelMsg(uMsg, MSG_BEFORE)
  Else
    Call sc.AddMsg(uMsg, MSG_BEFORE)
  End If
End Sub

Private Sub optAfter_Click(Index As Integer)
  Dim i As Long
  Dim b As Boolean
  
  If Index = 0 Then
    b = False
    Call sc.AddMsg(ALL_MESSAGES, MSG_AFTER)
  Else
    b = True
    Call sc.DelMsg(ALL_MESSAGES, MSG_AFTER)
  End If
  
  For i = 0 To 5
    chkAfterMsg(i).Enabled = b
    chkAfterMsg(i).Value = vbUnchecked
  Next i
End Sub

Private Sub optBefore_Click(Index As Integer)
  Dim i As Long
  Dim b As Boolean
  
  If Index = 0 Then
    b = False
    Call sc.AddMsg(ALL_MESSAGES, MSG_BEFORE)
  Else
    b = True
    Call sc.DelMsg(ALL_MESSAGES, MSG_BEFORE)
  End If
  
  For i = 0 To 5
    chkBeforeMsg(i).Enabled = b
    chkBeforeMsg(i).Value = vbUnchecked
  Next i
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
  Call Display("After ", lReturn, hWnd, uMsg, wParam, lParam)
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
  Call Display("Before", lReturn, hWnd, uMsg, wParam, lParam)
End Sub

Private Sub Display(sWhen As String, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
  Static nMsgs  As Long
  Dim itm       As MSComctlLib.ListItem
  
  nMsgs = nMsgs + 1
  Set itm = lv.ListItems.Add(, , fmt(nMsgs))
  
  With itm
    .SubItems(1) = sWhen
    .SubItems(2) = GetMsgName(uMsg)
    .SubItems(3) = fmt(wParam)
    .SubItems(4) = fmt(lParam)
    .SubItems(5) = fmt(lReturn)
    .EnsureVisible
  End With
End Sub

'Return the Value parameter converted to a hex string padded to 8 characters with a leading &H
Private Function fmt(Value As Long) As String
  Dim s As String
  
  s = Hex$(Value)
  fmt = String$(8 - Len(s), "0") & s
End Function

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
