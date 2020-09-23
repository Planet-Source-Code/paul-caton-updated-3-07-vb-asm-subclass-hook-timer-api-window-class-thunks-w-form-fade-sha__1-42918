VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WH_MSGFILTER"
   ClientHeight    =   1725
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar vsb 
      Height          =   1740
      Left            =   3090
      TabIndex        =   1
      Top             =   -15
      Width           =   330
   End
   Begin VB.CommandButton cmdMsgBox 
      Caption         =   "MsgBox"
      Height          =   405
      Left            =   555
      TabIndex        =   0
      Top             =   960
      Width           =   1890
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItem 
         Caption         =   "Item &1"
         Index           =   0
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Item &2"
         Index           =   1
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuItem 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMsgBox_Click()
  MsgBox "MsgBox", vbInformation
End Sub

Private Sub mnuItem_Click(Index As Integer)
  If Index = 3 Then Unload Me
End Sub
