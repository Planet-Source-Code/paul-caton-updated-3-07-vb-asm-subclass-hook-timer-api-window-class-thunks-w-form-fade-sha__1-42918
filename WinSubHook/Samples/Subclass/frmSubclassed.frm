VERSION 5.00
Begin VB.Form frmSubclassed 
   AutoRedraw      =   -1  'True
   Caption         =   "Subclassed form"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3150
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   3150
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lst 
      Height          =   840
      ItemData        =   "frmSubclassed.frx":0000
      Left            =   180
      List            =   "frmSubclassed.frx":0013
      TabIndex        =   1
      Top             =   180
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Button"
      Height          =   405
      Left            =   1710
      TabIndex        =   0
      Top             =   180
      Width           =   1260
   End
End
Attribute VB_Name = "frmSubclassed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    MsgBox "Close the main form first"
  End If
End Sub
