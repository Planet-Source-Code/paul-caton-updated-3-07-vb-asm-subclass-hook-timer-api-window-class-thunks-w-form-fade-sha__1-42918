VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Shadow"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4065
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
   ScaleHeight     =   4530
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Shadow.ucShadow Shadow 
      Left            =   2550
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      FadeTime        =   1000
      Transparency    =   128
   End
   Begin VB.Frame fraFade 
      Caption         =   "Fade"
      Height          =   1440
      Left            =   150
      TabIndex        =   13
      Top             =   2250
      Width           =   3765
      Begin VB.CheckBox chkFadeIn 
         Caption         =   "Fade In"
         Height          =   195
         Left            =   2220
         TabIndex        =   17
         Top             =   435
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSComCtl2.UpDown updFadeIn 
         Height          =   330
         Left            =   1620
         TabIndex        =   16
         Top             =   360
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         _Version        =   393216
         Value           =   100
         Increment       =   100
         Max             =   2000
         Min             =   100
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkFadeOut 
         Caption         =   "Fade Out"
         Height          =   195
         Left            =   2220
         TabIndex        =   14
         Top             =   975
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComCtl2.UpDown updFadeOut 
         Height          =   330
         Left            =   1620
         TabIndex        =   15
         Top             =   900
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         _Version        =   393216
         Value           =   200
         Increment       =   100
         Max             =   2000
         Min             =   100
         Enabled         =   -1  'True
      End
      Begin VB.Label lblFadeIn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fade In:"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   405
         Width           =   1485
      End
      Begin VB.Label lblFadeOut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fade Out:"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   945
         Width           =   1485
      End
   End
   Begin VB.Frame fraShadow 
      Caption         =   "Shadow"
      Height          =   1965
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   3765
      Begin VB.CheckBox chkVisible 
         Caption         =   "Shadow Visible"
         Height          =   195
         Left            =   2220
         TabIndex        =   9
         Top             =   435
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   330
         Left            =   1620
         TabIndex        =   6
         Top             =   1455
         Width           =   345
      End
      Begin VB.CheckBox chkSoft 
         Caption         =   "Soft Shadow"
         Height          =   195
         Left            =   2220
         TabIndex        =   5
         Top             =   805
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkHideS 
         Caption         =   "Hide on size"
         Height          =   195
         Left            =   2220
         TabIndex        =   4
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox chkHideM 
         Caption         =   "Hide on move"
         Height          =   195
         Left            =   2220
         TabIndex        =   3
         Top             =   1175
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin MSComCtl2.UpDown updDepth 
         Height          =   330
         Left            =   1620
         TabIndex        =   7
         Top             =   915
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         _Version        =   393216
         Value           =   8
         Max             =   32
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTransparency 
         Height          =   330
         Left            =   1635
         TabIndex        =   8
         Top             =   360
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         _Version        =   393216
         Value           =   80
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTransparency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Transparency: "
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   405
         Width           =   1485
      End
      Begin VB.Label lblDepth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Depth:"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   945
         Width           =   1485
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdChild 
      Caption         =   "New Form"
      Height          =   360
      Left            =   1290
      TabIndex        =   0
      Top             =   4050
      Width           =   1485
   End
   Begin VB.Label lbl 
      Height          =   280
      Left            =   1875
      TabIndex        =   1
      Top             =   2010
      Width           =   1890
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

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  With Shadow
    chkSoft.Value = IIf(.SoftShadow, vbChecked, vbUnchecked)
    chkHideM.Value = IIf(.HideMove, vbChecked, vbUnchecked)
    chkHideS.Value = IIf(.HideSize, vbChecked, vbUnchecked)
    chkVisible.Value = IIf(.Visible, vbChecked, vbUnchecked)
    chkFadeIn.Value = IIf(.FadeIn, vbChecked, vbUnchecked)
    chkFadeOut.Value = vbChecked
    lblColor.BackColor = .Color
    updDepth.Value = .Depth
    updTransparency.Value = .Transparency
    updFadeIn.Value = .FadeTime
    updFadeOut.Value = .FadeTime
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Me.chkFadeOut = vbChecked Then
    Call Shadow.FadeOut(updFadeOut.Value)
  End If
End Sub

Private Sub chkFadeIn_Click()
  Shadow.FadeIn = (chkFadeIn.Value = vbChecked)
End Sub

Private Sub chkHideM_Click()
  Shadow.HideMove = (chkHideM.Value = vbChecked)
End Sub

Private Sub chkHideS_Click()
  Shadow.HideSize = (chkHideS.Value = vbChecked)
End Sub

Private Sub chkSoft_Click()
  Shadow.SoftShadow = (chkSoft.Value = vbChecked)
End Sub

Private Sub chkVisible_Click()
  Shadow.Visible = (chkVisible.Value = vbChecked)
End Sub

Private Sub cmdChild_Click()
  Dim frm As frmMain
  
  Set frm = New frmMain
  Call Load(frm)
  
  With frm
    .lblColor.BackColor = lblColor.BackColor
    .Shadow.Color = Shadow.Color
    .updDepth = updDepth
    .chkHideM = chkHideM
    .chkHideS = chkHideS
    .chkSoft = chkSoft
    .updTransparency = updTransparency
    .chkVisible = chkVisible
  
    .chkFadeIn = chkFadeIn
    .updFadeIn = updFadeIn.Value
    .chkFadeOut = chkFadeOut
    .updFadeOut = updFadeOut
  End With
    
  Call frm.Show(vbModeless)
End Sub

Private Sub cmdColor_Click()
  On Error GoTo Catch
  
  With dlgColor
    .CancelError = True
    .Color = lblColor.BackColor
    .DialogTitle = "Shadow Color"
    .ShowColor
    Shadow.Color = .Color
    lblColor.BackColor = .Color
  End With
Catch:
End Sub

Private Sub updDepth_Change()
  Shadow.Depth = updDepth.Value
  lblDepth = " Depth: " & updDepth.Value
End Sub

Private Sub updFadeIn_Change()
  Shadow.FadeTime = updFadeIn.Value
  lblFadeIn = " FadeIn: " & updFadeIn.Value
End Sub

Private Sub updFadeOut_Change()
  lblFadeOut = " FadeOut: " & updFadeOut.Value
End Sub

Private Sub updTransparency_Change()
  Shadow.Transparency = updTransparency.Value
  lblTransparency = " Transparency: " & updTransparency.Value
End Sub
