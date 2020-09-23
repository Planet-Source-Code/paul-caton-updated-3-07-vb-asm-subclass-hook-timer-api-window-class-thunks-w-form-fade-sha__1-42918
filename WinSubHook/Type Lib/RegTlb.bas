Attribute VB_Name = "mRegTlb"

'Handy little Type Library registration tool
'
Option Explicit

Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Main()
  Dim sTypeLib  As String
  Dim o_TypeLib As TypeLibInfo
  
  On Error GoTo Catch_22
  
  sTypeLib = Show
  
  If sTypeLib <> "" Then
    Set o_TypeLib = TLIApplication.TypeLibInfoFromFile(sTypeLib)
    
    If MsgBox("Do you want to register the type library?" & vbNewLine & vbNewLine & _
              "File..." & vbNewLine & _
              "  " & sTypeLib & vbNewLine & vbNewLine & _
              "HelpString..." & vbNewLine & _
              " '" & o_TypeLib.HelpString & "'" & vbNewLine, _
              vbQuestion Or vbYesNo, _
              "Register Type Library?") = vbYes Then
              
      o_TypeLib.Register
      MsgBox "Type Library Registered", vbInformation
    End If
  End If
  
GetOuttaHere:
  Set o_TypeLib = Nothing
  Exit Sub
  
Catch_22:
  MsgBox Err.Description, vbExclamation
  Resume GetOuttaHere
End Sub

Private Function Show() As String
  Dim of As OPENFILENAME
  
  With of
    .Flags = &H281004
    .lpstrDefExt = "tlb"
    .lpstrFile = String$(260, 0)
    .lpstrFilter = "Type Libraries" & vbNullChar & "*.tlb" & vbNullChar & vbNullChar
    .lpstrTitle = "Select Type Library"
    .nFilterIndex = 1
    .nMaxFile = 260
    .nMaxFileTitle = 260
    .lStructSize = LenB(of)
  End With
  
  If GetOpenFileName(of) Then
    Show = TrimNulls(of.lpstrFile)
  End If
End Function

Private Function TrimNulls(ByVal StrIn As String) As String
  Dim nPos As Long

  nPos = InStr(StrIn, vbNullChar)

  If nPos = 0 Then
    TrimNulls = StrIn
  Else
    If nPos = 1 Then
      TrimNulls = ""
    Else
      TrimNulls = Left$(StrIn, nPos - 1)
    End If
  End If
End Function
