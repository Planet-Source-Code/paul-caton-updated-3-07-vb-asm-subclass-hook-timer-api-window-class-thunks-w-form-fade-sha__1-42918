VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cHook - module-less, IDE safe, machine code windows hooking thunk
'
'v1.00 20030107 First cut..........................................................................
'

Option Explicit

Private Const OFFSET_P1 As Long = 8                       'Callback gate address
Private Const OFFSET_P2 As Long = 32                      'Owner object address for iHook_Before
Private Const OFFSET_P3 As Long = 93                      'Current hook handle
Private Const OFFSET_P4 As Long = 98                      'CallNextHookEx address
Private Const OFFSET_P5 As Long = 117                     'Owner object address for iHook_After
Private Const OFFSET_P6 As Long = 162                     'Current hook handle
Private Const OFFSET_P7 As Long = 167                     'UnhookWindowsHookEx address
Private Const ARRAY_LB  As Long = 1                       'Low bound of arrays.

Private Type tCode
  Buf(ARRAY_LB To 178)  As Byte                           'Code buffer
End Type

Private Type tCodeBuf
  Code                  As tCode                          'Hook thunk code
End Type

Private CodeBuf         As tCodeBuf                       'Hook thunk code instance
Private nBreakGate      As Long                           'Callback breakpoint gate
Private nCurrentHook    As Long                           'Current hook handle
Private nHookThunk      As Long                           'The address of our hook thunk

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
Const OPS As String = "558BEC83C4F856BE_patch1_33C08945FC8945F8833E00753AC70601000000BA_patch2_8B0283F80074768D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007546FF7510FF750CFF750868_patch3_E8_patch4_8945FC833E00752BC70601000000BA_patch5_8B0283F8007421FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E8B45FCC9C20C0068_patch6_E8_patch7_33C08945FCEBE7"
Dim i     As Long, _
    j     As Long

'Convert the string of opcodes from hex pairs to bytes and store in the code buffer
  With CodeBuf.Code
    j = 1                                                 'Set the character index to the start of the opcode string
    For i = ARRAY_LB To UBound(.Buf)                      'For each byte of the code buffer
      .Buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
      j = j + 2                                           'Bump to the next pair of characters
    Next i                                                'Next byte of the code buffer
    
    nHookThunk = VarPtr(.Buf(ARRAY_LB))                   'Address of the hook thunk entry point
  End With

'Patch the Hook thunk with runtime values
  Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
  Call PatchRel(OFFSET_P4, AddrFunc("CallNextHookEx"))    'Address of the CallNextHookEx api function
  Call PatchRel(OFFSET_P7, _
                        AddrFunc("UnhookWindowsHookEx"))  'Address of the UnhookWindowsHookEx api function
End Sub

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
  If nCurrentHook <> 0 Then
    Call UnHook                                           'UnHook if the hook thunk is active
  End If
End Sub

'-----------------------------
'Public interface

'Call this method to hook app
Public Sub Hook(HookType As WinSubHook.eHookType, Owner As WinSubHook.iHook)
  Debug.Assert (nCurrentHook = 0)
  
  Call PatchVal(OFFSET_P2, ObjPtr(Owner))                 'Owner object address for iHook_Before
  Call PatchVal(OFFSET_P5, ObjPtr(Owner))                 'Owner object address for iHook_After
  nCurrentHook = WinSubHook.SetWindowsHookEx(HookType, nHookThunk, App.hInstance, App.ThreadID)
  Debug.Assert nCurrentHook
  
  Call PatchVal(OFFSET_P3, nCurrentHook)                  'Current hook
  Call PatchVal(OFFSET_P6, nCurrentHook)                  'Current hook
End Sub

'Call this method to unhook
Public Sub UnHook()
  If nCurrentHook <> 0 Then
    Call WinSubHook.UnhookWindowsHookEx(nCurrentHook)
    nCurrentHook = 0                                      'Indicate the hook thunk is inactive
  End If
End Sub

'lParam cast helpers
Public Property Get xCWPSTRUCT(ByVal lParam As Long) As WinSubHook.tCWPSTRUCT
  Call WinSubHook.CopyMemory(xCWPSTRUCT, ByVal lParam, LenB(xCWPSTRUCT))
End Property

Public Property Get xCWPRETSTRUCT(ByVal lParam As Long) As WinSubHook.tCWPRETSTRUCT
  Call WinSubHook.CopyMemory(xCWPRETSTRUCT, ByVal lParam, LenB(xCWPRETSTRUCT))
End Property

Public Property Get xCBT_CREATEWND(ByVal lParam As Long) As WinSubHook.tCBT_CREATEWND
  Call WinSubHook.CopyMemory(xCBT_CREATEWND, ByVal lParam, LenB(xCBT_CREATEWND))
End Property

Public Property Get xCREATESTRUCT(ByVal lParam As Long) As WinSubHook.tCREATESTRUCT
  Call WinSubHook.CopyMemory(xCREATESTRUCT, ByVal lParam, LenB(xCREATESTRUCT))
End Property

Public Property Get xDEBUGSTRUCT(ByVal lParam As Long) As WinSubHook.tDEBUGHOOKINFO
  Call WinSubHook.CopyMemory(xDEBUGSTRUCT, ByVal lParam, LenB(xDEBUGSTRUCT))
End Property

Public Property Get xEVENTMSG(ByVal lParam As Long) As WinSubHook.tEVENTMSG
  Call WinSubHook.CopyMemory(xEVENTMSG, ByVal lParam, LenB(xEVENTMSG))
End Property

Public Property Get xMSG(ByVal lParam As Long) As WinSubHook.tMSG
  Call WinSubHook.CopyMemory(xMSG, ByVal lParam, LenB(xMSG))
End Property

Public Property Get xMOUSEHOOKSTRUCT(ByVal lParam As Long) As WinSubHook.tMOUSEHOOKSTRUCT
  Call WinSubHook.CopyMemory(xMOUSEHOOKSTRUCT, ByVal lParam, LenB(xMOUSEHOOKSTRUCT))
End Property

Public Property Get xRECT(ByVal lParam As Long) As WinSubHook.tRECT
  Call WinSubHook.CopyMemory(xRECT, ByVal lParam, LenB(xRECT))
End Property

'-----------------------------
' Private subroutines

'Return the address of the passed user32.dll api function
Private Function AddrFunc(sProc As String) As Long
  AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)
End Function

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
  Call WinSubHook.CopyMemory(ByVal (nHookThunk + nOffset), nValue, 4)
End Sub

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
  Call WinSubHook.CopyMemory(ByVal (nHookThunk + nOffset), nTargetAddr - nHookThunk - nOffset - 4, 4)
End Sub
