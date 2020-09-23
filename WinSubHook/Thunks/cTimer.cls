VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cTimer - module-less, IDE safe, machine code timer
'
'v1.00 20030107 First cut..........................................................................
'

Option Explicit

Private Const OFFSET_P1 As Long = 5                       'Callback gate address
Private Const OFFSET_P2 As Long = 10                      'Owner object address for iTimer_Fire
Private Const OFFSET_P3 As Long = 36                      'Start time
Private Const OFFSET_P4 As Long = 59                      'Timer ID
Private Const OFFSET_P5 As Long = 66                      'KillTimer address
Private Const ARRAY_LB  As Long = 1                       'Low bound of arrays

Private Type tCode
  Buf(ARRAY_LB To 72)   As Byte                           'Code buffer
End Type

Private Type tCodeBuf
  Code                  As tCode                          'Timer thunk code
End Type

Private CodeBuf         As tCodeBuf                       'Timer thunk code instance
Private nBreakGate      As Long                           'Callback breakpoint gate
Private nTimerID        As Long                           'Timer ID
Private nTimerThunk     As Long                           'The address of our timer thunk

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
Const OPS As String = "558BEC56BE_patch1_BA_patch2_8B0283F8007425833E00751BC706010000008B45142D_patch3_508B0252FF501CC706000000005EC9C2100068_patch4_6A00E8_patch5_EBED"
Dim i     As Long, _
    j     As Long

'Convert the string of opcodes from hex pairs to bytes and store in the code buffer
  With CodeBuf.Code
    j = 1                                                 'Set the character index to the start of the opcode string
    For i = ARRAY_LB To UBound(.Buf)                      'For each byte of the code buffer
      .Buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
      j = j + 2                                           'Bump to the next pair of characters
    Next i                                                'Next byte of the code buffer
    
    nTimerThunk = VarPtr(.Buf(ARRAY_LB))                  'Address of the timer thunk entry point
  End With

'Patch the timer thunk with runtime values
  Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
  Call PatchRel(OFFSET_P5, AddrFunc("KillTimer"))         'Address of the KillTimer api function
End Sub

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
  If nTimerID <> 0 Then
    Call WinSubHook.KillTimer(0, nTimerID)
    nTimerID = 0                                          'Indicate the timer is inactive
  End If
End Sub

'-----------------------------
'Public interface

'Call this method to start the timer
Public Sub Start(nInterval As Long, Owner As WinSubHook.iTimer)
  Debug.Assert (nTimerID = 0)                             'LOGIC ERROR: The timer is active
  
  Call PatchVal(OFFSET_P2, ObjPtr(Owner))                 'Owner object address for iTimer_Fire
  Call PatchVal(OFFSET_P3, WinSubHook.GetTickCount)       'Set the start time

'Create the timer
  nTimerID = WinSubHook.SetTimer(0, 0, nInterval, nTimerThunk)
  Call PatchVal(OFFSET_P4, nTimerID)                      'Timer ID pathched into the thunk so it can kill the timer itself in event of problems
End Sub

'-----------------------------
' Private subroutines

'Return the address of the passed user32.dll api function
Private Function AddrFunc(sProc As String) As Long
  AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)
End Function

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
  Call WinSubHook.CopyMemory(ByVal (nTimerThunk + nOffset), nValue, 4)
End Sub

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
  Call WinSubHook.CopyMemory(ByVal (nTimerThunk + nOffset), nTargetAddr - nTimerThunk - nOffset - 4, 4)
End Sub
