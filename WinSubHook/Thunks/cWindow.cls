VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cWindow - module-less, IDE safe machine code WndProc for API windows
'
'v1.00 20030107 First cut..........................................................................
'

Option Explicit

Private Const OFFSET_P1 As Long = 9                       'Callback gate address
Private Const OFFSET_P2 As Long = 22                      'Owner object address for iWindow_WndProc
Private Const OFFSET_P3 As Long = 34                      'Before table entry count
Private Const OFFSET_P4 As Long = 49                      'Before table address
Private Const OFFSET_P5 As Long = 63                      'IDE check
Private Const OFFSET_P6 As Long = 137                     'DefWindowProc address
Private Const OFFSET_P7 As Long = 157                     'DestroyWindow address
Private Const ARRAY_LB  As Long = 1                       'Lowest bound of arrays

Private Type tCode
  Buf(ARRAY_LB To 168)  As Byte                           'Code buffer
End Type

Private Type tCodeBuf
  Code                  As tCode                          'Subclass WndProc code
End Type

Private CodeBuf         As tCodeBuf                       'WndProc code instance
Private nBreakGate      As Long                           'Callback breakpoint gate
Private nMsgCnt         As Long                           'Msg table entry count
Private aMsgTbl()       As WinSubHook.eMsg                'Msg table array
Private nWndProc        As Long                           'The address of the WndProc
Private col_hWnds       As Collection                     'Collection of window handles
Private m_sClass        As String                         'Class name
Private m_Owner         As iWindow                        'Private member property variable

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
Const OPS As String = "558BEC83C4F85756BE_patch1_33C08945FC8945F8BA_patch2_8B0283F8007478B90000000083F900745183F9FF740CBF000000008B450CF2AF754033C03D_patch5_740B833E007532C706010000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF501CC706000000008B45F883F8007514FF7514FF7510FF750CFF7508E8_patch6_8945FC5E5F8B45FCC9C21000FF7508E8_patch7_33C08945FCEBE8"
Dim i     As Long, _
    j     As Long

'Convert the string of opcodes from hex pairs to bytes and store in the code buffer
  With CodeBuf.Code
    j = 1                                                 'Set the character index to the start of the opcode string
    For i = ARRAY_LB To UBound(.Buf)                      'For each byte of the code buufer
      .Buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
      j = j + 2                                           'Bump to the next pair of characters
    Next i                                                'Next byte of the code buffer
    
    nWndProc = VarPtr(.Buf(ARRAY_LB))                     'Address of the cWindow WndProc entry point
  End With
      
'Patch the WndProc thunk with runtime values
  Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
  Call PatchVal(OFFSET_P5, InIDE)                         'Whether we need check the breakpoint gate and the vtable
  Call PatchRel(OFFSET_P6, AddrFunc("DefWindowProcA"))    'Address of the DefWindowProc api function
  Call PatchRel(OFFSET_P7, AddrFunc("DestroyWindow"))     'Address of the DestroyWindow api function
  
  Set col_hWnds = New Collection                          'Create instance of window handles collection
End Sub

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
  Dim i As Long
  
  Call PatchVal(OFFSET_P3, 0)                             'Patch the code to ensure no further iWindow_WndProc callbacks
  
  For i = col_hWnds.Count To 1 Step -1                    'For each window created (and not yet destroyed)
    Call WinSubHook.DestroyWindow(col_hWnds.Item(i))      'Destroy the window
    Call col_hWnds.Remove(i)                              'Remove from the collection
  Next i                                                  'Next window
  Set col_hWnds = Nothing                                 'Destroy the collection
  
  If Len(m_sClass) > 0 Then                               'If a class was registered
    Call UnregisterClass(m_sClass, App.hInstance)         'Unregister the window class
  End If
End Sub

'-----------------------------
'Public interface

'Call this method to add a message to the msg callback table. NB This method can be called at any time
Public Sub AddMsg(uMsg As WinSubHook.eMsg)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                             'If ALL_MESSAGES
    nMsgCnt = -1                                          'Indicates that all messages are to callback
  Else                                                    'Else a specific message number
    For nEntry = ARRAY_LB To nMsgCnt                      'For each existing entry. NB will skip if 0 or -1 (ALL_MESSAGES)
      Select Case aMsgTbl(nEntry)                         'Select on the message number stored in this table entry
      Case -1                                             'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                            'Re-use this entry
        Exit Sub                                          'Bail
      Case uMsg                                           'The msg is already in the table!
        Exit Sub                                          'Bail
      End Select
    Next nEntry                                           'Next entry
    
'Make space for the new entry
    ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)            'Increase the size of the table. NB nEntry = nMsgCnt + 1
    nMsgCnt = nEntry                                      'Bump the entry count
    aMsgTbl(nEntry) = uMsg                                'Store the message in the table
  End If
  
  Call PatchVal(OFFSET_P3, nMsgCnt)                       'Patch the Before table entry count
  Call PatchVal(OFFSET_P4, AddrMsgTbl())                  'Patch the address of the Before message table. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.
End Sub

'Arbitarily call the DefWindowProc - Normally, if bHandled isn't set in the WndProc callback, the DefWindowProc
'is called in the assembler thunk after the callback. Use this method to call the DefWindowProc first.
Public Function CallDefWndProc(hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam) As Long
  CallDefWndProc = WinSubHook.DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Return the window class name
Public Property Get Class() As String
  Class = m_sClass
End Property

'Call this method to delete a message from the msg table. NB This method can be called at any time
Public Sub DelMsg(uMsg As WinSubHook.eMsg)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                             'If deleting all messages (specific or ALL_MESSAGES)
    nMsgCnt = 0                                           'Message count is now zero
    Call PatchVal(OFFSET_P3, 0)                           'Patch the before table message count
  Else                                                    'Else deleteting a specific message
    For nEntry = ARRAY_LB To nMsgCnt                      'For each table entry
      If aMsgTbl(nEntry) = uMsg Then                      'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = -1                              'Mark the table slot as available
        Exit For                                          'Bail
      End If
    Next nEntry                                           'Next entry
  End If
End Sub

'Set the window class owner, Form/Class/UserControl
Public Property Set Owner(NewOwner As WinSubHook.iWindow)
  Set m_Owner = NewOwner
  Call PatchVal(OFFSET_P2, ObjPtr(m_Owner))               'Owner object address for iWindow_WndProc
End Property

'Register the window class, call this before creating windows--unless one of the predefined window classes is required.
Public Function WindowClassRegister(sClass As String, _
                                    Optional colBackground As Long = &HFFFFFF, _
                                    Optional Style As WinSubHook.eClassStyle = 0, _
                                    Optional hCursor As Long = 0, _
                                    Optional hIcon As Long = 0, _
                                    Optional hIconSm As Long = 0, _
                                    Optional cbClassExtra As Long = 0, _
                                    Optional cbWndExtra As Long = 0) As Boolean
  Dim wc As tWNDCLASSEX

  Debug.Assert (m_sClass = vbNullString)                  'This method should only be called once or never for a predefined class
  
  m_sClass = sClass                                       'Store the class name

  With wc
    .cbSize = Len(wc)                                     'Size of the window class type
    .cbClsExtra = cbClassExtra                            'Number of class extra bytes
    .cbWndExtra = cbWndExtra                              'Number of window extra bytes
    .hbrBackground = CreateSolidBrush(colBackground)      'Class background
    .hCursor = hCursor                                    'Class cursor
    .hIcon = hIcon                                        'Class icon
    .hIconSm = hIconSm                                    'Class small icon
    .hInstance = App.hInstance                            'Application instance handle
    .lpfnWndProc = nWndProc                               'Class WndProc address
    .Style = Style                                        'Class style
    .lpszClassName = StrPtr( _
                      StrConv(m_sClass, vbFromUnicode))   'Class name
  End With

  WindowClassRegister = (RegisterClassEx(wc) <> 0)        'Register the window class
End Function

'Create a window, return the window handle
Public Function WindowCreate(dwExStyle As WinSubHook.eWindowStyleEx, _
                             dwStyle As WinSubHook.eWindowStyle, _
                             Optional Class As WinSubHook.eWindowClass = AS_WINDOWCLASS, _
                             Optional x As Long = 0, _
                             Optional y As Long = 0, _
                             Optional nWidth As Long = 0, _
                             Optional nHeight As Long = 0, _
                             Optional sCaption As String = "", _
                             Optional hWndParent As Long = 0, _
                             Optional hMenu As Long = 0, _
                             Optional lParam As Long = 0) As Long
Dim hWnd    As Long, _
    sClass  As String
  
  Debug.Assert (Not (m_Owner Is Nothing))                 'LOGIC ERROR! the Owner must be set before calling this method
  
  Select Case Class
'User defined window class
    Case WinSubHook.eWindowClass.AS_WINDOWCLASS:             sClass = m_sClass
    
'Predefined window classes
    Case WinSubHook.eWindowClass.PREDEFINED_BUTTON:          sClass = "BUTTON"
    Case WinSubHook.eWindowClass.PREDEFINED_COMBOBOX:        sClass = "COMBOBOX"
    Case WinSubHook.eWindowClass.PREDEFINED_EDIT:            sClass = "EDIT"
    Case WinSubHook.eWindowClass.PREDEFINED_LISTBOX:         sClass = "LISTBOX"
    Case WinSubHook.eWindowClass.PREDEFINED_MDICLIENT:       sClass = "MDICLIENT"
    Case WinSubHook.eWindowClass.PREDEFINED_RICHEDIT:        sClass = "RichEdit"
    Case WinSubHook.eWindowClass.PREDEFINED_RICHEDIT_CLASS:  sClass = "RICHEDIT_CLASS"
    Case WinSubHook.eWindowClass.PREDEFINED_SCROLLBAR:       sClass = "SCROLLBAR"
    Case WinSubHook.eWindowClass.PREDEFINED_STATIC:          sClass = "STATIC"
  End Select
  Debug.Assert (sClass <> vbNullString)                   'LOGIC ERROR! Class name not defined
  
'Create the window
  hWnd = WinSubHook.CreateWindowEx(dwExStyle, _
                                    sClass, _
                                    sCaption, _
                                    dwStyle, _
                                    x, y, nWidth, nHeight, _
                                    hWndParent, _
                                    hMenu, _
                                    App.hInstance, _
                                    lParam)
  Debug.Assert hWnd                                       'CreateWindow failed
  Call col_hWnds.Add(hWnd, "h" & hWnd)                    'Add the window handle to the collection
  WindowCreate = hWnd
End Function

'Destroy window
Public Function WindowDestroy(ByVal hWnd As Long) As Boolean
Dim sKey As String
  
  On Error GoTo Catch
    sKey = "h" & hWnd
    hWnd = col_hWnds.Item(sKey)                           'Ensure the handle is in the collection
    Call WinSubHook.DestroyWindow(hWnd)                   'Destroy the window
    Call col_hWnds.Remove(sKey)                           'Remove the handle from the collection
    WindowDestroy = True
Catch:
  On Error GoTo 0
End Function

'-----------------------------
' Private subroutines

'Return the address of the passed user32.dll api function
Private Function AddrFunc(sProc As String) As Long
  AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)
End Function

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl() As Long
  On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
    AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))                'Get the address of the first element of the passed message table
  On Error GoTo 0                                         'Switch off error protection
End Function

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
  Call WinSubHook.CopyMemory(ByVal (nWndProc + nOffset), nValue, 4)
End Sub

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
  Call WinSubHook.CopyMemory(ByVal (nWndProc + nOffset), nTargetAddr - nWndProc - nOffset - 4, 4)
End Sub

'Return -1 if we're running in the IDE or 0 if were running compiled
Private Function InIDE() As Long
  Static Value As Long
  
  If Value = 0 Then
    Value = 1
    Debug.Assert InIDE() Or True                          'This line won't exist in the compiled app
    InIDE = Value - 1
  End If
  
  Value = 0
End Function
