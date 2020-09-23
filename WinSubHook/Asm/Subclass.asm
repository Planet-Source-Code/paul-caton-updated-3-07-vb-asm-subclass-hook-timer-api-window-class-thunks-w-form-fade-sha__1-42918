;##########################################################################################################################
;#
;# cSubclass,cls model assembler source

.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Thunk.inc                       ;# Macros 'n stuff

GWL_WNDPROC = -4
_patch1_    = 01BCCAABh                 ;# Callback break gate address
_patch2_    = 02BCCAABh                 ;# Table B (before) entry count
_patch3_    = 03BCCAABh                 ;# Table B (before) address
_patch4_    = 04BCCAABh                 ;# In IDE?
_patch5_    = 05BCCAABh                 ;# Object address of the owner Form/UserControl/Class
_patch6_    = 06BCCAABh                 ;# Address of the previous WndProc
                                        ;# _patch7_ call Dummy, patched to CallWindowProc
_patch8_    = 08BCCAABh                 ;# Table A (after) entry count
_patch9_    = 09BCCAABh                 ;# Table A (after) address
_patchA_    = 0ABCCAABh                 ;# In IDE?
_patchB_    = 0BBCCAABh                 ;# Object address of the owner Form/UserControl/Class
_patchC_    = 0CBCCAABh                 ;# Address of the previous WndProc
                                        ;# _patchD_ call Dummy, patched to SetWindowLong
.code

start:

WndProc proc    hWin    :DWORD,         ;# Window handle
                uMsg    :DWORD,         ;# Message number
                wParam  :DWORD,         ;# Message related parameter
                lParam  :DWORD          ;# Message related parameter
                
    LOCAL   lReturn     :DWORD          ;# Value returned to windows
    LOCAL   bHandled    :DWORD          ;# If set in the iSubclass_Before callback then the original WndProc is skipped
_init:    
    push    edi                         ;# Preserve edi
    push    esi                         ;# Preserve esi
    mov     esi, _patch1_               ;# Callback break gate address, patched at runtime
    xor     eax, eax                    ;# Clear eax
    mov     lReturn, eax                ;# Clear lReturn
    mov     bHandled, eax               ;# Clear bHandled
_check_before:
    mov     ecx, _patch2_               ;# Table B entry count, patched at runtime
    cmp     ecx, 0                      ;# If no entries
    je      _call_original              ;# Call original WndProc
    cmp     ecx, 0FFFFFFFFh             ;# If entries = -1
    je      _check_ide_before           ;# All messages call iSubCls_Before
    mov     edi, _patch3_               ;# Table B address, patched at runtime
    mov     eax, uMsg                   ;# Message number to search for
    repne   scasd                       ;# Scan the table
    jne     _call_original              ;# If not found call the original WndProc
_check_ide_before:
    xor     eax, eax                    ;# Clear eax
    cmp     eax, _patch4_               ;# IDE check, patched at runtime
    je      _validate_vtbl_before       ;# Skip gate check
_check_callback_gate_before:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _call_original              ;# If the callback gate is set call the original WndProc
    mov     dword ptr [esi], 1          ;# Close the callback gate
_validate_vtbl_before:
    mov     edx, _patch5_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0                      ;# Validate Vtbl
    je      _un_subclass                ;# Vtbl invalid
_callback_before:
    lea     eax, lParam                 ;# ByRef lParam 
    push    eax                         ;#
    lea     eax, wParam                 ;# ByRef wParam
    push    eax                         ;#
    lea     eax, uMsg                   ;# ByRef uMsg 
    push    eax                         ;#
    lea     eax, hWin                   ;# ByRef hWin 
    push    eax                         ;#
    lea     eax, lReturn                ;# ByRef lReturn 
    push    eax                         ;#
    lea     eax, bHandled               ;# ByRef bHandled 
    push    eax                         ;#
    mov     eax, [edx]                  ;# Vtbl
    push    edx                         ;#
    call    dword ptr [eax][20h]        ;# Call iSubclass_Before
    mov     dword ptr [esi], 0          ;# Open the callback gate
    mov     eax, bHandled               ;#
    cmp     eax, 0                      ;# If iSubclass_Before has handled the call
    jne      _return                    ;# Don't call the original WndProc
_call_original:
    push    lParam                      ;# ByVal lParam
    push    wParam                      ;# ByVal wParam
    push    uMsg                        ;# ByVal uMsg
    push    hWin                        ;# ByVal hWin
    push    _patch6_                    ;# Address of the previous WndProc, patched at runtime
    call    Dummy                       ;# Call CallWindowProc, patched at runtime
    mov     lReturn, eax                ;# Preserve the return value
    mov     ecx, _patch8_               ;# Table A entries, patched at runtime
    cmp     ecx, 0                      ;# If no entries
    je      _return                     ;# Call return
    cmp     ecx, 0FFFFFFFFh             ;# If entries = -1
    je      _check_ide_after            ;# All messages call iSubclass_After
    mov     edi, _patch9_               ;# Table A address, patched at runtime
    mov     eax, uMsg                   ;# Message number to search for
    repne   scasd                       ;# Scan the table
    jne     _return                     ;# If not found then return
_check_ide_after:
    xor     eax, eax                    ;# Clear eax
    cmp     eax, _patchA_               ;# IDE check, patched at runtime
    je      _validate_vtbl_after        ;# Skip gate check
_check_callback_gate_after:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _return                     ;# If the callback gate is set return
    mov     dword ptr [esi], 1;         ;# Close the callback gate
_validate_vtbl_after:
    mov     edx, _patchB_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0
    je      _un_subclass
_callback_after:    
    push    lParam                      ;# ByVal lParam
    push    wParam                      ;# ByVal wParam
    push    uMsg                        ;# ByVal uMsg
    push    hWin                        ;# ByVal hWin
    lea     eax, lReturn                ;# ByRef lReturn
    push    eax                         ;#
    mov     eax, [edx]                  ;# Vtbl
    push    edx                         ;#    
    call    dword ptr [eax][1Ch]        ;# Call iSubclass_After
    mov     dword ptr [esi], 0          ;# Open the callback gate
_return:                                ;#
    pop     esi                         ;#
    pop     edi                         ;# Restore edi
    mov     eax, lReturn                ;# Function return value
    ret
_un_subclass:
    push    _patchC_                    ;# Address of the previous WndProc
    push    GWL_WNDPROC                 ;# WndProc index
    push    hWin                        ;# Push the window handle
    call    Dummy                       ;# Call SetWindowsLong, patched at runtime
    xor     eax, eax
    mov     lReturn, eax
    jmp     _return

WndProc endp
Dummy:
end start