;##########################################################################################################################
;#
;# cWindow,cls model assembler source

.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Thunk.inc                       ;# Macros 'n stuff

_patch1_    = 01BCCAABh                 ;# Callback break gate address
_patch2_    = 02BCCAABh                 ;# Object address of the owner Form/UserControl/Class
_patch3_    = 03BCCAABh                 ;# Table B (before) entry count
_patch4_    = 04BCCAABh                 ;# Table B (before) address
_patch5_    = 05BCCAABh                 ;# In IDE?
                                        ;# _patch6_ call Dummy, patched to DefWindowProc
                                        ;# _patch7_ call Dummy, patched to DestroyWindow
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
    mov     esi, _patch1_               ;# Break gate address, patched at runtime
    xor     eax, eax                    ;# Clear eax
    mov     lReturn, eax                ;# Clear lReturn
    mov     bHandled, eax               ;# Clear bHandled
_validate_vtbl:
    mov     edx, _patch2_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0                      ;# Validate Vtbl
    je      _destroy_window             ;# Vtbl invalid
_check_message:
    mov     ecx, _patch3_               ;# Table entry count, patched at runtime
    cmp     ecx, 0                      ;# If no entries
    je      _def_window_proc            ;# Call original default processing
    cmp     ecx, 0FFFFFFFFh             ;# If entries = -1
    je      _check_ide                  ;# All messages call iWindow_WndProc
    mov     edi, _patch4_               ;# Table address, patched at runtime
    mov     eax, uMsg                   ;# Message number to search for
    repne   scasd                       ;# Scan the table
    jne     _def_window_proc            ;# Call default processing
_check_ide:
    xor     eax, eax                    ;# Clear eax
    cmp     eax, _patch5_               ;# IDE check, patched at runtime
    je      _no_ide                     ;# Skip gate check
_check_callback_gate:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _def_window_proc            ;# If the callback gate is set call the original WndProc
    mov     dword ptr [esi], 1          ;# Close the callback gate
_no_ide:
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
    call    dword ptr [eax][1Ch]        ;# Call iWindow_WndProc
    mov     dword ptr [esi], 0          ;# Open the callback gate
    mov     eax, bHandled               ;#
    cmp     eax, 0                      ;# If iWindow_WndProc has handled the call
    jne      _return                    ;# Return
_def_window_proc:
    push    lParam                      ;# ByVal lParam
    push    wParam                      ;# ByVal wParam
    push    uMsg                        ;# ByVal uMsg
    push    hWin                        ;# ByVal hWin
    call    Dummy                       ;# Call DefWindowProc, patched at runtime
    mov     lReturn, eax                ;# Preserve the return value
_return:                                ;#
    pop     esi                         ;#
    pop     edi                         ;# Restore edi
    mov     eax, lReturn                ;# Function return value
    ret
_destroy_window:
    push    hWin                        ;# Push the window handle
    call    Dummy                       ;# Call DestroyWindow, patched at runtime
    xor     eax, eax
    mov     lReturn, eax
    jmp     _return
WndProc endp
Dummy:
end start
