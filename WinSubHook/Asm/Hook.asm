;##########################################################################################################################
;#
;# cHook,cls model assembler source

.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Thunk.inc                       ;# Macros 'n stuff

_patch1_    = 01BCCAABh                 ;# Callback break gate address
_patch2_    = 02BCCAABh                 ;# Object address of the owner Form/UserControl/Class for vtbl validation and before callback
_patch3_    = 03BCCAABh                 ;# Current hook handle for CallNextHookEx
                                        ;# _patch4_ Call CallNextHookEx
_patch5_    = 05BCCAABh                 ;# Object address of the owner Form/UserControl/Class for vtbl validation and after callback
_patch6_    = 06BCCAABh                 ;# Current hook handle for UnhookWindowsHookEx
                                        ;# _patch7_ Call UnhookWindowsHookEx
.code

start:

HookProc proc   nCode   :DWORD,         ;# Hook code
                wParam  :DWORD,         ;# Message related parameter
                lParam  :DWORD          ;# Message related parameter
                
    LOCAL   lReturn     :DWORD          ;# Hook return value
    LOCAL   bHandled    :DWORD          ;# If set in the iHook_Before callback then the original WndProc is skipped
_init:    
    push    esi                         ;# Preserve esi
    mov     esi, _patch1_               ;# Callback break gate address, patched at runtime
    xor     eax, eax                    ;# Clear eax
    mov     lReturn, eax                ;# Clear lReturn
    mov     bHandled, eax               ;# Clear bHandled
_check_callback_gate_before:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _call_next_hook             ;# If the callback gate is set call the next hook
    mov     dword ptr [esi], 1          ;# Close the callback gate
_validate_vtbl_before:
    mov     edx, _patch2_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0                      ;# Validate Vtbl
    je      _unhook                     ;# Vtbl invalid
_callback_before:
    lea     eax, lParam                 ;# ByRef lParam 
    push    eax                         ;#
    lea     eax, wParam                 ;# ByRef wParam
    push    eax                         ;#
    lea     eax, nCode                  ;# ByRef nCode
    push    eax                         ;#
    lea     eax, lReturn                ;# ByRef lReturn 
    push    eax                         ;#
    lea     eax, bHandled               ;# ByRef bHandled 
    push    eax                         ;#
    mov     eax, [edx]                  ;# Vtbl
    push    edx                         ;#
    call    dword ptr [eax][20h]        ;# Call iHook_Before
    mov     dword ptr [esi], 0          ;# Open the callback gate
    mov     eax, bHandled               ;#
    cmp     eax, 0                      ;# If iHook_Before has handled the call
    jne      _return                    ;# Don't call the next hook
_call_next_hook:
    push    lParam                      ;# ByVal lParam
    push    wParam                      ;# ByVal wParam
    push    nCode                       ;# ByVal nCode
    push    _patch3_                    ;# Current hook handle
    call    Dummy                       ;# Call CallNextHookEx, patched at runtime (_patch4_)
    mov     lReturn, eax                ;# Preserve the return value
_check_callback_gate_after:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _return                     ;# If the callback gate is set return
    mov     dword ptr [esi], 1;         ;# Close the callback gate
_validate_vtbl_after:
    mov     edx, _patch5_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0
    je      _unhook
_callback_after:    
    push    lParam                      ;# ByVal lParam
    push    wParam                      ;# ByVal wParam
    push    nCode                       ;# ByVal uMsg
    lea     eax, lReturn                ;# ByRef lReturn
    push    eax                         ;#
    mov     eax, [edx]                  ;# Vtbl
    push    edx                         ;#    
    call    dword ptr [eax][1Ch]        ;# Call iHook_After
    mov     dword ptr [esi], 0          ;# Open the callback gate
_return:                                ;#
    pop     esi                         ;#
    mov     eax, lReturn                ;# Function return value
    ret
_unhook:
    push    _patch6_                    ;# Current hook handle
    call    Dummy                       ;# Call UnhookWindowsHookEx, patched at runtime
    xor     eax, eax
    mov     lReturn, eax
    jmp     _return

HookProc endp
Dummy:
end start