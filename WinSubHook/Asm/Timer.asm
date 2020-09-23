;##########################################################################################################################
;#
;# cTimer,cls model assembler source

.486                                    ;# Create 32 bit code
.model flat, stdcall                    ;# 32 bit memory model
option casemap :none                    ;# Case sensitive
include Thunk.inc                       ;# Macros 'n stuff

_patch1_    = 01BCCAABh                 ;# Callback break gate address
_patch2_    = 02BCCAABh                 ;# Object address of the owner Form/UserControl/Class
_patch3_    = 03BCCAABh                 ;# Start time (GetTickCount)
_patch4_    = 04BCCAABh                 ;# Timer ID
                                        ;# _patch5_ call Dummy, patched to KillTimer at runtime
.code

start:

WndProc proc    hWin    :DWORD,         ;# Window handle
                uMsg    :DWORD,         ;# Message number
                idEvent :DWORD,         ;# Event ID
                lTime   :DWORD          ;# System time
_init:    
    push    esi                         ;# Preserve esi
    mov     esi, _patch1_               ;# Break gate address, patched at runtime
_validate_vtbl:
    mov     edx, _patch2_               ;# Object address of the owner Form/UserControl/Class, patched at runtime
    mov     eax, [edx]                  ;# Vtbl
    cmp     eax, 0                      ;# Validate Vtbl
    je      _destroy_timer              ;# Vtbl invalid
_check_callback_gate:
    cmp     dword ptr [esi], 0          ;# Check the callback gate
    jne     _return                     ;# If the callback gate is set, return
    mov     dword ptr [esi], 1          ;# Close the callback gate
    mov     eax, lTime                  ;# Prepare elapsed time calc
    sub     eax, _patch3_               ;# Calculate elapsed time, patched at runtime
    push    eax                         ;# ByVal elapsed time
    mov     eax, [edx]                  ;# Vtbl
    push    edx                         ;#
    call    dword ptr [eax][1Ch]        ;# Call iTimer_Fire
    mov     dword ptr [esi], 0          ;# Open the callback gate
_return:                                ;#
    pop     esi                         ;#
    ret
_destroy_timer:
    push    _patch4_                    ;# Push the timer id
    push    0                           ;# NULL
    call    Dummy                       ;# Call KillTimer, patched at runtime
    jmp     _return
WndProc endp
Dummy:
end start
