; https://github.com/ayaka14732/TinyPE-on-Win10/blob/main/step7/stretch.asm
; nasm SetForegroundWindow.asm -o SetForegroundWindow.exe
BITS 64

%define align(n,r) (((n+(r-1))/r)*r)
                            ; DOS Header
    dw 'MZ'                 ; e_magic
    dw 0                    ; [UNUSED] e_cblp
pe_hdr:                                                 ; PE Header
    dw 'PE'                 ; [UNUSED] c_cp             ; Signature
    dw 0                    ; [UNUSED] e_crlc           ; Signature (Cont)
                                                        ; Image File Header
    dw 0x8664               ; [UNUSED] e_cparhdr        ; Machine
code:
    dw 0x01                 ; NumberOfSections
    dd 0                    ; [UNUSED] TimeDateStamp
    dd 0                    ; PointerToSymbolTable
    dd 0                    ; NumberOfSymbols
    dw opt_hdr_size         ; SizeOfOptionalHeader
    dw 0x22                 ; Characteristics
opt_hdr:                                                ; Optional Header, COFF Standard Fields
    dw 0x020b               ; [UNUSED] e_res            ; Magic (PE32+)
    db 0                    ; [UNUSED] e_res (Cont)     ; [UNUSED] MajorLinkerVersion
    db 0                    ; [UNUSED] e_res (Cont)     ; [UNUSED] MinorLinkerVersion
    dd code_size            ; [UNUSED] e_res (Cont)     ; SizeOfCode
    dw 0                    ; [UNUSED] e_oemid          ; [UNUSED] SizeOfInitializedData
    dw 0                    ; [UNUSED] e_oeminfo        ; [UNUSED] SizeOfInitializedData (Cont)
    dd 0                    ; [UNUSED] e_res2           ; [UNUSED] SizeOfUninitializedData
    dd entry                ; [UNUSED] e_res2 (Cont)    ; AddressOfEntryPoint
    dd code                 ; [UNUSED] e_res2 (Cont)    ; BaseOfCode
                                                        ; Optional Header, NT Additional Fields
    dq 0x000140000000       ; [UNUSED] e_res2 (Cont)    ; ImageBase
    dd pe_hdr               ; e_lfanew                  ; [MODIFIED] SectionAlignment (0x10 -> 0x04)
    dd 0x04                                             ; [MODIFIED] FileAlignment (0x10)
    dw 0x06                                             ; [UNUSED] MajorOperatingSystemVersion
    dw 0                                                ; [UNUSED] MinorOperatingSystemVersion
    dw 0                                                ; [UNUSED] MajorImageVersion
    dw 0                                                ; [UNUSED] MinorImageVersion
    dw 0x06                                             ; MajorSubsystemVersion
    dw 0                                                ; MinorSubsystemVersion
    dd 0                                                ; [UNUSED] Reserved1
    dd file_size                                        ; SizeOfImage
    dd hdr_size                                         ; SizeOfHeaders
    dd 0                                                ; [UNUSED] CheckSum
    dw 0x02                                             ; Subsystem (Windows GUI)
    dw 0x8160                                           ; DllCharacteristics
    dq 0x100000                                         ; SizeOfStackReserve
    dq 0x1000                                           ; SizeOfStackCommit
    dq 0x100000                                         ; SizeOfHeapReserve
dll_name:                                                                                   ; DLLName
    db 'USER32.dll', 0                                                                      ; DLLName
    times 12-($-dll_name) db 0                          ; [UNUSED] SizeOfHeapCommit
                                                        ; [UNUSED] LoaderFlags
    dd 0x02                                             ; [MODIFIED] NumberOfRvaAndSizes (0x10)

; Optional Header, Data Directories
lol:
    dd 0                    ; [UNUSED] Export, RVA
    dd 0                    ; [UNUSED] Export, Size
iatbl:                                                  ; Import Address Directory
    dd itbl                 ; Import, RVA               ; [USEDAFTERLOAD] DLLFuncEntry
    dd itbl_size            ; Import, Size              ; [USEDAFTERLOAD] DLLFuncEntry (Cont)

opt_hdr_size equ $-opt_hdr

                            ; Section Table
    section_name db '.', 0  ; Name
    times 8-($-section_name) db 0
    dd sect_size            ; VirtualSize
    dd iatbl                ; VirtualAddress
    dd code_size            ; SizeOfRawData
    dd iatbl                ; PointerToRawData
content:                                                ; Strings
    db 0x41,0x00,0x42,0x00,0x43,0x00,0x44,0x00
    db 0x45,0x00,0x46,0x00,0x47,0x00,0,0
                            ; [UNUSED] PointerToRelocations
                            ; [UNUSED] PointerToLinenumbers
                            ; [UNUSED] NumberOfRelocations
                            ; [UNUSED] NumberOfLinenumbers
                            ; [UNUSED] Characteristics
hdr_size equ $-$$

intblFirstThunk1:                                                  ; Import Name Table
    dq symbol11               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
intblFirstThunk2:                                                  ; Import Name Table
    dq symbol21               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
intblFirstThunk3:                                                  ; Import Name Table
    dq symbol31               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
    times align($-$$,16)-($-$$) db 0xcc

; Entry
entry:
    ;int3
    sub rsp, 28h
    call [rel intbl2]            ; GetCommandLineW
    mov rcx, rax
    lea rdx, [rel pNumArgs2]
    call [rel intbl3]            ; CommandLineToArgvW

    ;AtoI
    mov    rcx, [08h + rax]
    movzx  edx, BYTE [rcx]
    test   dl,dl
    je     endXor
    add    rcx, 02h
    xor    eax,eax
    ;I have no clue what this data16 does: padding for performance ???
    ;data16 data16 data16 data16 data16 cs nopw 0x0(%rax,%rax,1)

    ;db 0x66
    ;db 0x66
    ;db 0x66
    ;db 0x66
    ;db 0x66
    ;nop word cs:[rax + rax*1]
    db 0x66,0x66,0x66,0x66,0x66,0x66,0x2E,0x0F,0x1F,0x84,0x00,0x00,0x00,0x00,0x00

midAtoI:
    lea    eax, [rax + 4*rax]
    xor    dl,30h
    movsx  edx,dl
    lea    eax, [rdx + 2*rax]
    movzx  edx,BYTE [rcx]
    add    rcx, 02h
    test   dl,dl
    jne    midAtoI
    jmp endAtoI
endXor:
    xor    eax,eax
endAtoI:
    mov    ecx,eax
    call [rel intbl]             ; SetForegroundWindow
    add rsp, 28h
    ret
pNumArgs2:
    dd 0
    times align($-$$,16)-($-$$) db 0xcc

itbl:                       ; Import Directory
    dd intbl                ; OriginalFirstThunk
    dd 0                    ; [UNUSED] TimeDateStamp
    dd 0                    ; [UNUSED] Forwarder Chain
    dd dll_name             ; Name
    dd intbl                  ; FirstThunk
    dd intbl2               ; OriginalFirstThunk
    dd 0                    ; [UNUSED] TimeDateStamp
    dd 0                    ; [UNUSED] Forwarder Chain
    dd dll_name2            ; Name
    dd intbl2     ; FirstThunk
    dd intbl3               ; OriginalFirstThunk
    dd 0                    ; [UNUSED] TimeDateStamp
    dd 0                    ; [UNUSED] Forwarder Chain
    dd dll_name3            ; Name
    dd intbl3     ; FirstThunk
    dd 0                    ; OriginalFirstThunk
    dd 0                    ; [UNUSED] TimeDateStamp
    dd 0                    ; [UNUSED] Forwarder Chain
    dd 0                    ; Name
itbl_size equ $-itbl
intbl:                                                  ; Import Name Table
    dq symbol11               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
intbl2:                                                  ; Import Name Table
    dq symbol21               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
intbl3:                                                  ; Import Name Table
    dq symbol31               ; [UNUSED] FirstThunk       ; Symbol
    dq 0
symbol11:
    dw 0x0294               ; [UNUSED] Function Order
    db 'SetForegroundWindow', 0     ; Function Name
symbol21:
    dw 0x0296               ; [UNUSED] Function Order
    db 'GetCommandLineW', 0     ; Function Name
symbol31:
    dw 0x0297               ; [UNUSED] Function Order
    db 'CommandLineToArgvW', 0     ; Function Name
dll_name2:
    db 'KERNEL32.dll', 0
dll_name3:
    db 'Shell32.dll', 0

;sect_size equ $-code
;code_size equ $-code
sect_size equ $-iatbl
code_size equ $-iatbl
file_size equ $-$$