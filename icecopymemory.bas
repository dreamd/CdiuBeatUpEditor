Attribute VB_Name = "cq"
Private Declare Function CallWindowProcW Lib "user32" (ByRef lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "ntdll" (ByVal pDst As Long, ByVal pSrc As Long, ByVal dwLength As Long)
Dim Asm(37) As Long, AsmStub(1) As Long, AsmPtr As Long, IDEMode As Long
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private Function RetLng(ByVal dwAny As Long) As Long
    RetLng = dwAny
End Function

Private Sub InitAsm()
    Dim FuncAddr As Long, OldProtect As Long
    Asm(0) = &H8B575651: Asm(1) = &H8B10247C: Asm(2) = &H8B142474: Asm(3) = &H5118244C
    Asm(4) = &HF33FE183: Asm(5) = &HE18359A4: Asm(6) = &H6850FC0: Asm(7) = &H5F000000
    Asm(8) = &HCC2595E: Asm(9) = &H3E9C100: Asm(10) = &H8DCE348D: Asm(11) = &HD9F7CF3C
    Asm(12) = &HCE84180F: Asm(13) = &H200&: Asm(14) = &HCE046F0F: Asm(15) = &HCE4C6F0F
    Asm(16) = &H4E70F08: Asm(17) = &H4CE70FCF: Asm(18) = &H6F0F08CF: Asm(19) = &HF10CE44
    Asm(20) = &H18CE4C6F: Asm(21) = &HCF44E70F: Asm(22) = &H4CE70F10: Asm(23) = &H6F0F18CF
    Asm(24) = &HF20CE44: Asm(25) = &H28CE4C6F: Asm(26) = &HCF44E70F: Asm(27) = &H4CE70F20
    Asm(28) = &H6F0F28CF: Asm(29) = &HF30CE44: Asm(30) = &H38CE4C6F: Asm(31) = &HCF44E70F
    Asm(32) = &H4CE70F30: Asm(33) = &HC18338CF: Asm(34) = &HFA57508: Asm(35) = &H770FF8AE
    Asm(36) = &HC2595E5F: Asm(37) = &HC&
    AsmStub(0) = &HFF505A58: AsmStub(1) = &HE2&
    AsmPtr = VarPtr(Asm(0))
    FuncAddr = RetLng(AddressOf IcyCopyMemory)
    VirtualProtect FuncAddr, 5, PAGE_EXECUTE_READWRITE, OldProtect
    RtlMoveMemory FuncAddr, VarPtr(&HE9), 1
    RtlMoveMemory FuncAddr + 1, VarPtr(AsmPtr - (FuncAddr + 5)), 4
    VirtualProtect FuncAddr, 5, OldProtect, OldProtect
End Sub

Private Function SetIDE() As Boolean
    IDEMode = -1
    SetIDE = True
End Function

Public Sub IcyCopyMemory(ByVal pDst As Long, ByVal pSrc As Long, ByVal dwLength As Long)
    If AsmPtr = 0 Then InitAsm
    Debug.Assert SetIDE
    If IDEMode = 0 Then
        IcyCopyMemory pDst, pSrc, dwLength
    Else
        CallWindowProcW AsmStub(0), AsmPtr, pDst, pSrc, dwLength
    End If
End Sub

