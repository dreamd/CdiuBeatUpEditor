Attribute VB_Name = "MCPU_PRoc"
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Any, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

Private Const INFINITE = &HFFFF
Private Const WAIT_TIMEOUT = &H102

Const Asm_XorData = "558BEC8B45088B088B79088B48048B70088B500C33DB8A07320433880747433BDA7605BB000000004975EBC9C20400"
Const Asm_StrToBin = "558BEC8B45088B088B79088B48048B71088B480803700C8A0650E81B0000008AD8C0E304468A0650E80D0000000AD8881F47464975E1C9C20400558BEC8A45083C3072083C3977042C30EB163C6172083C6677042C57EB0A3C4172063C4677022C37C9C20400"
Const Asm_ReverseByte = "558BEC8B45088B48048B71088BFE8B480803F94FD1F98A068A27880788264F464975F3C9C20400"
Const Asm_ByteToHex = "558BEC8B45088B088B79088B48048B71088B48088A1EC0FB0480E30F8AC350E8180000008807478A06240F50E80B000000880747464975DCC9C20400558BEC8A45083C00720C3C0977080430C9C20400EB0E3C0A720A3C0F77060437C9C20400C9C20400"


Private Function RunAsmCode(ACodeS As String, ParamArray Data()) As Long
Dim iParam() As Long, ACodeR() As Byte
Dim hProc As Long, i As Long
Dim StartTime As Long


    ACodeR = StrToBin(ACodeS, StartTime)
StartTime = timeGetTime

    ReDim iParam(UBound(Data))

    For i = 0 To UBound(Data)
        iParam(i) = Data(i)
    Next

    hProc = CreateThread(0, 0, ACodeR(0), iParam(0), 0, ByVal 0)
    Do
        DoEvents
    Loop Until WaitForSingleObject(hProc, 100) = 0
    
    CloseHandle hProc
    
RunAsmCode = timeGetTime - StartTime
End Function

Public Function LargeXorData(AData, KeyData As String, XorTime)
Dim KeyByte() As Byte, Time As Long

    KeyByte = StrConv(KeyData, vbFromUnicode)
    Time = RunAsmCode(Asm_XorData, VarPtr(AData(0)), UBound(AData) + 1, VarPtr(KeyByte(0)), UBound(KeyByte))
    XorTime = Time
End Function

Public Function LargeStrToBin(Data, OData, Pos As Long, StrToBinTime)
Dim ByteUBound As Long, Time As Long

    ByteUBound = ((UBound(Data) - Pos) / 2) - 1
    ReDim OData(ByteUBound)
    Time = RunAsmCode(Asm_StrToBin, VarPtr(OData(0)), VarPtr(Data(0)), ByteUBound + 1, Pos + 1)
    StrToBinTime = Time
End Function

Public Function LargeByteToHex(Data, OData, ByteToHexTime)
Dim ByteUBound As Long, Time As Long
If UBound(Data) > 0 Then
    ByteUBound = UBound(Data)
    ReDim OData(ByteUBound * 2 + 1)
    Time = RunAsmCode(Asm_ByteToHex, VarPtr(OData(0)), VarPtr(Data(0)), ByteUBound + 1)
End If
    ByteToHexTime = Time
End Function

Public Function LargeReverseByte(Data, ReverseByteTime)
Dim Time As Long

     Time = RunAsmCode(Asm_ReverseByte, CLng(0), VarPtr(Data(0)), UBound(Data) + 1)
     ReverseByteTime = Time
End Function

