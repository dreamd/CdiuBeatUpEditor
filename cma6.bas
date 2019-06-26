Attribute VB_Name = "cma6"
Public DI As DirectInput8
Public DIDev As DirectInputDevice8
'Public dxa As New DirectX8
Public dxa As DirectX8
Public DIState As DIKEYBOARDSTATE
Public DIState2 As DIKEYBOARDSTATE

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Dim GetKey As Boolean

Public Sub InitDI()

If InitedA = True Then Exit Sub
InitedA = True

If Admin = False Then On Error Resume Next

  Set dxa = CreateObject("DIRECT.DirectX8.0")
  Set DI = dxa.DirectInputCreate()
  Set DIDev = DI.CreateDevice("GUID_SysKeyboard")
  DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
  DIDev.SetCooperativeLevel cmt.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'DISCL_FOREGROUND '
  DIDev.Acquire
End Sub

Public Sub CloseAll(Optional Back As Boolean)

If Admin = False Then On Error Resume Next

    cma2.FreeLibrary hLibR
    cma1.CloseSound
    cma4.UnloadD3D
    cma6.UnloadDI
    If Back = False Then cma7.ExitExe

End Sub

Public Sub UnloadDI()

If Admin = False Then On Error Resume Next

    Set DI = Nothing
    Set DIDev = Nothing
    Set dxa = Nothing

End Sub

Public Function CheckTime(ToNowBeat As Long, ToOffset As Single, Optional CType As Boolean = True)

If Admin = False Then On Error Resume Next

Dim ToNowBeat2 As Double
Static LastBeat As Long

If ChangeBpm = True Then
        If CType = False Then
        
                ToNowBeat = FindWhichBeatA(LastBeat, cmt.Times.value - OffSet, 0.5) + 1
                If ToNowBeat < 0 Then ToNowBeat = FindWhichBeat(PData, cmt.Times.value - OffSet, 0.5) + 1
                ToOffset = FindWhichOffset(PData, cmt.Times.value - OffSet, ToNowBeat)
                ToNowBeat = ToNowBeat + 1
                ToOffset = FormatNumber(ToOffset, 4)
        Else
                ToNowBeat = FindWhichBeatA(LastBeat, cmt.Times.value - OffSet) + 1
                If ToNowBeat < 0 Then ToNowBeat = FindWhichBeat(PData, cmt.Times.value - OffSet) + 1
                ToOffset = FindWhichOffset(PData, cmt.Times.value - OffSet, ToNowBeat) + 1
                If ToOffset >= 1 Then ToOffset = ToOffset - 1: ToNowBeat = ToNowBeat + 1
                ToOffset = FormatNumber(ToOffset, 4)
        End If
Else
        If CType = False Then
                ToNowBeat2 = CDbl(cmt.Times.value - OffSet) * CBT + CDbl(1.5)
                ToNowBeat = CLng(ToNowBeat2)
                ToOffset = FormatNumber(ToNowBeat2 - ToNowBeat, 4)
                If ToOffset = 0.5 Then ToOffset = -0.5: ToNowBeat = ToNowBeat + 1
        Else
                ToNowBeat2 = CDbl(cmt.Times.value - OffSet) * CBT + CDbl(1)
                ToNowBeat = CLng(ToNowBeat2)
                ToOffset = FormatNumber(ToNowBeat2 - ToNowBeat + 0.5, 4)
                If ToOffset = 1 Then ToOffset = 0: ToNowBeat = ToNowBeat + 1
        End If
End If

        LastBeat = ToNowBeat

End Function

Public Sub DXKeyboard()

Dim ToNowBeat As Long, ToOffset As Single, k As Long, Which As Byte, Add As Long, cI As Long, CheckPoffset(2) As Single

If Admin = False Then On Error Resume Next

Static cDoTime As Long

CheckPoffset(0) = 0.3
CheckPoffset(1) = 0.4
CheckPoffset(2) = 0.5

If GetForegroundWindow() <> cmt.hwnd Then Exit Sub

DIDev.GetDeviceStateKeyboard DIState

cma6.CheckTime ToNowBeat, ToOffset, False

If ToNowBeat < 1 Then Exit Sub
CheckGData ToNowBeat
    
    If UseMode <> "game" And cmt.Frame2.Visible = True And DIState.Key(DIK_SPACE) <> 0 And DIState.Key(DIK_SPACE) <> DIState2.Key(DIK_SPACE) Then
        If timeGetTime() - cDoTime > 1000 Then
            cDoTime = timeGetTime()
            cma1.DoPlayOrStop
        End If
    End If
    
    If (UseMode = "normal" And Mode = "playing" And (ToNowBeat + 2) > 63) Or UseMode = "game" Then
    
            For k = 0 To 8
            
                Select Case k
                    Case 0: Which = DIK_NUMPAD9
                    Case 1: Which = DIK_NUMPAD6
                    Case 2: Which = DIK_NUMPAD3
                    Case 3: Which = DIK_NUMPAD7
                    Case 4: Which = DIK_NUMPAD4
                    Case 5: Which = DIK_NUMPAD1
                    Case 6: Which = DIK_NUMPAD0
                    Case 7: Which = DIK_NUMPAD5
                    Case 8: Which = DIK_SPACE
                End Select
                
                Add = k
                If Add > 6 Then Add = 6
                    If DIState.Key(Which) <> 0 And DIState.Key(Which) <> DIState2.Key(Which) Then
                    
                        If UseMode = "normal" Then
                                GData(ToNowBeat * 8 + Add) = True
                                PlaySoundMusic IIf(Add = 6, 1, 0)
                                KeyPKey = ToNowBeat
                        ElseIf UseMode = "game" Then
                            If GData(ToNowBeat * 8 + Add) = True And GameCheck(ToNowBeat * 8 + Add) = 0 Then '
                                         For cI = 0 To 2
                                              If IIf(InStr(CStr(ToOffset), "-") > 0, CSng(Replace(CStr(ToOffset), "-", "")), ToOffset) <= CheckPoffset(cI) Then DoForm cI, ToNowBeat, Add: Exit For
                                          Next cI
                            ElseIf GData((ToNowBeat + 1) * 8 + Add) = True And GameCheck((ToNowBeat + 1) * 8 + Add) = 0 Then
                                        If ToOffset >= 0 Then DoForm 2, ToNowBeat + 1, Add
                                        If ToOffset < 0 Then DoForm 3, ToNowBeat + 1, Add
                            ElseIf GData((ToNowBeat - 1) * 8 + Add) = True And GameCheck((ToNowBeat - 1) * 8 + Add) = 0 Then '
                                        If ToOffset >= 0 Then DoForm 3, ToNowBeat - 1, Add
                                        If ToOffset < 0 Then DoForm 2, ToNowBeat - 1, Add
                            End If
                        End If
                                KeyTime(Add) = ToNowBeat + ToOffset
                                CheckOffset = FormatNumber(ToOffset + OffsetSet, 4)
                                PressOffset = ToOffset
                    ElseIf DIState.Key(Which) <> 0 And DIState.Key(Which) = DIState2.Key(Which) Then
                        KeyTime(Add) = ToNowBeat + ToOffset
                    End If
            Next k

    End If
    DIState2 = DIState
End Sub

Public Sub DoForm(Which As Long, Beat As Long, Add As Long)

If Admin = False Then On Error Resume Next

Dim XP As Single

                                  NowCombo = NowCombo + 1
                                  GameR = Which
                                  GameCheck(Beat * 8 + Add) = Which + 1
                                  GameP(Which) = GameP(Which) + 1
                                  PlaySoundMusic IIf(Add = 6, 1, IIf(Which >= 1, 2, 0))
                                  
                                  XP = 1
                                  If NowCombo >= 100 Then XP = 1.3
                                  If NowCombo >= 400 Then XP = 1.3
                                  
                                  If Add < 6 Then
                                    Select Case Which
                                      Case 0: Score = Score + 520 * XP
                                      Case 1: Score = Score + 260 * XP
                                      Case 2: Score = Score + 130 * XP
                                      Case 3: Score = Score + 26 * XP
                                    End Select
                                  Else
                                     Select Case Which
                                      Case 0: Score = Score + 2000 * XP
                                      Case 1: Score = Score + 1500 * XP
                                      Case 2: Score = Score + 1000 * XP
                                      Case 3: Score = Score + 500 * XP
                                    End Select
                                  End If

End Sub

Public Sub PlaySoundMusic(SoWhich As Integer)

If Admin = False Then On Error Resume Next

Select Case SoWhich
    Case 0: cma1.PlayNote
    Case 1: cma1.PlaySpace
    Case 2: cma1.PlayNote 1
End Select

End Sub
