Attribute VB_Name = "cma1"
Option Explicit

Dim Fso As New FileSystemObject

Public Const VerionA = "3"
Public Const VerionB = "5"

Dim SoundH As Long
Dim BeatSound(2 To 15) As Long, SpaceSound(2 To 15)  As Long
Dim ReadySound As Long, GoSound As Long, Miss(2 To 15) As Long, Great(2 To 15) As Long

Public SoundL As Long
Public TotalBeat As Long '總beat數
Public Mode As String '放播模式
Public Mouse As Boolean 'check 播放拉果行野
Public CBT As Single 'bpm每粒時間
Public OffSet As Double '歌曲offset
Public SongPath As String '歌曲位置
Public TempPath As String '臨時資料夾
Public hLibR As Long '自動載入dll
Public hLibL As Long '自動載入dll
Public GData() As Boolean 'key 資料
Public UData() As Boolean

Public MData() As Boolean
Public NData() As Boolean
Public BData() As Boolean

Public SData() As Byte
Public OUData() As Byte

Public PData() As Long

Public CurTime As Long 'ogg播放現時時間
Public ShitTime As Integer 'ogg播放現時時間
Public ShitTime2 As Integer 'ogg播放現時時間
Public UseMode As String
Public CurrPos As Long
Public User As String
Public SaveFileINI As String
Public OggF As String
Public ScrF As String
Public ASL As String
Public BUL As String
Public Ddr_Desc As String
Public Ddr_devide As String
Public Ddr_image As String
Public C_Bpm As Single
Public Singer As String
Public Melody As String
Public Author As String
Public level As String
Public MusicCode As String
Public Inited As Boolean
Public InitedA As Boolean
Public SaUnDo As Boolean

Public Room As Boolean

Public BpmSet(8) As Single

Public OffsetSet As Single
Public GameCheck() As Byte
Public KeyPKey As Long

Public Language As Long
Public NowCombo As Long
Public GameP(4) As Long
Public GameR As Long

Public Score As Long

Public ChooseBackGround As Long
Public CheckOffset As Single
Public PressOffset As Single

Public KeyTime(6) As Single

Public SetRx As Long
Public SaveSelect() As Long
Public FastTeam() As String
Public Admin As Boolean

Public ChangeBpm As Boolean

Public Slk1() As Byte
Public Slk2() As Byte
Public Slk3() As Byte
Public Slk4() As Byte

Public Code1 As Long
Public Code2 As Long

Public CodeA As Long
Public code As Long
Public GameAuthor As String
Public GameMelody As String
Public GameSinger As String

Public ChatRoom As Boolean

Public NetWork As Boolean

Public PressA As Boolean
Public PressTime As Long

Public RoomNumber As Long
Public Roomid As Long
Public RoomPassword As String
Public RoomName As String

Public CheckFile() As String

Public Sub LoadEffect(Name As String, Add() As Byte)

Dim CData As Long, i As Long

    If Name = "SOUND\BEAT.ogg" Then
        For i = LBound(BeatSound) To UBound(BeatSound)
            CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))

            CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)

            BeatSound(i) = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Next i
        Exit Sub
    End If
    
    If Name = "SOUND\SPACE.ogg" Then
        For i = LBound(SpaceSound) To UBound(SpaceSound)

            CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))

            CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)

            SpaceSound(i) = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Next i
        Exit Sub
    End If

    If Name = "SOUND\START.ogg" Then
        CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))
        
        CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)
        
        GoSound = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Exit Sub
    End If
    
    If Name = "SOUND\READY.ogg" Then
        CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))
        
        CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)
        
        ReadySound = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Exit Sub
    End If



    If Name = "SOUND\MISS.ogg" Then
        For i = LBound(Miss) To UBound(Miss)
            CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))
            
            CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)
            
            Miss(i) = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Next i
        Exit Sub
    End If

    If Name = "SOUND\GREAT.ogg" Then
        For i = LBound(Great) To UBound(Great)
            CData = GlobalAlloc(&H40, CLng(UBound(Add) + 1))
            
            CopyMemory ByVal CData, Add(0), CLng(UBound(Add) + 1)
            
            Great(i) = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(Add) + 1))
        Next i
        Exit Sub
    End If

End Sub

Public Sub LoadSound()

Dim Result As Boolean

If Admin = False Then On Error Resume Next

Result = FSOUND_Init(44100, 32, 0)

End Sub

Public Function OpenSoundByte(Path As String, Optional NotShow As Boolean)

Dim i As Long, k As Long, SongFile() As Byte, CData As Long, SigT As String * 22, LenTPData As Long, SigT2 As String * 22, RealName As String * 512, Tmp As Long, CBT2 As Single, Offset2 As Double, ChooseBackGroundA As Long, NoUse As Long, NoUse2 As Integer

If Admin = False Then On Error Resume Next

        If SoundH <> 0 Then FSOUND_Stream_Close SoundH
        
        Decrypt_12 Path, 0

            If Fso.FileExists(App.Path + "\Game\cma\" + "song.ogg") = True Then
                Open App.Path + "\Game\cma\" + "song.ogg" For Binary As #1
                    ReDim SongFile(FileLen(App.Path + "\Game\cma\" + "song.ogg") - 1)
                    Get #1, 1, SongFile
                Close #1
            End If

        DeleteDir App.Path + "\Game\cma\"
            
            CData = GlobalAlloc(&H40, CLng(UBound(SongFile) + 1))
            
            CopyMemory ByVal CData, SongFile(0), CLng(UBound(SongFile) + 1)
            
            SoundH = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(SongFile) + 1))
        
        If SoundH <> 0 Then SoundL = FSOUND_Stream_GetLengthMs(SoundH)

TotalBeat = SoundL * CBT

If SoundL > 255 Then cmt.Times.Max = SoundL

ReDim GData((TotalBeat + 35) * 8)

ReDim SData((TotalBeat + 35) * 8)

    For i = 0 To TotalBeat + 34
        For k = 0 To 6
            SData(i * 8 + k) = 4
        Next k
    Next i

FindBeatTime PData

        Decrypt_12 Path, 0

        If Fso.FileExists(App.Path + "\Game\cma\" + "maindata.cbe") = True Then
            Open App.Path + "\Game\cma\" + "maindata.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileData" Then
                Get #1, , LenTPData
                    ReDim GData(LenTPData)
                    TotalBeat = (LenTPData / 8)
                Get #1, , GData
            End If
            Close #1
        End If
        
        If Fso.FileExists(App.Path + "\Game\cma\" + "sdata.cbe") = True Then
            Open App.Path + "\Game\cma\" + "sdata.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileSata" Then
                Get #1, , LenTPData
                    ReDim SData(LenTPData)
                Get #1, , SData
            End If
            Close #1
        End If
        
        If Fso.FileExists(App.Path + "\Game\cma\" + "infodata.cbe") = True Then
            Open App.Path + "\Game\cma\" + "infodata.cbe" For Binary As #1
            Get #1, 1, SigT2
            If SigT2 = "CdiuBeatEditorFileInfo" Then
                Get #1, , CurTime
                Get #1, , ShitTime
                Get #1, , CurrPos
                Get #1, , SetRx
                Get #1, , Tmp
                cmt.Times.value = Tmp
                
                Get #1, , RealName
                Singer = Trim(RealName)
                cmt.Single_Text = Singer
        
                Get #1, , RealName
                Melody = Trim(RealName)
                cmt.Melody_Text = Melody
                
                Get #1, , RealName
                Author = Trim(RealName)
                cmt.Author_Text = Author
                
                Get #1, , RealName
                level = Trim(RealName)
                cmt.Level_Text = level
                
                Get #1, , RealName
                MusicCode = Trim(RealName)
                cmt.MusicCode_Text = MusicCode
                
                Get #1, , RealName
                'UseMode = Trim(RealName)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(0) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(0) = BpmSet(0)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then OffsetSet = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Offset = OffsetSet
                
                Get #1, , CBT2
                Get #1, , Offset2
                If IsNumeric(CBT2) = True Then CBT = CBT2
                If IsNumeric(Offset2) = True Then OffSet = Offset2
                
                Get #1, , RealName
                SongPath = Trim(RealName)
                
                Get #1, , ChooseBackGroundA
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(1) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(1) = BpmSet(1)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(2) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(2) = BpmSet(2)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(3) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(3) = BpmSet(3)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(4) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(4) = BpmSet(4)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(5) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(5) = BpmSet(5)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(6) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(6) = BpmSet(6)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(7) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(7) = BpmSet(7)
                
                Get #1, , RealName
                If IsNumeric(Trim(RealName)) = True Then BpmSet(8) = Trim(RealName)
                If IsNumeric(Trim(RealName)) = True Then cmt.OK_Bpm.Item(8) = BpmSet(8)
            End If
            Close #1
        End If

        DeleteDir App.Path + "\Game\cma\"
        
cma3.CheckChangeBpm
ChangeMapByUser ChooseBackGroundA

End Function

Public Function OpenSound(Path, Optional NotShow As Boolean)

Dim i As Long, k As Long, SongFile() As Byte, CData As Long

If Admin = False Then On Error Resume Next

If NotShow = False Then cmt.Frame1.Visible = True
If NotShow = False Then cmt.OK_Bpm.Item(0).SetFocus
cmt.openmusic.Enabled = False
cmt.NewFile.Enabled = True
cmt.setting.Enabled = True
cmt.ProSetting.Enabled = True

        If SoundH <> 0 Then FSOUND_Stream_Close SoundH
        
        'SoundH = FSOUND_Stream_Open(Path, FSOUND_NORMAL, 0, 0)
        
        
            Open Path For Binary As #1
                ReDim SongFile(FileLen(SongPath) - 1)
                Get #1, 1, SongFile
            Close #1
            CData = GlobalAlloc(&H40, CLng(UBound(SongFile) + 1))
            
            CopyMemory ByVal CData, SongFile(0), CLng(UBound(SongFile) + 1)
            
            SoundH = FSOUND_Stream_Open2(ByVal CData, &H2A130, 0, CLng(UBound(SongFile) + 1))
        
        If SoundH <> 0 Then SoundL = FSOUND_Stream_GetLengthMs(SoundH)

TotalBeat = SoundL * CBT

If SoundL > 255 Then cmt.Times.Max = SoundL

ReDim GData((TotalBeat + 35) * 8)

ReDim SData((TotalBeat + 35) * 8)

ReDim MData((TotalBeat + 35) * 8)
ReDim NData((TotalBeat + 35) * 8)
ReDim BData((TotalBeat + 35) * 8)

    For i = 0 To TotalBeat + 34
        For k = 0 To 6
            SData(i * 8 + k) = 4
        Next k
    Next i

FindBeatTime PData
If NotShow = True Then cma3.LoadCbe App.Path + "\Temp.cdiu"

UseMode = "see"
cma2.UseModeChange

End Function

Public Function NoSound()

Dim Result As Boolean

If Admin = False Then On Error Resume Next

Mode = "close"
CurTime = 0
cmt.Times.value = 0
CurrPos = 0
NowCombo = 0

'Result = FSOUND_Stream_Close(SoundH)

'SoundH = FSOUND_Stream_Open(SongPath, FSOUND_NORMAL, 0, 0)

'cmt.Times.value = 0

FSOUND_Stream_Stop SoundH

End Function

Public Function CloseSound()

Dim Result As Boolean

If Admin = False Then On Error Resume Next

Mode = "close"

Result = FSOUND_Stream_Close(SoundH)
Result = FSOUND_Close

End Function

Public Function FindWhichOffset(PData, TimeA As Long, NowBeat As Long) As Single

If Admin = False Then On Error Resume Next

If NowBeat <> 0 Then
    If TimeA >= PData(NowBeat + 1) Then
        FindWhichOffset = (TimeA - PData(NowBeat + 1)) / (PData(NowBeat + 1) - PData(NowBeat))
    Else
        FindWhichOffset = (TimeA - PData(NowBeat)) / (PData(NowBeat) - PData(NowBeat - 1))
    End If
End If

End Function

Public Function FindWhichBeat(PData, TimeA As Long, Optional NoBeat As Single) As Long

Dim i As Long, ABeat As Single

If Admin = False Then On Error Resume Next

ABeat = NoBeat / (BpmSet(0) / 1000 / 60 * 4)

For i = 0 To UBound(PData) - 2
    If TimeA >= PData(i) + ABeat And TimeA < PData(i + 1) + ABeat Then FindWhichBeat = i: Exit For
Next i

End Function

Public Function FindWhichBeatA(LastBeat, TimeA As Long, Optional NoBeat As Single) As Long

Dim i As Long, ABeat As Single, Check As Boolean

If Admin = False Then On Error Resume Next

ABeat = NoBeat / (BpmSet(0) / 1000 / 60 * 4)

For i = LastBeat - 5 To LastBeat + 5

    If i < 0 Then Exit For
    If i > UBound(PData) - 2 Then Exit For
    
    If TimeA >= PData(i) + ABeat And TimeA < PData(i + 1) + ABeat Then FindWhichBeatA = i: Check = True: Exit For
Next i

If Check = False Then FindWhichBeatA = -100

End Function

Public Function FindBeatTime(PData)

Dim i As Long, Tmp() As Single

If Admin = False Then On Error Resume Next

ReDim PData(TotalBeat + 3)
ReDim Tmp(TotalBeat + 3)
If TotalBeat <= 10 Then
    ReDim Preserve SData(70)
Else
    ReDim Preserve SData((TotalBeat + 4) * 8)
End If

Tmp(0) = -0.5 / (BpmSet(SData(7)) / 1000 / 60 * 4)
For i = 1 To UBound(Tmp)
    Tmp(i) = Tmp(i - 1) + 1 / (BpmSet(SData(i * 8 + 7)) / 1000 / 60 * 4)
Next i


For i = 0 To UBound(Tmp)
    PData(i) = CLng(Tmp(i))
Next i

End Function

Public Function PlaySound(Pos As Long, Optional StartBeat As Long, Optional NextBeat As Long)

Dim Result As Boolean, ToNowBeat As Long, Start As Integer, Times As Integer, i As Integer, PlayBeat As Long, ResultA As String

If Admin = False Then On Error Resume Next

'FindBeatTime PData

Mode = "playing"
If UseMode = "game" Then
        NowCombo = 0
        GameP(0) = 0
        GameP(1) = 0
        GameP(2) = 0
        GameP(3) = 0
        GameP(4) = 0
        KeyTime(0) = 0
        KeyTime(1) = 0
        KeyTime(2) = 0
        KeyTime(3) = 0
        KeyTime(4) = 0
        KeyTime(5) = 0
        KeyTime(6) = 0
        KeyPKey = 0
        Score = 0
        ReDim GameCheck(TotalBeat * 8)
End If

If cmt.Times.value < 1 Then cmt.Times.value = 0

        If cmt.Times.value <> 0 Then
            FSOUND_Stream_SetTime SoundH, cmt.Times.value
        ElseIf CurrPos <> 0 Then
            FSOUND_Stream_SetTime SoundH, CurrPos - OffSet
            CurrPos = 0
        Else
            FSOUND_Stream_SetTime SoundH, 0
        End If

        If SoundL <> 0 Then
            FSOUND_Stream_Play 1, SoundH
        Else
            EndTheSong
        End If

    Do
            FSOUND_SetVolumeAbsolute 1, 255 - cmt.sldSpeed.value
            
            CurTime = FSOUND_Stream_GetTime(SoundH)
            If (Mouse = False) Then cmt.Times.value = CurTime
            
            If ChangeBpm = True Then
                ToNowBeat = FindWhichBeat(PData, CurTime - OffSet) + 1
            Else
                ToNowBeat = (CurTime - OffSet) * CBT + 1
            End If

            If ToNowBeat < 1 Then GoTo DoNext
   
            If StartBeat <> 0 And ToNowBeat >= StartBeat + NextBeat And UseMode <> "game" Then cma1.ChangeTime PData(ToNowBeat - NextBeat)
            cma4.Render
            
            If (PlayBeat <> ToNowBeat) Then
            
                PlayBeat = ToNowBeat
                
                If UseMode <> "game" Then
                        For i = 0 To 5
                                If ToNowBeat * 8 + 7 >= UBound(GData) Then ReDim Preserve GData((ToNowBeat + 1) * 8)
                                If GData(ToNowBeat * 8 + i) = True And ToNowBeat <> KeyPKey Then PlayNote: Exit For
                        Next i
                        
                        If GData(ToNowBeat * 8 + 6) = True And ToNowBeat <> KeyPKey Then PlaySpace
                Else
                    CheckMiss
                End If
            End If
DoNext:
    DoEvents
    Loop While CurTime < SoundL And Mode = "playing"

                If NetWork = True Then
                    cmt.Hide
                    cmt.Enabled = False
                    cmt.Height = 12300
                    Room = True
                    cma1.CloseSound
                    cma4.UnloadD3D
                    cma6.UnloadDI
                    ChatBox.Enabled = True
                    ChatBox.Show
                    ChatBox.UpdateText.Enabled = True
                    ChatBox.Timer1.Enabled = True
                    
                    ResultA = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F555E541E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16535F54550D"))
                    
                    ResultA = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F425543455C441E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16535F54550D") + GetLink("16400D") + GetCode(CStr(GameP(0))) + GetLink("16570D") + GetCode(CStr(GameP(1))) + GetLink("16530D") + GetCode(CStr(GameP(2))) + GetLink("16520D") + GetCode(CStr(GameP(3))) + GetLink("165D0D") + GetCode(CStr(GameP(4))) + GetLink("16430D") + GetCode(CStr(Score)))
                    ChatBox.UpdateTextShow
                    
                Else
                    If Mode <> "stop" Then EndTheSong
                End If

End Function

Public Sub CheckMiss()

Dim ToNowBeat As Long, ToOffset As Single, k As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 10 Then Exit Sub
cma4.CheckGData ToNowBeat

    For k = 0 To 6
        If GameCheck((ToNowBeat - 2) * 8 + k) = 0 And GData((ToNowBeat - 2) * 8 + k) = True Then PlaySpace 1: FixCombo: GameP(4) = GameP(4) + 1: GameR = 4: GameCheck((ToNowBeat - 2) * 8 + k) = 5
    Next k

End Sub

Public Sub FixCombo()

If Admin = False Then On Error Resume Next

    If NowCombo >= 100 Then
        NowCombo = 0
    ElseIf NowCombo > 10 Then
        If CLng(Mid(CStr(NowCombo), 2, 1)) > 0 Then
            NowCombo = NowCombo - CLng(Mid(CStr(NowCombo), 2, 1))
        Else
            NowCombo = NowCombo - CLng(Mid(CStr(NowCombo), 2, 1)) - 11
        End If
        
    Else
        NowCombo = 0
    End If

End Sub

Public Sub ChangeTime(Pos As Long)

If Admin = False Then On Error Resume Next

If UseMode = "game" Then
        NowCombo = 0
        GameP(0) = 0
        GameP(1) = 0
        GameP(2) = 0
        GameP(3) = 0
        GameP(4) = 0
        KeyTime(0) = 0
        KeyTime(1) = 0
        KeyTime(2) = 0
        KeyTime(3) = 0
        KeyTime(4) = 0
        KeyTime(5) = 0
        KeyTime(6) = 0
        KeyPKey = 0
        Score = 0
        ReDim GameCheck(TotalBeat * 8)
End If

FSOUND_Stream_SetTime SoundH, Pos

End Sub

Public Sub PlayReady()

If Admin = False Then On Error Resume Next

FSOUND_Stream_Play 16, ReadySound

End Sub

Public Sub PlayGo()

If Admin = False Then On Error Resume Next

FSOUND_Stream_Play 16, GoSound

End Sub

Public Function StopSound()

If Admin = False Then On Error Resume Next

CurrPos = FSOUND_Stream_GetTime(SoundH)
FSOUND_Stream_Stop SoundH

Mode = "stop"
            
End Function

Public Sub PlayNote(Optional PorG As Long)

Static Number As Integer

If Admin = False Then On Error Resume Next

If (Number >= 16) Or (Number < 2) Then Number = 2

FSOUND_Stream_Play Number, IIf(PorG = 0, BeatSound(Number), Great(Number))
FSOUND_SetVolume Number, 255 - cmt.noteSpeed.value

Number = Number + 1

End Sub

Public Sub PlaySpace(Optional SpaceOrMiss As Long)

Static SpaceNumber As Integer

If Admin = False Then On Error Resume Next

If (SpaceNumber >= 16) Or (SpaceNumber < 2) Then SpaceNumber = 2

FSOUND_Stream_Play SpaceNumber + 14, IIf(SpaceOrMiss = 0, SpaceSound(SpaceNumber), Miss(SpaceNumber))
FSOUND_SetVolume SpaceNumber + 14, 255 - cmt.noteSpeed.value

SpaceNumber = SpaceNumber + 1

End Sub

Public Function DoPlayOrStop(Optional StartBeat As Long, Optional NextBeat As Long)

If Admin = False Then On Error Resume Next

        If cmt.PlayOrStop.Caption = "播放" Or cmt.PlayOrStop.Caption = "Play" Then
                cmt.Button.Item(0).Visible = False
                cmt.PlayOrStop.Caption = IIf(Language = 0, "暫停", "Pause")
                cmt.RightOne.Enabled = False
                cmt.LeftOne.Enabled = False
                cmt.dmenu.Enabled = False
                cmt.emenu.Enabled = False
                cmt.fmenu.Enabled = False
                cmt.gmenu.Enabled = False
                cmt.MapMenu.Enabled = False
                cmt.SetLanguage.Enabled = False
                cmt.PlaySpace.Enabled = False
                cma1.PlaySound 0, StartBeat, NextBeat
        Else
                cmt.Button.Item(0).Visible = True
                cmt.PlayOrStop.Caption = IIf(Language = 0, "播放", "Play")
                cmt.RightOne.Enabled = True
                cmt.LeftOne.Enabled = True
                cmt.dmenu.Enabled = True
                cmt.emenu.Enabled = True
                cmt.fmenu.Enabled = True
                cmt.gmenu.Enabled = True
                cmt.MapMenu.Enabled = True
                cmt.SetLanguage.Enabled = True
                cmt.PlaySpace.Enabled = True
                cma1.StopSound
        End If
End Function
