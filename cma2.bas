Attribute VB_Name = "cma2"
Option Explicit

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public ExitMsg As Boolean

Public Function OpenFile(FileWhich As String, SaveOrOpen As String, Optional File2Which As String) As String

On Error GoTo NoFile

Dim i As Integer
Dim c As New cCommonDialog
    
    With c
        .DialogTitle = SaveOrOpen + " File"
        .CancelError = True
        
        If FileWhich = "Ogg" Then
            .InitDir = OggF
        ElseIf FileWhich = "cbg" Then
            .InitDir = App.Path + "\Game\"
        Else
            .InitDir = ScrF
        End If
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
        If File2Which = "" Then
            .Filter = FileWhich + " File (*." + FileWhich + ")|*." + FileWhich + "|All Files (*.*)|*.*"
        Else
            .Filter = FileWhich + " Or " + File2Which + " File (*." + FileWhich + ";*." + File2Which + ")|*." + FileWhich + ";*." + File2Which + "" + "|All Files (*.*)|*.*"
            '.Filter = FileWhich + " Or " + File2Which + " File (*." + FileWhich + ";*." + File2Which + ")|*." + FileWhich + "|All Files (*.*)|*.*"
        End If
        .FilterIndex = 0
        
            Select Case SaveOrOpen
                Case "Open": .ShowOpen
                Case "Save": .ShowSave
            End Select
            
        OpenFile = .FileName
    End With

cmt.PlayOrStop.Enabled = True
cmt.EndSong.Enabled = True

NoFile:

End Function

Public Function CheckDFrame2()

If Admin = False Then On Error Resume Next

cmt.Frame2.Visible = IIf(cmt.Frame2.Visible = True, False, True)

End Function

Public Function SaveData(Optional Done As Boolean)

Dim i As Long, LNumber As Long, u As Long

If Admin = False Then On Error Resume Next

BpmSet(0) = cmt.OK_Bpm(0).Text
BpmSet(1) = cmt.OK_Bpm(1).Text
BpmSet(2) = cmt.OK_Bpm(2).Text
BpmSet(3) = cmt.OK_Bpm(3).Text
BpmSet(4) = cmt.OK_Bpm(4).Text
BpmSet(5) = cmt.OK_Bpm(5).Text
BpmSet(6) = cmt.OK_Bpm(6).Text
BpmSet(7) = cmt.OK_Bpm(7).Text
BpmSet(8) = cmt.OK_Bpm(8).Text

OffsetSet = cmt.OK_Offset.Text ' - 0.7

CBT = BpmSet(0) / 1000 / 60 * 4

OffSet = CInt(15000 / BpmSet(0) * OffsetSet)

cmt.Frame1.Visible = False

    If TotalBeat < SoundL * CBT Then
        TotalBeat = SoundL * CBT
        ReDim Preserve GData(TotalBeat * 8)
        ReDim Preserve MData(TotalBeat * 8)
        ReDim Preserve NData(TotalBeat * 8)
        ReDim Preserve BData(TotalBeat * 8)
        'ReDim Preserve SData(TotalBeat * 8)
    
            For i = 0 To 6
                SData(i) = 4
            Next i
    
    
        If TotalBeat * 8 > UBound(SData) Then
            LNumber = Fix(UBound(SData) / 8)
            ReDim Preserve SData(TotalBeat * 8)
        
                For i = LNumber To Fix(UBound(SData) / 8) - 1
                    For u = 0 To 7
                        SData(i * 8 + u) = SData((LNumber - 1) * 8 + u)
                    Next u
                Next i
        
        End If
    
    End If

If Done = True Then CheckDFrame2: cmt.ShowOrHide.Enabled = True: cmt.BMenu.Enabled = True: cmt.CMenu.Enabled = True: cmt.dmenu.Enabled = True: cmt.emenu.Enabled = True: cmt.fmenu.Enabled = True: cmt.gmenu.Enabled = True: cmt.MapMenu.Enabled = True

End Function

Public Function MstoMin(ms As Long) As String

If Admin = False Then On Error Resume Next

MstoMin = CStr(Fix(ms / 1000 / 60)) + ":" + Right("0" + CStr(Fix(ms / 1000) Mod 60), 2)

End Function

Public Function CheckBeat(Start As Integer, NowBeat As Long) As Boolean

Dim i As Integer

If Admin = False Then On Error Resume Next

        For i = 1 To 100
                If (NowBeat = Start + (32 * i)) Then
                        CheckBeat = True
                        GoTo EndCheckBeat
                End If
        Next i
EndCheckBeat:
End Function

Public Sub MakeTemp(ByVal FileName As String, ByRef Path As String, ByVal NameA As String, ByVal NameB As String)

Dim i As Integer, TmpP() As Byte

If Admin = False Then On Error Resume Next

ReDim TmpP(256)
GetTempPath 256, TmpP(0)

        For i = 0 To 256
                If TmpP(i) = 0 Then
                    ReDim Preserve TmpP(i - 1)
                    Exit For
                End If
        Next i
        
Path = StrConv(TmpP, vbUnicode)

SaveRes Path, FileName, NameA, NameB

End Sub

Public Sub SaveRes(ByVal Path As String, ByVal FileName As String, ByVal Which As String, ByVal WID As String)

If Admin = False Then On Error Resume Next

SaveFileArr Path, FileName, LoadResData(WID, Which)

End Sub

Public Sub SaveFileArr(ByVal SaveFolder As String, ByVal FileName As String, Message() As Byte)

If Admin = False Then On Error Resume Next

CheckUrl SaveFolder

Cdiu_Folder "Make", SaveFolder

FreeLibrary hLibR

Cdiu_File "Delete", SaveFolder, FileName

        If Cdiu_File("Check", TempPath, FileName) Then
                hLibR = cma2.LoadLibrary(CStr(TempPath + FileName))
                
        Else
                Open SaveFolder + "\" + FileName For Binary As #1
                    Put #1, , Message
                Close #1
        End If

End Sub

Public Sub CheckUrl(ByRef Check As String)

If Admin = False Then On Error Resume Next

Check = IIf(Right(Check, 1) = "\", Left(Check, Len(Check) - 1), Check)

End Sub

Public Function Cdiu_Folder(ByVal Action As String, ByVal DoFolder As String, Optional ByVal NewFolder As String) As Boolean

If Admin = False Then On Error Resume Next

    Select Case Action
        Case "Make": CreateDir DoFolder
    End Select
   
End Function

Public Function CreateDir(DirName As String) As Boolean

Dim TempArray() As String, TempFileHolder As String, x As Integer

On Error Resume Next

        If Len(Dir$(DirName, vbDirectory)) = 0 Then
            TempArray = Split(DirName, "\")
            TempFileHolder = TempArray(0)
            
                For x = 1 To UBound(TempArray)
                    TempFileHolder = TempFileHolder & "\" & TempArray(x)
                    
                        If Not (Dir$(TempFileHolder)) Then
                            MkDir TempFileHolder
                        End If
                Next x
            On Error GoTo ErrDebug
            
            CreateDir = IIf(Len(Dir$(DirName, vbDirectory)) = 0, False, True)

        Else
            CreateDir = True
        End If
Exit Function

ErrDebug:
CreateDir = False
    
End Function

Public Function Cdiu_File(ByVal Action As String, ByVal DoFolder As String, ByVal DoFile As String, Optional ByVal NewFolder As String, Optional ByVal NewFile As String) As Boolean

If Admin = False Then On Error Resume Next

    Select Case Action
        Case "Delete": Cdiu_DelFile DoFolder, DoFile
        Case "Check": If Not Dir(DoFolder + "\" + DoFile) = "" Then Cdiu_File = True
        Case "Rename": Cdiu_RenameFile DoFolder, DoFile, NewFolder, NewFile
    End Select

End Function

Public Sub Cdiu_RenameFile(ByVal DoFolder As String, ByVal DoFile As String, ByVal NewFolder As String, ByVal NewFile As String)

On Error GoTo Finish_Cdiu_RenameFile

If DoFolder + "\" + DoFile <> NewFolder + "\" + NewFile Then Cdiu_DelFile NewFolder, NewFile
Name DoFolder + "\" + DoFile As NewFolder + "\" + NewFile

Finish_Cdiu_RenameFile:

End Sub

Public Sub Cdiu_DelFile(ByVal Folder As String, ByVal FileName As String)

On Error GoTo Finish_Cdiu_DelFile

Kill Folder + "\" + FileName

Finish_Cdiu_DelFile:

End Sub

Function CheckCombo(Beat As Long) As Long

Dim Number As Long, i As Long, j As Long

If Admin = False Then On Error Resume Next

If (Beat * 8 + 7) > UBound(GData) Then ReDim Preserve GData((Beat + 1) * 8)

        For i = Beat To 1 Step -1
                For j = 0 To 6
                        If GData(i * 8 + j) = True Then Number = Number + 1
                Next j
        Next i

CheckCombo = Number
End Function

Public Sub UseModeChange()

If Admin = False Then On Error Resume Next

        If UseMode = "normal" Then
            cmt.SetNormalMode.Checked = True
            cmt.SetSeeMode.Checked = False
            cmt.SetGameMode.Checked = False
            
            cmt.SetGameMode.Enabled = True
            cmt.SetSeeMode.Enabled = True
            cmt.SetNormalMode.Enabled = False
            
        ElseIf UseMode = "game" Then
        
            cmt.SetNormalMode.Checked = False
            cmt.SetSeeMode.Checked = False
            cmt.SetGameMode.Checked = True
            
            cmt.SetGameMode.Enabled = False
            cmt.SetSeeMode.Enabled = True
            cmt.SetNormalMode.Enabled = True
        
            NowCombo = 0
            GameP(0) = 0
            GameP(1) = 0
            GameP(2) = 0
            GameP(3) = 0
            GameP(4) = 0
            Score = 0
            ReDim GameCheck(TotalBeat * 8)
        
        ElseIf UseMode = "see" Then
            cmt.SetNormalMode.Checked = False
            cmt.SetGameMode.Checked = False
            cmt.SetSeeMode.Checked = True
            
            cmt.SetSeeMode.Enabled = False
            cmt.SetNormalMode.Enabled = True
            cmt.SetGameMode.Enabled = True
        End If

End Sub

Public Sub EndTheSong()

If Admin = False Then On Error Resume Next

cma1.NoSound
cmt.PlayOrStop.Caption = IIf(Language = 0, "¼½©ñ", "Play")
cmt.Button.Item(0).Visible = True
cmt.RightOne.Enabled = True
cmt.LeftOne.Enabled = True
cmt.dmenu.Enabled = True
cmt.emenu.Enabled = True
cmt.fmenu.Enabled = True
cmt.gmenu.Enabled = True
cmt.MapMenu.Enabled = True
cmt.PlaySpace.Enabled = True

End Sub

Public Sub Frame4Show()

If Admin = False Then On Error Resume Next

cmt.Frame4.Visible = IIf(cmt.Frame4.Visible = True, False, True)

If cmt.Frame4.Visible = True Then
    cmt.Frame1.Visible = False
    cmt.Frame3.Visible = False
    cmt.Frame5.Visible = False
End If

End Sub

Public Function SetSData(Which As Long, Add As Long)

Dim Tmp As Byte, i As Long

If Admin = False Then On Error Resume Next

Tmp = SData(Add * 8 + Which)

    For i = Add To TotalBeat - 1
        If SData(i * 8 + Which) = Tmp Then
            SData(i * 8 + Which) = Tmp + 1
            If SData(i * 8 + Which) >= 10 And Which < 7 Then SData(i * 8 + Which) = 1
            If SData(i * 8 + Which) > 8 And Which = 7 Then SData(i * 8 + Which) = 0
        Else
            Exit For
        End If
    Next i

If Which = 7 Then cma3.CheckChangeBpm


End Function
