Attribute VB_Name = "cma3"
Dim Fso As New FileSystemObject

Public Sub LoadSlk(LoadFile As String, YNumber As Integer, XNumber As Integer, KData() As String)

Dim Number As Long, TData() As String, i As Long, FileName As String, MaxX As Integer

If Admin = False Then On Error Resume Next

YNumber = 0: XNumber = 0

FileName = Mid(LoadFile, 1, Len(LoadFile) - 4)

        Open LoadFile For Input As #1
            Do
            ReDim Preserve TData(Number)
            Line Input #1, TData(Number)
            
            If InStr(TData(Number), ";Y" + CStr(YNumber + 1)) > 1 Then
            
                XNumber = 0
                YNumber = YNumber + 1
                ReDim Preserve KData(20, YNumber)
            
            End If
            
            For i = 1 To 20
                    If InStr(TData(Number), ";X" + CStr(i)) > 1 Then XNumber = i: If XNumber > MaxX Then MaxX = XNumber: Exit For
            Next i
            
            If (InStr(TData(Number), ";K") > 1) And (XNumber > 0) And (YNumber > 0) Then
                KData(XNumber, YNumber) = Mid(TData(Number), InStrRev(TData(Number), ";K") + 2, Len(TData(Number)))

                    If Left(KData(XNumber, YNumber), 1) = Chr$(34) Then KData(XNumber, YNumber) = Mid(KData(XNumber, YNumber), 2, InStr(2, KData(XNumber, YNumber), Chr$(34)) - 2)
                    
                    If InStr(KData(XNumber, YNumber), ";ER") > 0 Then KData(XNumber, YNumber) = Mid(KData(XNumber, YNumber), 1, InStr(KData(XNumber, YNumber), ";ER") - 1)
            End If

            Number = Number + 1
            Loop Until EOF(1)
        Close #1
        
XNumber = MaxX

End Sub

Public Function FileToStr(FileName As String) As String
Dim txt As TextStream

If Admin = False Then On Error Resume Next

Set txt = Fso.OpenTextFile(FileName, ForReading, False)
FileToStr = txt.ReadAll
txt.Close

End Function

Public Sub LoadSet(LoadFile As String)

Dim chkr() As String

If Admin = False Then On Error Resume Next

    If Fso.FileExists(Left(LoadFile, Len(LoadFile) - 4) + ".set") = True Then
        chkr() = Split(FileToStr(Left(LoadFile, Len(LoadFile) - 4) + ".set"), "ooooooo")
        
        If chkr(0) = "KarMooSetEditorFile" Then
            Singer = chkr(1)
            cmt.Single_Text.Text = chkr(1)
            
            Melody = chkr(2)
            cmt.Melody_Text.Text = chkr(2)
            
            Author = chkr(3)
            cmt.Author_Text.Text = chkr(3)
            
            MusicCode = chkr(4)
            cmt.MusicCode_Text.Text = chkr(4)
            
            level = chkr(5)
            cmt.Level_Text.Text = chkr(5)
            
            If IsNumeric(chkr(6)) = True Then BpmSet(0) = chkr(6)
            If IsNumeric(chkr(6)) = True Then cmt.OK_Bpm.Item(0).Text = chkr(6)
            
            If IsNumeric(chkr(7)) = True Then OffsetSet = chkr(7)
            If IsNumeric(chkr(7)) = True Then cmt.OK_Offset.Text = chkr(7)
            
            If IsNumeric(chkr(8)) = True Then SetRx = CLng(chkr(8)) + 1
    
            For i = 9 To UBound(chkr)
                If IsNumeric(chkr(i)) = True Then
                    ReDim Preserve GData(CLng(chkr(i)) * 8 + 7)
                    GData(CLng(chkr(i)) * 8 + 7) = True
                End If
            Next i
        End If
    End If
End Sub

Public Sub LoadDdr(LoadFile As String, YNumber As Integer, XNumber As Integer, LData() As String)

Dim TData() As String, Number As Long, KData() As String
YNumber = 3

If Admin = False Then On Error Resume Next

        Open LoadFile For Input As #1
            Do
            ReDim Preserve TData(Number)
            Line Input #1, TData(Number)
            
            If Number = 3 Then
                KData = Split(TData(Number), Chr$(34))
                
                Melody = KData(1)
                cmt.Melody_Text = Melody
                
                Singer = KData(3)
                cmt.Single_Text = Singer
                
                level = KData(5)
                cmt.Level_Text = level
                
                If IsNumeric(KData(7)) = True Then cmt.OK_Bpm.Item(0).Text = KData(7)
                If IsNumeric(KData(7)) = True Then BpmSet(0) = cmt.OK_Bpm.Item(0).Text
                
                Ddr_devide = KData(9)
                
                If IsNumeric(KData(11)) = True Then cmt.OK_Offset.Text = KData(11)
                If IsNumeric(KData(11)) = True Then OffsetSet = cmt.OK_Offset.Text
            
            ElseIf Number = 4 Then
                KData = Split(TData(Number), Chr$(34))
                
                Ddr_image = KData(5)
                
            ElseIf Number = 5 Then
                KData = Split(TData(Number), "<desc>")
                Ddr_Desc = KData(1)
                Ddr_Desc = Replace(Ddr_Desc, "</desc>", "")
                
            End If
            
                If InStr(TData(Number), "time=") > 1 Then
                    ReDim Preserve LData(20, YNumber)
                    KData = Split(TData(Number), Chr$(34))
                    LData(4, YNumber) = KData(1)
                        
                        If InStr(KData(3), "G") = 1 Then
                            LData(5, YNumber) = "s"
                            If InStr(KData(3), "I") > 0 Then LData(5, YNumber) = LData(5, YNumber) + ",f"
                            
                        ElseIf InStr(KData(3), "G") > 1 Then
                            LData(5, YNumber) = "s"
                            If InStr(KData(3), "I") > 0 Then LData(5, YNumber) = LData(5, YNumber) + ",f"
                            
                            LData(1, YNumber) = CStr(Fix(CInt(KData(1)) / 16))
                            LData(2, YNumber) = CStr(Fix((CInt(KData(1)) Mod 16) / 4))
                            LData(3, YNumber) = CStr((CInt(KData(1)) Mod 16) Mod 4)
                            
                            YNumber = YNumber + 1
                            ReDim Preserve LData(20, YNumber)
                        Else
                            LData(5, YNumber) = "n"
                            If InStr(KData(3), "I") > 0 Then LData(5, YNumber) = LData(5, YNumber) + ",f"
                        End If

                            If Left(KData(3), 1) = "A" Then LData(6, YNumber) = "7"
                            If Left(KData(3), 1) = "B" Then LData(6, YNumber) = "4"
                            If Left(KData(3), 1) = "C" Then LData(6, YNumber) = "1"
                            If Left(KData(3), 1) = "D" Then LData(6, YNumber) = "9"
                            If Left(KData(3), 1) = "E" Then LData(6, YNumber) = "6"
                            If Left(KData(3), 1) = "F" Then LData(6, YNumber) = "3"
                        
                            LData(1, YNumber) = CStr(Fix(CInt(KData(1)) / 16))
                            LData(2, YNumber) = CStr(Fix((CInt(KData(1)) Mod 16) / 4))
                            LData(3, YNumber) = CStr((CInt(KData(1)) Mod 16) Mod 4)

                    YNumber = YNumber + 1
                End If
            Number = Number + 1
            Loop Until EOF(1)
        Close #1

YNumber = YNumber - 1

End Sub

Public Sub LoadToGData(KData, YNumber As Integer)

Dim i As Integer, OKNumber As Integer, a As Long

If Admin = False Then On Error Resume Next

        ReDim GData(0)
        
        For i = 3 To YNumber
                
                'If KData(4, i) <> "" Then ReDim Preserve GData(CLng(KData(4, i)) * 8)
                
                If InStr(KData(5, i), "n") > 0 Then
                    ReDim Preserve GData(CLng(KData(4, i)) * 8 + 5)
                
                    Select Case KData(6, i)
                        Case 9
                            GData(CLng(KData(4, i)) * 8 + 0) = True
                        Case 6
                            GData(CLng(KData(4, i)) * 8 + 1) = True
                        Case 3
                            GData(CLng(KData(4, i)) * 8 + 2) = True
                        Case 7
                            GData(CLng(KData(4, i)) * 8 + 3) = True
                        Case 4
                            GData(CLng(KData(4, i)) * 8 + 4) = True
                        Case 1
                            GData(CLng(KData(4, i)) * 8 + 5) = True
                    End Select
                End If
                
                If InStr(KData(5, i), "s") > 0 Then
                    ReDim Preserve GData(CLng(KData(4, i)) * 8 + 6)
                            GData(CLng(KData(4, i)) * 8 + 6) = True
                End If
                
                If InStr(KData(5, i), "f") > 0 Then
                    ReDim Preserve GData(CLng(KData(4, i)) * 8 + 7)
                            GData(CLng(KData(4, i)) * 8 + 7) = True
                End If
        Next i

WhenOpen

End Sub

Public Sub WhenOpen()

If Admin = False Then On Error Resume Next

cmt.SaveSlkButton.Enabled = True
cmt.SaveKbe.Enabled = True
cmt.SaveDdr.Enabled = True
cmt.SaveCbe.Enabled = True
cmt.HighSave.Enabled = True
cmt.SaveAsCbg.Enabled = True
cmt.OpenSlk.Enabled = False
cmt.OpenKbe.Enabled = False
cmt.OpenCbe.Enabled = False
cmt.OpenDdr.Enabled = False

End Sub

Public Function ClearName(Clear As String) As String

Dim TName() As Byte, i As Long

If Admin = False Then On Error Resume Next

        TName = StrConv(Clear, vbFromUnicode)
        For i = 1 To UBound(TName)
            If TName(i) = 0 Then
                ReDim Preserve TName(i - 1)
                Exit For
            End If
        Next
        ClearName = StrConv(TName, vbUnicode)

End Function

Public Sub CbeToSlk(FileName As String, Optional SaveToDdr As Boolean, Optional SaveFileName As String)

Dim Number As Integer, KData() As String, i As Long, j As Long, Note As String, SaveFile As String

If Admin = False Then On Error Resume Next

ReDim KData(20, TotalBeat): Number = 3

KData(1, 1) = "葆蛤": KData(2, 1) = "夢(1/4)": KData(3, 1) = "1/16": KData(4, 1) = "嬪纂": KData(5, 1) = "翕濛": KData(6, 1) = "Key":
KData(1, 2) = "int": KData(2, 2) = "int": KData(3, 2) = "int": KData(4, 2) = "int": KData(5, 2) = "enum(n,s,f)": KData(6, 2) = "string":

        For i = 0 To TotalBeat
                
                For j = 0 To 7
                        Note = ""
                        Select Case j
                            Case 0: Note = "9"
                            Case 1: Note = "6"
                            Case 2: Note = "3"
                            Case 3: Note = "7"
                            Case 4: Note = "4"
                            Case 5: Note = "1"
                        End Select
                        
                        If j = 7 Then
                                If GData(i * 8 + j) = True Then KData(5, Number - 1) = KData(5, Number - 1) + ",f"
                                GoTo EndJ
                        End If
                                
                        If (i * 8 + j) > UBound(GData) Then Exit For
                        
                        If GData(i * 8 + j) = True Then
                            
                            KData(1, Number) = CStr(Fix(i / 16))
                            KData(2, Number) = CStr(Fix((i Mod 16) / 4))
                            KData(3, Number) = CStr((i Mod 16) Mod 4)
                            KData(4, Number) = CStr(i)
                            
                                If j <> 6 Then
                                    KData(5, Number) = "n"
                                    KData(6, Number) = Note
                                Else
                                    KData(5, Number) = "s"
                                End If

                            Number = Number + 1

                        End If
                        
                Next j
EndJ:
                
        Next i

        FileName = ClearName(FileName)

If SaveToDdr = False Then
    SaveSlkSub "CbeToSlk", FileName, Number - 1, 6, KData, 0
Else
    SaveDdr KData, SaveFileName, Number - 1
End If

End Sub

Public Function CheckFileBack(FileName As String, BackName As String) As String

If Admin = False Then On Error Resume Next

If Right(FileName, 4) = BackName Then
        CheckFileBack = FileName
Else
        CheckFileBack = FileName + BackName
End If

End Function

Public Sub SaveSlkSub(Which As String, FileName As String, YNumber As Integer, XNumber As Integer, GData() As String, StartNumber As Integer, Optional NotDisplay As Boolean)

Dim i As Integer, Data1() As Byte, Data2() As Byte, u As Integer, NumberOrString As String, ArrayNumber As Integer

If Admin = False Then On Error Resume Next

    Select Case Which
        Case "CbeToSlk"
            ReDim Data1(UBound(Slk4))
            ReDim Data2(UBound(Slk3))
            cq.IcyCopyMemory ByVal VarPtr(Data1(0)), ByVal VarPtr(Slk4(0)), UBound(Slk4) + 1
            cq.IcyCopyMemory ByVal VarPtr(Data2(0)), ByVal VarPtr(Slk3(0)), UBound(Slk3) + 1
        'Data1 = LoadResData("SLKHEADER", "SLK"): Data2 = LoadResData("SLK2HEADER", "SLK")
        Case "HighSave"
            ReDim Data1(UBound(Slk1))
            ReDim Data2(UBound(Slk2))
            cq.IcyCopyMemory ByVal VarPtr(Data1(0)), ByVal VarPtr(Slk1(0)), UBound(Slk1) + 1
            cq.IcyCopyMemory ByVal VarPtr(Data2(0)), ByVal VarPtr(Slk2(0)), UBound(Slk2) + 1
        'Data1 = LoadResData("LIST1", "SLK"): Data2 = LoadResData("LIST2", "SLK")
    End Select
   
FileName = CheckFileBack(FileName, ".slk")
If FileName = ".slk" Then Exit Sub

DelFile FileName

        Open FileName For Binary As #1
            Put #1, 1, Data1
            Put #1, , "B;Y" + CStr(YNumber) + ";X" + CStr(XNumber) + ";D0 0 " + CStr(YNumber - 1) + " " + CStr(XNumber - 1) + vbCrLf
            Put #1, , Data2
    
                For i = 1 To YNumber + ArrayNumber
                        If (GData(1, i) = "") And (Which = "BGMake") Then GData(1, i) = GData(2, i)
                        
                        For u = 1 To 20
                            NumberOrString = Chr$(34)
                                If IsNumeric(GData(u, i)) Then NumberOrString = ""
                                If XNumber >= u Then Put #1, , "C;Y" + CStr(i + StartNumber) + ";X" + CStr(u) + ";K" + NumberOrString + GData(u, i) + NumberOrString + vbCrLf
                        Next u
                Next i
    
            Put #1, , "E" + vbCrLf
        Close #1
        
        If NotDisplay <> True Then MsgBox IIf(Language = 0, "儲存完成", "Save Success"), 0, IIf(Language = 0, "系統訊息", "System Info")

End Sub

Public Sub DelFile(ByVal FileName As String)

On Error GoTo Finish_DelFile

Kill FileName

Finish_DelFile:

End Sub

Public Function FindFileName(ByVal FileName As String, Optional ByVal Which As Integer) As String

If Admin = False Then On Error Resume Next

        If Len(FileName) > 0 Then
            FindFileName = Mid(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
                If Which = 1 Then
                        If InStr(FindFileName, ".") > 0 Then FindFileName = Split(FindFileName, ".")(0)
                ElseIf Which = 2 Then
                        If InStr(FindFileName, ".") > 0 Then FindFileName = Split(FindFileName, ".")(1)
                ElseIf Which = 3 Then
                    FileName = Replace(FileName, "\" + FindFileName, "")
                    FindFileName = Replace(FileName, FindFileName, "")
                ElseIf Which = 4 Then
                    FindFileName = Mid(FileName, 1, InStr(FileName, "\"))
                ElseIf Which = 5 Then
                    FindFileName = Dir(Replace(FileName, "\" + FindFileName(FileName), ""))
                End If
        End If
        
End Function

Public Sub LoadKbe(LoadFile As String)

Dim SigT As String * 20, SigT2 As String * 22, LenTPData As Long, KData() As Byte, FileStart As Integer

If Admin = False Then On Error Resume Next

    Open LoadFile For Binary As #1
        Get #1, 1, SigT
        Get #1, 1, SigT2
            
        If SigT = "KarMooBeatEditorFile" Then FileStart = 21
        If SigT2 = "KarMooBeatUpEditorFile" Then FileStart = 23
            
        Get #1, FileStart, LenTPData
            ReDim GData(LenTPData / 7 * 8)
            ReDim KData(LenTPData)
            TotalBeat = LenTPData / 7
        Get #1, , KData
    Close #1

    For i = 0 To TotalBeat - 1
        GData(i * 8 + 0) = IIf(KData(i * 7 + 4) = 1, True, False)
        GData(i * 8 + 1) = IIf(KData(i * 7 + 5) = 1, True, False)
        GData(i * 8 + 2) = IIf(KData(i * 7 + 6) = 1, True, False)
        GData(i * 8 + 3) = IIf(KData(i * 7 + 1) = 1, True, False)
        GData(i * 8 + 4) = IIf(KData(i * 7 + 2) = 1, True, False)
        GData(i * 8 + 5) = IIf(KData(i * 7 + 3) = 1, True, False)
        GData(i * 8 + 6) = IIf(KData(i * 7 + 7) = 1, True, False)
    Next i
    
    WhenOpen
    
End Sub

Public Sub CbeToKbe(LoadFile As String)

Dim i As Long, KData() As Byte

If Admin = False Then On Error Resume Next

LoadFile = CheckFileBack(LoadFile, ".kbe")
If LoadFile = ".kbe" Then Exit Sub

    ReDim KData(TotalBeat * 7)

    For i = 0 To TotalBeat - 1
        If GData(i * 8 + 0) = True Then KData(i * 7 + 4) = 1
        If GData(i * 8 + 1) = True Then KData(i * 7 + 5) = 1
        If GData(i * 8 + 2) = True Then KData(i * 7 + 6) = 1
        If GData(i * 8 + 3) = True Then KData(i * 7 + 1) = 1
        If GData(i * 8 + 4) = True Then KData(i * 7 + 2) = 1
        If GData(i * 8 + 5) = True Then KData(i * 7 + 3) = 1
        If GData(i * 8 + 6) = True Then KData(i * 7 + 7) = 1
    Next i

LoadFile = ClearName(LoadFile)

DelFile LoadFile

    Open LoadFile For Binary As #1
    Put #1, 1, "KarMooBeatEditorFile"
    Put #1, , CLng(UBound(KData))
    Put #1, , KData
    Close #1

        MsgBox IIf(Language = 0, "儲存完成", "Save Success"), 0, IIf(Language = 0, "系統訊息", "System Info")

End Sub

Public Sub CbeToDdr(LoadFile As String)

If Admin = False Then On Error Resume Next

CbeToSlk LoadFile, True, LoadFile

End Sub

Public Sub SaveDdr(KData() As String, SaveFileDr As String, YNumber As Integer)

Dim TData() As String

If Admin = False Then On Error Resume Next

SaveFileDr = CheckFileBack(SaveFileDr, ".ddr")
If SaveFileDr = ".ddr" Then Exit Sub

AddBackArray TData, "<?xml version=" + Chr$(34) + "1.0" + Chr$(34) + "?>"
AddBackArray TData, ""
AddBackArray TData, "<ddr>"
AddBackArray TData, "    <head name=" + Chr$(34) + Melody + Chr$(34) + " author=" + Chr$(34) + Singer + Chr$(34) + " level=" + Chr$(34) + level + Chr$(34) + " bpm=" + Chr$(34) + BpmSet(0) + Chr$(34) + " devide=" + Chr$(34) + Ddr_devide + Chr$(34) + " offset=" + Chr$(34) + OffsetSet + Chr$(34) + ">"
AddBackArray TData, "        <music length=" + Chr$(34) + CStr(SoundL) + Chr$(34) + " file=" + Chr$(34) + SongPath + Chr$(34) + " image=" + Chr$(34) + Ddr_image + Chr$(34) + " />"
AddBackArray TData, "        <desc>" + Ddr_Desc + "</desc>"
AddBackArray TData, "    </head>"

        For i = 1 To YNumber
                
                Select Case KData(6, i)
                    Case "7": Note = "A"
                    Case "4": Note = "B"
                    Case "1": Note = "C"
                    Case "9": Note = "D"
                    Case "6": Note = "E"
                    Case "3": Note = "F"
                End Select
                
            If (KData(5, i) = "s") Or (KData(5, i) = "n,s") Or (KData(5, i) = "s,n") Or (KData(5, i) = "s,f") Or (KData(5, i) = "f,s") Or (KData(5, i) = "s,n,f") Or (KData(5, i) = "s,f,n") Or (KData(5, i) = "n,s,f") Or (KData(5, i) = "n,f,s") Or (KData(5, i) = "f,n,s") Or (KData(5, i) = "f,s,n") Then AddBackArray TData, "    <Tick time=" + Chr$(34) + CStr(KData(4, i)) + Chr$(34) + " keys=" + Chr$(34) + "G" + Chr$(34) + " />"
            If (KData(5, i) = "n") Or (KData(5, i) = "n,s") Or (KData(5, i) = "s,n") Or (KData(5, i) = "n,f") Or (KData(5, i) = "f,n") Or (KData(5, i) = "s,n,f") Or (KData(5, i) = "s,f,n") Or (KData(5, i) = "n,s,f") Or (KData(5, i) = "n,f,s") Or (KData(5, i) = "f,n,s") Or (KData(5, i) = "f,s,n") Then AddBackArray TData, "    <Tick time=" + Chr$(34) + CStr(KData(4, i)) + Chr$(34) + " keys=" + Chr$(34) + Note + Chr$(34) + " />"

        Next i
        
AddBackArray TData, "</ddr>"

DelFile SaveFileDr

        Open SaveFileDr For Binary As #1
            For i = 0 To UBound(TData)
                Put #1, , TData(i)
                    If i <> UBound(TData) Then Put #1, , vbCrLf
            Next i
        Close #1
        
        MsgBox IIf(Language = 0, "儲存完成", "Save Success"), 0, IIf(Language = 0, "系統訊息", "System Info")


End Sub

Public Sub AddBackArray(GData As Variant, ByVal AddWord As Variant)

Dim Number As Integer

If Admin = False Then On Error Resume Next

Number = GetUBound(GData) + 1
ReDim Preserve GData(Number)
GData(Number) = AddWord

End Sub

Public Function GetUBound(GData As Variant) As Integer

On Error Resume Next

GetUBound = -1
GetUBound = UBound(GData)

End Function

Public Sub GameFile(LoadFile As String)

Dim NowFolder As String, FileName As String

NowFolder = Replace(LoadFile, FindFileName(LoadFile), "")
FileName = FindFileName(LoadFile)
FileName = ClearName(FileName)

CbeOut LoadFile, False, True
cma2.Cdiu_File "Rename", NowFolder, FileName + ".cdiu", NowFolder, FileName + ".cbg"

End Sub

Public Sub CbeOut(LoadFile As String, Optional NotShow As Boolean, Optional SaveOgg As Boolean)

Dim FileName As String, NowFolder As String, Tmp As String * 512, SongFile() As Byte, RandomCode As Long

If Admin = False Then On Error Resume Next

NowFolder = Replace(LoadFile, FindFileName(LoadFile), "")
FileName = FindFileName(LoadFile)
FileName = ClearName(FileName)
If LoadFile = ".cdiu" Then Exit Sub

    cma5.CreateDir TempPath + "\cma\"

    If SaveOgg = True Then
            Open SongPath For Binary As #1
                ReDim SongFile(FileLen(SongPath) - 1)
                Get #1, 1, SongFile
            Close #1
    
            Open TempPath + "\cma\" + "song.ogg" For Binary As #1
                Put #1, 1, SongFile
            Close #1
            
            Randomize
            'RandomCode = Rnd() * 2555
            RandomCode = Fix(25555555 * Rnd)
            
            Open TempPath + "\cma\" + "RandomCode.cbe" For Binary As #1
                Put #1, 1, RandomCode
            Close #1
            
    End If


    Open TempPath + "\cma\" + "maindata.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileData"
    Put #1, , CLng(UBound(GData))
    Put #1, , GData
    Close #1
    
    Open TempPath + "\cma\" + "sdata.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileSata"
    Put #1, , CLng(UBound(SData))
    Put #1, , SData
    Close #1
    
    Open TempPath + "\cma\" + "mdata.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileMata"
    Put #1, , CLng(UBound(MData))
    Put #1, , MData
    Close #1
    
    Open TempPath + "\cma\" + "infodata.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileInfo"
    Put #1, , CurTime
    Put #1, , ShitTime
    Put #1, , CurrPos
    Put #1, , SetRx
    Put #1, , CLng(cmt.Times.value)
    Tmp = Singer
    Put #1, , Tmp
    Tmp = Melody
    Put #1, , Tmp
    Tmp = Author
    Put #1, , Tmp
    Tmp = level
    Put #1, , Tmp
    Tmp = MusicCode
    Put #1, , Tmp
    Tmp = UseMode
    Put #1, , Tmp
    Tmp = BpmSet(0)
    Put #1, , Tmp
    Tmp = OffsetSet
    Put #1, , Tmp
    Put #1, , CBT
    Put #1, , OffSet
    Tmp = SongPath
    Put #1, , Tmp
    Put #1, , ChooseBackGround
    Tmp = BpmSet(1)
    Put #1, , Tmp
    Tmp = BpmSet(2)
    Put #1, , Tmp
    Tmp = BpmSet(3)
    Put #1, , Tmp
    Tmp = BpmSet(4)
    Put #1, , Tmp
    Tmp = BpmSet(5)
    Put #1, , Tmp
    Tmp = BpmSet(6)
    Put #1, , Tmp
    Tmp = BpmSet(7)
    Put #1, , Tmp
    Tmp = BpmSet(8)
    Put #1, , Tmp
    Close #1
    

        Enrypt_12 TempPath + "cma\", FileName, NowFolder + "\"
        
        DeleteDir TempPath + "cma\"

    If NotShow = False Then MsgBox IIf(Language = 0, "儲存完成", "Save Success"), 0, IIf(Language = 0, "系統訊息", "System Info")
    
End Sub

Public Sub LoadCbe(LoadFile As String, Optional LoadGame As Boolean)

Dim NowFolder As String, SigT As String * 22, LenTPData As Long, SigT2 As String * 22, RealName As String * 512, Tmp As Long, CBT2 As Single, Offset2 As Double, ChooseBackGroundA As Long, NoUse As Long, NoUse2 As Integer

If Admin = False Then On Error Resume Next

NowFolder = Replace(LoadFile, FindFileName(LoadFile), "")

        Decrypt_12 LoadFile, 0

    If LoadGame = False Then
    
        If Fso.FileExists(NowFolder + "\cma\" + "maindata.cbe") = True Then
            Open NowFolder + "\cma\" + "maindata.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileData" Then
                Get #1, , LenTPData
                    ReDim GData(LenTPData)
                    TotalBeat = (LenTPData / 8)
                Get #1, , GData
            End If
            Close #1
        End If
        
        If Fso.FileExists(NowFolder + "\cma\" + "sdata.cbe") = True Then
            Open NowFolder + "\cma\" + "sdata.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileSata" Then
                Get #1, , LenTPData
                    ReDim SData(LenTPData)
                Get #1, , SData
            End If
            Close #1
        End If

        If Fso.FileExists(NowFolder + "\cma\" + "mdata.cbe") = True Then
            Open NowFolder + "\cma\" + "mdata.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileMata" Then
                Get #1, , LenTPData
                    ReDim MData(LenTPData)
                Get #1, , MData
            End If
            Close #1
        End If
        
        If Fso.FileExists(NowFolder + "\cma\" + "infodata.cbe") = True Then
            Open NowFolder + "\cma\" + "infodata.cbe" For Binary As #1
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
                UseMode = Trim(RealName)
                
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
    Else

        If Fso.FileExists(NowFolder + "\cma\" + "RandomCode.cbe") = True Then
            Open NowFolder + "\cma\" + "RandomCode.cbe" For Binary As #1
            Get #1, 1, CodeA
            Close #1
        End If
    
    If Fso.FileExists(NowFolder + "\cma\" + "infodata.cbe") = True Then
            Open NowFolder + "\cma\" + "infodata.cbe" For Binary As #1
            Get #1, 1, SigT2
            If SigT2 = "CdiuBeatEditorFileInfo" Then
                Get #1, , NoUse
                Get #1, , NoUse2
                Get #1, , NoUse
                Get #1, , NoUse
                Get #1, , NoUse
                
                Get #1, , RealName
                GameSinger = Trim(RealName)
        
                Get #1, , RealName
                GameMelody = Trim(RealName)
                
                Get #1, , RealName
                GameAuthor = Trim(RealName)
            End If
            Close #1
        End If
        
    End If
    
        DeleteDir NowFolder + "\cma\"
        
    If LoadGame = False Then

        cmt.openmusic.Enabled = False
        cmt.NewFile.Enabled = True
        cmt.setting.Enabled = True
        cmt.ProSetting.Enabled = True
        WhenOpen
        CheckDFrame2
        cmt.ShowOrHide.Enabled = True
        cmt.BMenu.Enabled = True
        cmt.CMenu.Enabled = True
        cmt.dmenu.Enabled = True
        cmt.emenu.Enabled = True
        cmt.fmenu.Enabled = True
        cmt.gmenu.Enabled = True
        cmt.MapMenu.Enabled = True
        cmt.Frame2.Visible = True
        
        cma3.CheckChangeBpm
        cma2.UseModeChange
        
        If ChooseBackGroundA = 9 Then
            ChooseBackGround = 1
            cma4.ChangeMapByUser 9
        Else
            ChooseBackGround = ChooseBackGroundA
        End If
        
    End If

End Sub

Public Sub CheckChangeBpm()

Dim i As Long

ChangeBpm = False

    For i = 0 To UBound(SData) / 8 - 1
        If SData(i * 8 + 7) <> 0 Then ChangeBpm = True: Exit For
    Next i

If ChangeBpm = True Then FindBeatTime PData

End Sub

Public Sub AutoSave()

If Admin = False Then On Error Resume Next

CbeOut App.Path + "\Temp", True

End Sub

Public Sub LoadCbeSong(LoadFile As String)

Dim NowFolder As String, SigT As String * 22, LenTPData As Long, SigT2 As String * 22, RealName As String * 512, Tmp As Long, TempI As Integer, TempS As Single, TempD As Double

If Admin = False Then On Error Resume Next

NowFolder = Replace(LoadFile, FindFileName(LoadFile), "")

        Decrypt_12 LoadFile, 0
    
    Open NowFolder + "\cma\" + "infodata.cbe" For Binary As #1
    Get #1, 1, SigT2
    If SigT2 = "CdiuBeatEditorFileInfo" Then
        Get #1, , Tmp
        Get #1, , TempI
        Get #1, , Tmp
        Get #1, , Tmp
        Get #1, , Tmp
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , RealName
        Get #1, , TempS
        Get #1, , TempD
        Get #1, , RealName
        SongPath = Trim(RealName)
        
    End If
    Close #1

        DeleteDir NowFolder + "\cma\"

End Sub

Public Sub AutoLoadDo()

If Admin = False Then On Error Resume Next

cma3.LoadCbeSong App.Path + "\Temp.cdiu"
cma1.OpenSound SongPath, True
cmt.Frame2.Visible = True

End Sub

Public Sub NewFileDo()

Dim i As Long

If Admin = False Then On Error Resume Next

cmt.NewFile.Enabled = False
cmt.openmusic.Enabled = True
cmt.OpenSlk.Enabled = False
cmt.SaveSlkButton.Enabled = False
cmt.OpenKbe.Enabled = False
cmt.SaveKbe.Enabled = False
cmt.OpenCbe.Enabled = False
cmt.SaveCbe.Enabled = False
cmt.OpenDdr.Enabled = False
cmt.SaveDdr.Enabled = False
cmt.SaveAsCbg.Enabled = False
cmt.HighSave.Enabled = False
cmt.BMenu.Enabled = False
cmt.CMenu.Enabled = False
cmt.dmenu.Enabled = False
cmt.emenu.Enabled = False
cmt.fmenu.Enabled = False
cmt.gmenu.Enabled = False
cmt.MapMenu.Enabled = True
cmt.Frame1.Visible = False
cmt.Frame2.Visible = False
cmt.Frame3.Visible = False
cmt.Frame4.Visible = False
cmt.Frame5.Visible = False
UseMode = "see"
Mode = "stop"
cmt.Times.value = 0
CurTime = 0
CurrPos = 0
SetRx = 1
cma1.StopSound
cmt.OK_Bpm.Item(0).Text = 150
cmt.OK_Bpm.Item(1).Text = 140
cmt.OK_Bpm.Item(2).Text = 130
BpmSet(0) = 150
BpmSet(1) = 75
BpmSet(2) = 130
cmt.OK_Offset.Text = 0.7
OffsetSet = 0.7
'cma4.UnloadD3D
'Inited = False
ReDim GData(14)
ReDim SData(14)

    For i = 0 To 6
        SData(i) = 4
    Next i

cmt.SetNormalMode.Enabled = True
cmt.SetSeeMode.Enabled = False
cmt.SetSeeMode.Checked = False
cmt.Single_Text = "Singer"
Singer = "Singer"
cmt.Melody_Text = "Melody"
Melody = "Melody"
cmt.Author_Text = "Author"
Author = "Author"
cmt.Level_Text = "3"
level = "3"
cmt.MusicCode_Text = "km001"
MusicCode = "km001"
ChooseBackGround = 1
cma2.EndTheSong

End Sub

Public Sub OpenSlkDo()

Dim SlkPath As String, KData() As String, YNumber As Integer, NowFolder As String, FileName As String

If Admin = False Then On Error Resume Next

SlkPath = cma2.OpenFile("Slk", "Open")
If SlkPath <> "" Then SlkPath = ClearName(SlkPath)


NowFolder = Replace(SlkPath, cma3.FindFileName(SlkPath), "")
FileName = cma3.FindFileName(SlkPath)

    If SlkPath <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then
        cma3.LoadSlk SlkPath, YNumber, 0, KData
        cma3.LoadToGData KData, YNumber
        cma3.LoadSet SlkPath
        cma3.AutoSave
    End If
    
End Sub

Public Sub OpenCbeDo()

Dim SlkPath As String, KData() As String, YNumber As Integer, NowFolder As String, FileName As String

If Admin = False Then On Error Resume Next

SlkPath = cma2.OpenFile("cdiu", "Open")
If SlkPath <> "" Then SlkPath = ClearName(SlkPath)

NowFolder = Replace(SlkPath, cma3.FindFileName(SlkPath), "")
FileName = cma3.FindFileName(SlkPath)

If SlkPath <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then cma3.LoadCbe SlkPath: cma3.AutoSave

End Sub

Public Sub OpenDdrDo()

Dim SlkPath As String, LData() As String, YNumber As Integer, NowFolder As String, FileName As String

If Admin = False Then On Error Resume Next

SlkPath = cma2.OpenFile("Ddr", "Open")
If SlkPath <> "" Then SlkPath = ClearName(SlkPath)

NowFolder = Replace(SlkPath, cma3.FindFileName(SlkPath), "")
FileName = cma3.FindFileName(SlkPath)

If SlkPath <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then
    cma3.LoadDdr SlkPath, YNumber, 0, LData
    cma3.LoadToGData LData, YNumber
    cma3.AutoSave
End If

End Sub

Public Sub OpenKbeDo()

Dim SlkPath As String, NowFolder As String, FileName As String

If Admin = False Then On Error Resume Next

SlkPath = cma2.OpenFile("Kbe", "Open")

If SlkPath <> "" Then SlkPath = ClearName(SlkPath)


NowFolder = Replace(SlkPath, cma3.FindFileName(SlkPath), "")
FileName = cma3.FindFileName(SlkPath)

If SlkPath <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then cma3.LoadKbe SlkPath: cma3.LoadSet SlkPath: cma3.AutoSave

End Sub

Public Sub LoadASL()

Dim KData() As String, YNumber As Integer, XNumber As Integer, i As Long, Check As Boolean

If Admin = False Then On Error Resume Next

LoadSlk ASL, YNumber, XNumber, KData

        For i = 1 To YNumber
            If Right(KData(3, i), 4) = ".ogg" Or Right(KData(3, i), 4) = ".kcb" Or Right(KData(3, i), 4) = ".tbm" Or Right(KData(3, i), 4) = ".abm" Then
                If Mid(KData(3, i), 1, InStr(KData(3, i), ".") - 1) = MusicCode Then
                    KData(1, i) = Singer
                    KData(2, i) = Melody
                    KData(6, i) = BpmSet(0)
                    Check = True
                End If
            End If
        Next i

    If Check = True Then SaveSlkSub "HighSave", ASL, YNumber, 13, KData, 0, True

End Sub

Public Sub LoadBUL()

Dim KData() As String, YNumber As Integer, XNumber As Integer, Check As Boolean

If Admin = False Then On Error Resume Next

LoadSlk BUL, YNumber, XNumber, KData

        For i = 1 To YNumber
            If Right(KData(3, i), 4) = ".ogg" Or Right(KData(3, i), 4) = ".kcb" Or Right(KData(3, i), 4) = ".tbm" Or Right(KData(3, i), 4) = ".abm" Then
                If Mid(KData(3, i), 1, InStr(KData(3, i), ".") - 1) = MusicCode Then
                    KData(1, i) = Singer
                    KData(2, i) = Melody
                    KData(4, i) = BpmSet(0)
                    KData(5, i) = level
                    KData(6, i) = OffsetSet
                    Check = True
                End If
            End If
        Next i

    If Check = True Then SaveSlkSub "HighSave", BUL, YNumber, 8, KData, 0, True

End Sub

Public Sub HighSaveDo(SaveFile As String)

Dim SaveCbeFile As String

If Admin = False Then On Error Resume Next

        If Fso.FileExists(ASL) = True Then LoadASL
        If Fso.FileExists(BUL) = True Then LoadBUL
        
SaveCbeFile = Replace(SaveFile, ".slk", "")
cma3.CbeOut SaveCbeFile, True

CbeToSlk SaveFile

End Sub

Public Sub DelANote()

Dim i As Long

If Admin = False Then On Error Resume Next

SaveUnDo

For i = 0 To UBound(GData)
    GData(i) = False
Next i

End Sub

