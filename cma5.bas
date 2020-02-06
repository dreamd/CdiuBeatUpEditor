Attribute VB_Name = "cma5"
Dim Fso As New FileSystemObject

Public Sub MakeItUpDo()

Dim First_D As Long, Last_D As Long, First_B As Long, Last_B As Long, i As Long, u As Long
SaveUnDo

On Error GoTo EndDelete

    First_D = SaveSelect(1) Mod 8
    Last_D = SaveSelect(UBound(SaveSelect)) Mod 8
    First_B = (SaveSelect(1) - First_D) / 8
    Last_B = (SaveSelect(UBound(SaveSelect)) - Last_D) / 8
    
    For i = Last_B To First_B
        For u = Last_D To First_D
           MData(i * 8 + u) = True
           GData(i * 8 + u) = True
        Next u
    Next i
    
EndDelete:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False
    cma3.AutoSave
    
End Sub

Public Sub MakeItOutDo()

Dim First_D As Long, Last_D As Long, First_B As Long, Last_B As Long, i As Long, u As Long
SaveUnDo

On Error GoTo EndDelete

    First_D = SaveSelect(1) Mod 8
    Last_D = SaveSelect(UBound(SaveSelect)) Mod 8
    First_B = (SaveSelect(1) - First_D) / 8
    Last_B = (SaveSelect(UBound(SaveSelect)) - Last_D) / 8
    
    For i = Last_B To First_B + 1
        For u = Last_D To First_D
           MData(i * 8 + u) = False
           GData(i * 8 + u) = False
        Next u
    Next i
    
EndDelete:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False
    cma3.AutoSave
    
End Sub

Public Sub AutoDoSpace(DoBeat As Long, HowManyBeat As Long)

Dim i As Long

If Admin = False Then On Error Resume Next

GData(DoBeat * 8 + 6) = True

ReDim Preserve GData((TotalBeat + 1) * 8 + 7)

        For i = DoBeat To TotalBeat
            GData(i * 8 + 6) = False
        Next i

        For i = DoBeat To TotalBeat Step HowManyBeat
            GData(i * 8 + 6) = True
        Next i

End Sub

Public Sub AutoDelSpace(DoBeat As Long)

Dim i As Long

If Admin = False Then On Error Resume Next

ReDim Preserve GData((TotalBeat + 1) * 8 + 7)

        For i = DoBeat + 1 To TotalBeat
            GData(i * 8 + 6) = False
        Next i


End Sub


Public Sub SpaceBack16(DoBeat As Long)

Dim i As Long

If Admin = False Then On Error Resume Next

ReDim Preserve GData((TotalBeat + 1) * 8 + 7)

        For i = TotalBeat To DoBeat + 1 Step -1
            If GData(i * 8 + 6) = True Then
                GData(i * 8 + 6) = False
                If i + 16 < TotalBeat Then GData((i + 16) * 8 + 6) = True
            End If
        Next i


End Sub


Public Function NextSpace(DoBeat As Long) As Long

Dim i As Long

If Admin = False Then On Error Resume Next

        For i = DoBeat + 3 To TotalBeat
                If GData(i * 8 + 6) = True Then NextSpace = i: Exit Function
        Next i

End Function

Public Function CheckSmall(HData) As Long

Dim i As Long, Tmp As Long

If Admin = False Then On Error Resume Next

Tmp = 1000000
    For i = 1 To UBound(HData)
        If HData(i) < Tmp Then Tmp = HData(i)
    Next i

CheckSmall = Tmp

End Function

Public Sub GoCopy(Optional Cut As Boolean)
On Error GoTo EndCopy

Dim CopyData As String, Base As Long, j As Long, Tmp As Long, l As Long, o As Long

    Base = Fix((CheckSmall(SaveSelect)) / 8) * 8
    
    CopyData = "BBQ"
    For o = 1 To UBound(SaveSelect)
        CopyData = CopyData + "," + CStr(SaveSelect(o) - Base)
    Next
    
    Clipboard.SetText CopyData
    
EndCopy:

If Cut = False Then
    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False
End If

End Sub

Public Sub GoDelete()

SaveUnDo

On Error GoTo EndDelete

Dim j As Long

    For j = 1 To UBound(SaveSelect)
           GData(SaveSelect(j)) = False
    Next
EndDelete:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False
    
End Sub

Public Sub GoAutoDo()

Dim First_D As Long, Last_D As Long, First_B As Long, Last_B As Long, i As Long, u As Long
SaveUnDo

On Error GoTo EndDelete

    First_D = SaveSelect(1) Mod 8
    Last_D = SaveSelect(UBound(SaveSelect)) Mod 8
    First_B = (SaveSelect(1) - First_D) / 8
    Last_B = (SaveSelect(UBound(SaveSelect)) - Last_D) / 8
    
    For i = First_B To Last_B
        For u = First_D To Last_D
           GData(i * 8 + u) = True
        Next u
    Next i
    
EndDelete:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False
    
End Sub

Public Sub PushOutKey()

Dim curPos As Long, cX As Single, ToOffset As Single, ToNowBeat As Long

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

SaveUnDo

On Error GoTo EndPaste
If CurrPos >= 0 Then

    Dim PasteData As String, cPos As Long, APos As Long, DataC() As Integer
    
    PasteData = Clipboard.GetText
    ReDim DataC(0)
  If Mid(PasteData, 1, 4) = "BBQ," Then
        cPos = 5
        Do
          APos = InStr(cPos, PasteData, ",") + 1
          ReDim Preserve DataC(UBound(DataC) + 1)
          If APos > 1 Then
            DataC(UBound(DataC)) = CInt(Val(Mid(PasteData, cPos, APos - cPos)))
          Else
            DataC(UBound(DataC)) = CInt(Val(Mid(PasteData, cPos)))
          End If
          cPos = APos
        Loop While APos > 1
        
        If UBound(DataC) > 0 Then
            For j = 1 To UBound(DataC)
                If GData(ToNowBeat * 8 + DataC(j)) = False Then GData(ToNowBeat * 8 + DataC(j)) = True
            Next
        End If
    End If
End If
EndPaste:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub GoAllRandom()

Dim St As Integer, Kt As Integer, i As Long, j As Long

SaveUnDo

On Error GoTo EndRanKey

            For i = 0 To UBound(GData) - 1 Step 8
                St = 0
                
                For j = 0 To 5
                   St = St + IIf(GData(i + j) = True, 1, 0)
                Next
                
                If St > 0 Then
                    For j = 0 To 5
                        GData(i + j) = False
                    Next
                    
                    Do While St > 0
                        Randomize
                        Kt = Fix(6 * Rnd)
                        Kt = IIf((Kt = 6), 0, Kt)
                        
                        If GData(i + Kt) = False Then
                            GData(i + Kt) = True
                            St = St - 1
                        End If
                    Loop
                End If
            Next
EndRanKey:
End Sub

Public Sub GoLeftRightRandom()

Dim St As Integer, Kt As Integer, i As Long, j As Long, LOR As Long, lR As Long

SaveUnDo

On Error GoTo EndRanKey

Randomize
LOR = Fix(35 * Rnd)

            For i = 0 To UBound(GData) - 1 Step 8
                St = 0
                
                For j = 0 To 5
                   St = St + IIf(GData(i + j) = True, 1, 0)
                Next
                
                If St > 0 Then
                    
                    For j = 0 To 5
                        GData(i + j) = False
                    Next
                    
                    Do While St > 0
                        Randomize
                        Kt = Fix(3 * Rnd)
                        Kt = IIf((Kt = 3), 0, Kt)
                        
                        If LOR Mod 2 = 0 Then
                            lR = 3
                        Else
                            lR = 0
                        End If
                        
                        If GData(i + Kt + lR) = False Then
                            GData(i + Kt + lR) = True
                            St = St - 1
                        End If
                    Loop
                    
                    LOR = LOR + 1
                End If
            Next
EndRanKey:
End Sub

Public Sub AllKeyLeft()

Dim i As Long, ky As Long

If Admin = False Then On Error Resume Next

SaveUnDo

For i = 0 To TotalBeat - 2
        For ky = 0 To 7
            If (i + 1) * 8 + 7 > UBound(GData) Then ReDim Preserve GData((i + 1) * 8 + 7)
            GData(i * 8 + ky) = GData((i + 1) * 8 + ky)
        Next
Next

End Sub

Public Sub AllKeyRight()

Dim i As Long, ky As Long

If Admin = False Then On Error Resume Next

SaveUnDo

For i = TotalBeat - 1 To 1 Step -1
        For ky = 0 To 7
            If (i) * 8 + 7 > UBound(GData) Then ReDim Preserve GData((i + 1) * 8 + 7)
           GData(i * 8 + ky) = GData((i - 1) * 8 + ky)
        Next
Next

End Sub

Public Sub OneKey()

Dim St As Integer, i As Long, ky As Long

If Admin = False Then On Error Resume Next

SaveUnDo

On Error GoTo EndOneKey
            For i = 0 To UBound(GData) - 1 Step 8
            
                St = 0
                For j = 0 To 5
                   St = St + IIf(GData(i + j) = True, 1, 0)
                Next
                
                If St > 0 Then
                    For j = 0 To 5
                        GData(i + j) = False
                    Next
                    
                    Do While St > 0
                        If GData(i + 5) = False Then
                            GData(i + 5) = True
                            St = St - 1
                        End If
                    Loop
                End If
            Next
EndOneKey:
End Sub

Public Sub SaveSetting(Optional NewFile As Boolean, Optional NeedMsg As Boolean)

Dim SaveString As String

If Admin = False Then On Error Resume Next

SaveString = "OggFolder=" + cmt.OggFolder_Text.Text + vbCrLf + "ScriptFolder=" + cmt.ScriptFolder_Text.Text + vbCrLf + "AllSongList=" + cmt.AllSongList_Text.Text + vbCrLf + "BeatUpList=" + cmt.BeatUpList_Text.Text + vbCrLf + "Language=" + CStr(Language) + vbCrLf

        CreateDir TempPath + "\cma\"
        
        Open TempPath + "\cma\" + "Setting.ini" For Binary As #1
            Put #1, , SaveString
        Close #1
        
        Enrypt_12 TempPath + "cma\", "setting", App.Path + "\"
        
        DeleteDir TempPath + "cma\"
        
        If NewFile = False Then
            cmt.Frame3.Visible = False
            cmt.Frame2.Visible = True
            cmt.Frame1.Visible = False
            cmt.Frame4.Visible = False
            cmt.Frame5.Visible = False

            Singer = cmt.Single_Text.Text
            Melody = cmt.Melody_Text.Text
            Author = cmt.Author_Text.Text
            level = cmt.Level_Text.Text
            MusicCode = cmt.MusicCode_Text.Text
            OggF = cmt.OggFolder_Text.Text
            ScrF = cmt.ScriptFolder_Text.Text
            ASL = cmt.AllSongList_Text.Text
            BUL = cmt.BeatUpList_Text.Text
        
            If NeedMsg = False Then MsgBox "儲存完成", 0, "系統訊息"
        
        End If
        
End Sub

Public Sub SetL()

If Admin = False Then On Error Resume Next

        If Language = 0 Then
            cmt.SetChinese.Enabled = False
            cmt.SetChinese.Checked = True
            cmt.SetEnglish.Enabled = True
            cmt.SetEnglish.Checked = False
            
            cmt.AMenu.Caption = "檔案"
            cmt.NewFile.Caption = "開新檔案"
            cmt.openmusic.Caption = "開啟音樂"
            cmt.OpenSlk.Caption = "導入slk文件"
            cmt.SaveSlkButton.Caption = "導出slk文件"
            cmt.OpenKbe.Caption = "導入kbe文件"
            cmt.SaveKbe.Caption = "導出kbe文件"
            cmt.OpenDdr.Caption = "導入ddr文件"
            cmt.SaveDdr.Caption = "導出ddr文件"
            cmt.OpenCbe.Caption = "導入cbe文件"
            cmt.SaveCbe.Caption = "導出cbe文件"
            cmt.HighSave.Caption = "高級儲存"
            cmt.LoadAutoSave.Caption = "開啟臨時儲存"
            cmt.BMenu.Caption = "設定"
            cmt.setting.Caption = "普通設定"
            cmt.ProSetting.Caption = "高級設定"
            cmt.CMenu.Caption = "播放器"
            cmt.ShowOrHide.Caption = "顯示/隱藏"
            cmt.PlayOrStop.Caption = "播放"
            cmt.EndSong.Caption = "停止"
            cmt.PlaySpace.Caption = "播放空白鍵部分"
            cmt.dmenu.Caption = "模式"
            cmt.SetNormalMode.Caption = "正常模式"
            cmt.SetSeeMode.Caption = "觀戰模式"
            cmt.SetGameMode.Caption = "遊戲模式"
            cmt.emenu.Caption = "功能"
            cmt.AutoSpace16.Caption = "自動16空白鍵"
            cmt.AutoSpace24.Caption = "自動24空白鍵"
            cmt.AutoSpace32.Caption = "自動32空白鍵"
            cmt.AutoSpace48.Caption = "自動48空白鍵"
            cmt.DelSpace.Caption = "後面空白鍵全刪"
            cmt.SpaceOut16.Caption = "後面空白鍵退後16"
            cmt.AllRandomKey.Caption = "全部箭頭隨機"
            cmt.LRRandomKey.Caption = "全部箭頭左右隨機"
            cmt.KeyLeft.Caption = "全部箭頭左移一格"
            cmt.KeyRight.Caption = "全部箭頭右移一格"
            cmt.AllKeyOne.Caption = "全部箭頭變為1"
            cmt.DelAllNote.Caption = "刪除全部箭頭"
            cmt.RxMove.Caption = "直線移動"
            cmt.fmenu.Caption = "編輯"
            cmt.CutButton.Caption = "剪下"
            cmt.CopyButton.Caption = "復制"
            cmt.PushUpButton.Caption = "貼上"
            cmt.DelButton.Caption = "刪除"
            cmt.UnDo.Caption = "復原"
            cmt.SomeKeyRandom.Caption = "部分箭頭隨機"
            cmt.SomeKeyLRRandom.Caption = "部分箭頭左右隨機"
            cmt.SomeKeyLeft.Caption = "部分箭頭左移一格"
            cmt.SomeKeyRight.Caption = "部分箭頭右移一格"
            cmt.SomeKeyBeOne.Caption = "部分箭頭變為1"
            cmt.gmenu.Caption = "段落"
            cmt.gshow.Caption = "顯示/隱藏"
            cmt.MapMenu.Caption = "地圖"
            cmt.ChangeBack1.Caption = "籃球場"
            cmt.ChangeBack3.Caption = "漢江公園"
            cmt.ChangeBack5.Caption = "光華門"
            cmt.ChangeBack6.Caption = "紫禁場"
            cmt.ChangeBack7.Caption = "萬聖節"
            cmt.ChangeBack8.Caption = "結他場"
            cmt.ChangeBack9.Caption = "滑雪場"
            cmt.HighAdmin.Caption = "管理"
            cmt.SetLanguage.Caption = "Language"
        
            cmt.AddUser.Caption = " 儲存 "
            cmt.CancelAddUser.Caption = " 取消 "
            cmt.AddTeam.Caption = " 加入 "
            cmt.UseTeam.Caption = " 貼上 "
            cmt.HideTeam.Caption = " 隱藏 "
            cmt.Singer_Set.Caption = "歌手名:"
            cmt.Melody_Set.Caption = "歌名:"
            cmt.Autho_Set.Caption = "編步者:"
            cmt.Level_Set.Caption = "等級:"
            cmt.MusicCode_Set.Caption = "音樂編號:"
            cmt.OggFolder_Set.Caption = "音樂資料夾:"
            cmt.ScriptFolder_Set.Caption = "步舞資料夾:"
            cmt.AllSongList_Set.Caption = "同步列表:"
            cmt.BeatUpList_Set.Caption = "BeatUp列表:"
            cmt.User_Set.Caption = "使用者:"
            cmt.SaveHighSetting.Caption = " 儲存 "
            cmt.CancelHighSetting.Caption = " 取消 "
            cmt.Label_Load.Caption = " 開啟臨時儲存 "
            cmt.Label_save.Caption = " 儲存 "
            
            cmt.Frame1.Caption = "普通設定"
            cmt.Frame2.Caption = "播放器"
            cmt.Frame3.Caption = "高級設定"
            cmt.Frame4.Caption = "段落工具"
            cmt.Frame5.Caption = "內部管理"
            
            cmt.SetByUser.Caption = "使用者定義"
            
            cmt.NetWork.Caption = "網路連線"
            cmt.SaveAsCbg.Caption = "導出遊戲檔"
            cmt.CheckError.Caption = "檢查Slk中的錯誤"
            cmt.AutoFillAll.Caption = "自動填滿"
            
        ElseIf Language = 1 Then
            cmt.SetChinese.Enabled = True
            cmt.SetChinese.Checked = False
            cmt.SetEnglish.Enabled = False
            cmt.SetEnglish.Checked = True
            
            cmt.AMenu.Caption = "File"
            cmt.NewFile.Caption = "New File"
            cmt.openmusic.Caption = "Open Music"
            cmt.OpenSlk.Caption = "Open Slk File"
            cmt.SaveSlkButton.Caption = "Save As Slk"
            cmt.OpenKbe.Caption = "Open Kbe File"
            cmt.SaveKbe.Caption = "Save As Kbe"
            cmt.OpenDdr.Caption = "Open Ddr File"
            cmt.SaveDdr.Caption = "Save As Ddr"
            cmt.OpenCbe.Caption = "Open Cbe File"
            cmt.SaveCbe.Caption = "Save As Cbe"
            cmt.HighSave.Caption = "Advance Save"
            cmt.LoadAutoSave.Caption = "Load Temporary Save"
            cmt.BMenu.Caption = "Setting"
            cmt.setting.Caption = "Default Setting"
            cmt.ProSetting.Caption = "Advance Setting"
            cmt.CMenu.Caption = "Music Player"
            cmt.ShowOrHide.Caption = "Toggle Menu On / Off"
            cmt.PlayOrStop.Caption = "Play"
            cmt.EndSong.Caption = "Stop"
            cmt.PlaySpace.Caption = "Only Play Music Form This Space"
            cmt.dmenu.Caption = "Mode"
            cmt.SetNormalMode.Caption = "Classic Mode"
            cmt.SetSeeMode.Caption = "AutoPlay Mode"
            cmt.SetGameMode.Caption = "Game Mode"
            cmt.emenu.Caption = "Function"
            cmt.AutoSpace16.Caption = "Auto Spaces (16 Beats)"
            cmt.AutoSpace24.Caption = "Auto Spaces (24 Beats)"
            cmt.AutoSpace32.Caption = "Auto Spaces (32 Beats)"
            cmt.AutoSpace48.Caption = "Auto Spaces (48 Beats)"
            cmt.DelSpace.Caption = "Del After Spaces Form This Space"
            cmt.SpaceOut16.Caption = "Move Spaces (16 Beats) Form This Space"
            cmt.AllRandomKey.Caption = "All Notes To Random"
            cmt.LRRandomKey.Caption = "All Notes To Left Right Random"
            cmt.KeyLeft.Caption = "Shift Notes Left"
            cmt.KeyRight.Caption = "Shift Notes Right"
            cmt.AllKeyOne.Caption = "All Note To Be 1"
            cmt.DelAllNote.Caption = "Del All Notes"
            cmt.RxMove.Caption = "Toggle Grid"
            cmt.fmenu.Caption = "Edit"
            cmt.CutButton.Caption = "Cut"
            cmt.CopyButton.Caption = "Copy"
            cmt.PushUpButton.Caption = "Paste"
            cmt.DelButton.Caption = "Delete"
            cmt.UnDo.Caption = "UnDo"
            cmt.SomeKeyRandom.Caption = "Selects Notes Random"
            cmt.SomeKeyLRRandom.Caption = "Selects Notes Left Right Random"
            cmt.SomeKeyLeft.Caption = "Shift Selects Notes Left"
            cmt.SomeKeyRight.Caption = "Shift Selects Notes Right"
            cmt.SomeKeyBeOne.Caption = "Selects Note To Be 1"
            cmt.gmenu.Caption = "Paragraph"
            cmt.gshow.Caption = "Toggle Menu On / Off"
            cmt.MapMenu.Caption = "Map"
            cmt.ChangeBack1.Caption = "Hip Hop Avenue"
            cmt.ChangeBack3.Caption = "Han Jang Park"
            cmt.ChangeBack5.Caption = "Kwang Hwa Mun"
            cmt.ChangeBack6.Caption = "Forbidden City"
            cmt.ChangeBack7.Caption = "Halloween Stage"
            cmt.ChangeBack8.Caption = "Guitar Stage"
            cmt.ChangeBack9.Caption = "Snow Valley"
            cmt.HighAdmin.Caption = "Admin"
            cmt.SetLanguage.Caption = "語言"
            
            cmt.AddUser.Caption = "  Save  "
            cmt.CancelAddUser.Caption = "  Cancel  "
            cmt.AddTeam.Caption = "  Add  "
            cmt.UseTeam.Caption = "  Paste  "
            cmt.HideTeam.Caption = "  Hide  "
            cmt.Singer_Set.Caption = "Singer:"
            cmt.Melody_Set.Caption = "Melody:"
            cmt.Autho_Set.Caption = "Author:"
            cmt.Level_Set.Caption = "Level:"
            cmt.MusicCode_Set.Caption = "Music Code:"
            cmt.OggFolder_Set.Caption = "Ogg Folder:"
            cmt.ScriptFolder_Set.Caption = "Script Folder:"
            cmt.AllSongList_Set.Caption = "All Song List:"
            cmt.BeatUpList_Set.Caption = "BeatUp List:"
            cmt.User_Set.Caption = "User:"
            cmt.SaveHighSetting.Caption = "  Save  "
            cmt.CancelHighSetting.Caption = "  Cancel  "
            cmt.Label_Load.Caption = " Load Temp  "
            cmt.Label_save.Caption = "  Save  "

            cmt.Frame1.Caption = "Default Setting"
            cmt.Frame2.Caption = "Music Player"
            cmt.Frame3.Caption = "Advance Setting"
            cmt.Frame4.Caption = "Paragraph Tool"
            cmt.Frame5.Caption = "Admin"
            
            cmt.SetByUser.Caption = "User Setting"
            
            cmt.NetWork.Caption = "NetWork"
            cmt.SaveAsCbg.Caption = "Save As Game File"
            cmt.CheckError.Caption = "Check Error Form Slk"
            cmt.AutoFillAll.Caption = "Auto Fill It All"
            
        End If

End Sub

Public Sub LoadINI()

Dim TData() As String, Number As Long, i As Long

If Admin = False Then On Error Resume Next
        
        If Fso.FileExists(App.Path + "\" + "setting.cdiu") = False Then SaveSetting True
        
        Decrypt_12 App.Path + "\" + "setting.cdiu", 0
        
        If Fso.FileExists(App.Path + "\cma\" + SaveFileINI) = False Then SaveSetting True: Decrypt_12 App.Path + "\" + "setting.cdiu", 0
        
        Open App.Path + "\cma\" + SaveFileINI For Input As #1
            Do
            ReDim Preserve TData(Number)
            Line Input #1, TData(Number)
            Number = Number + 1
            Loop Until EOF(1)
        Close #1
    
        DeleteDir App.Path + "\cma\"
       
        For i = LBound(TData) To UBound(TData)
            If InStr(TData(i), "OggFolder=") > 0 Then OggF = Mid(TData(i), 11, Len(TData(i))): cmt.OggFolder_Text.Text = OggF
            If InStr(TData(i), "ScriptFolder=") > 0 Then ScrF = Mid(TData(i), 14, Len(TData(i))): cmt.ScriptFolder_Text.Text = ScrF
            If InStr(TData(i), "AllSongList=") > 0 Then ASL = Mid(TData(i), 13, Len(TData(i))): cmt.AllSongList_Text.Text = ASL
            If InStr(TData(i), "BeatUpList=") > 0 Then BUL = Mid(TData(i), 12, Len(TData(i))): cmt.BeatUpList_Text.Text = BUL
            If InStr(TData(i), "Language=") > 0 Then Language = Mid(TData(i), 10, Len(TData(i)))
        Next i

Singer = cmt.Single_Text.Text
Melody = cmt.Melody_Text.Text
Author = cmt.Author_Text.Text
level = cmt.Level_Text.Text
MusicCode = cmt.MusicCode_Text.Text

End Sub

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

Public Function DeleteDir(DirName As String) As Boolean

Dim DirList As String, TempDirName As String
    
        If Not (Right(DirName, 1) = "\") Then DirName = DirName & "\"

On Error Resume Next

        Do While Len(Dir$(DirName, vbDirectory)) <> 0
            TempDirName = DirName
            DirList = Dir$(TempDirName, vbDirectory)
            Kill TempDirName & "*.*"
            
                    Do While Len(DirList) <> 0
                        DoEvents
                            If DirList <> "." And DirList <> ".." Then
                                TempDirName = TempDirName & DirList & "\"
                                DirList = Dir$(TempDirName, vbDirectory)
                                Kill TempDirName & "*.*"
                            End If
                        DirList = Dir
                    Loop
            RmDir TempDirName
        Loop

DeleteDir = IIf(Len(Dir$(DirName, vbDirectory)) = 0, True, False)

End Function

Public Sub DoSomeKeyRight()

If Admin = False Then On Error Resume Next

SaveUnDo

On Error GoTo EndDDoSomeKeyRight

Dim j As Long

    For j = UBound(SaveSelect) To 1 Step -1
            GData(SaveSelect(j) + 8) = GData(SaveSelect(j))
            GData(SaveSelect(j)) = False
    Next
    
EndDDoSomeKeyRight:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub DoSomeKeyLeft()

SaveUnDo

On Error GoTo EndDDoSomeKeyLeft

Dim j As Long

    For j = 1 To UBound(SaveSelect)
            GData(SaveSelect(j) - 8) = GData(SaveSelect(j))
            GData(SaveSelect(j)) = False
    Next j
    
EndDDoSomeKeyLeft:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub DoSomeOneKey()

SaveUnDo

Dim j As Long, NowKey As Long

On Error GoTo EndDoSomeOneKey

    For j = 1 To UBound(SaveSelect)
        NowKey = (SaveSelect(j) - (SaveSelect(j) Mod 8)) / 8
            GData(NowKey * 8 + 5) = True
            GData(SaveSelect(j)) = False
    Next j
EndDoSomeOneKey:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub DoSomeRandom()

SaveUnDo

Dim j As Long, NowKey As Long, St As Integer, Kt As Integer, i As Long

On Error GoTo EndDoSomeOneKey

    For j = 1 To UBound(SaveSelect)
        NowKey = (SaveSelect(j) - (SaveSelect(j) Mod 8)) / 8
        St = 0
                For i = 0 To 5
                   St = St + IIf(GData(NowKey * 8 + i) = True, 1, 0)
                Next i
                
                For i = 0 To 5
                   GData(NowKey * 8 + i) = False
                Next i
                
                    Do While St > 0
                        Randomize
                        Kt = Fix(6 * Rnd)
                        Kt = IIf((Kt = 6), 0, Kt)
                  
                        If GData(NowKey * 8 + Kt) = False Then
                            GData(NowKey * 8 + Kt) = True
                            St = St - 1
                        End If
                    Loop
    Next j
EndDoSomeOneKey:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub DoSomeLRRandom()

SaveUnDo

Dim j As Long, NowKey As Long, St As Integer, Kt As Integer, i As Long, LOR As Long, lR As Long

On Error GoTo EndDoSomeOneKey

        Randomize
        LOR = Fix(35 * Rnd)

    For j = 1 To UBound(SaveSelect)
        NowKey = (SaveSelect(j) - (SaveSelect(j) Mod 8)) / 8
        St = 0
                For i = 0 To 5
                   St = St + IIf(GData(NowKey * 8 + i) = True, 1, 0)
                Next i
        
        If St > 0 Then
                For i = 0 To 5
                   GData(NowKey * 8 + i) = False
                Next i
                
                    Do While St > 0
                        Randomize
                        Kt = Fix(3 * Rnd)
                        Kt = IIf((Kt = 3), 0, Kt)
                        
                        If LOR Mod 2 = 0 Then
                            lR = 3
                        Else
                            lR = 0
                        End If
                        
                  
                        If GData(NowKey * 8 + Kt + lR) = False Then
                            GData(NowKey * 8 + Kt + lR) = True
                            St = St - 1
                        End If
                    Loop
            LOR = LOR + 1
        End If
    Next j
EndDoSomeOneKey:

    ReDim SaveSelect(0)
    cmt.MouseMove = False
    cmt.MouseDown = False

End Sub

Public Sub SaveThisSpace(SaveBeat As Long)

Dim NewData(247) As Boolean, i As Long, St As Boolean, j As Long, Name As String, CopyData As String

    For i = 0 To 247
        NewData(i) = GData((SaveBeat + 1) * 8 + i)
    Next i

Name = "S"
CopyData = "BBQ,6"

On Error GoTo EndSaveThisSpace
            For i = 0 To UBound(NewData) Step 8
            
                St = False
                For j = 0 To 5
                    If NewData(i + j) = True Then
                        St = True
                        
                        If St = True And j = 0 Then Name = Name + "9"
                        If St = True And j = 1 Then Name = Name + "6"
                        If St = True And j = 2 Then Name = Name + "3"
                        If St = True And j = 3 Then Name = Name + "7"
                        If St = True And j = 4 Then Name = Name + "4"
                        If St = True And j = 5 Then Name = Name + "1"
                        
                        For k = 0 To 5
                            If k <> j Then NewData(i + k) = False
                        Next k
                        
                        Exit For
                    End If
                Next
                
                If St <> True Then
                    Name = Name + "_"
                End If
            Next
            
            Name = Name + "S"
EndSaveThisSpace:

        For i = 1 To 31
            Check = Mid(Name, i + 1, 1)
            If Check = "9" Then CopyData = CopyData + "," + CStr(i * 8 + 0)
            If Check = "6" Then CopyData = CopyData + "," + CStr(i * 8 + 1)
            If Check = "3" Then CopyData = CopyData + "," + CStr(i * 8 + 2)
            If Check = "7" Then CopyData = CopyData + "," + CStr(i * 8 + 3)
            If Check = "4" Then CopyData = CopyData + "," + CStr(i * 8 + 4)
            If Check = "1" Then CopyData = CopyData + "," + CStr(i * 8 + 5)
        Next i

cmt.Team_List.AddItem Name, UBound(FastTeam) + 1
ReDim Preserve FastTeam(UBound(FastTeam) + 1)
FastTeam(UBound(FastTeam)) = CopyData
'MsgBox Name
'MsgBox CopyData

SaveFast

End Sub

Public Sub SaveFast()

Dim Name As String * 50, NameB As String * 150

If Admin = False Then On Error Resume Next

    cma5.CreateDir TempPath + "\cma\"

    Open TempPath + "\cma\fast.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileFast"
    Put #1, , CLng(UBound(FastTeam))
    
    For i = 0 To UBound(FastTeam)
        Name = cmt.Team_List.List(i)
        NameB = FastTeam(i)
        Put #1, , Name
        Put #1, , NameB
    Next i
    
    Close #1
    
        Enrypt_12 TempPath + "cma\", "More", App.Path + "\"
        
        DeleteDir TempPath + "cma\"
    
End Sub

Public Sub LoadFast()

Dim SigT As String * 22, Tmp As Long, i As Long, Name As String * 50, NameB As String * 150

If Admin = False Then On Error Resume Next

        If Fso.FileExists(App.Path + "\" + "More.cdiu") = False Then OnceFast
        
        Decrypt_12 App.Path + "\" + "More.cdiu", 0
        
        If Fso.FileExists(App.Path + "\cma\fast.cbe") = False Then OnceFast: Decrypt_12 App.Path + "\" + "More.cdiu", 0
        
        Open App.Path + "\cma\" + "fast.cbe" For Binary As #1
            Get #1, 1, SigT
            If SigT = "CdiuBeatEditorFileFast" Then
        Get #1, , Tmp
        ReDim FastTeam(Tmp)
        
        For i = 0 To Tmp
            Get #1, , Name
            Get #1, , NameB
            
            cmt.Team_List.AddItem Trim(Name), i
            FastTeam(i) = Trim(NameB)
        Next i
        End If
        Close #1

cmt.Team_List.ListIndex = 0
End Sub

Public Sub OnceFast()

If Admin = False Then On Error Resume Next

cmt.Team_List.AddItem "S_1_1_1_1_1_1_1_1_1_1_1_1_1_1_1_S", 0
cmt.Team_List.ListIndex = 0
ReDim FastTeam(0)
FastTeam(0) = "BBQ,6,21,37,53,69,85,101,117,133,149,165,181,197,213,229,245"
SaveFast
End Sub

Public Sub SaveUnDo()

If Admin = False Then On Error Resume Next

ReDim UData(UBound(GData))
ReDim OUData(UBound(SData))

cq.IcyCopyMemory ByVal VarPtr(UData(0)), ByVal VarPtr(GData(0)), (UBound(GData) + 1) * 2
cq.IcyCopyMemory ByVal VarPtr(OUData(0)), ByVal VarPtr(SData(0)), (UBound(SData) + 1)

SaUnDo = True

End Sub

Public Sub UnDoIt()

If Admin = False Then On Error Resume Next

    If SaUnDo = True Then
        cq.IcyCopyMemory ByVal VarPtr(GData(0)), ByVal VarPtr(UData(0)), (UBound(UData) + 1) * 2
        cq.IcyCopyMemory ByVal VarPtr(SData(0)), ByVal VarPtr(OUData(0)), (UBound(OUData) + 1)
    End If

End Sub

Public Sub LoadAcv()

If Admin = False Then On Error Resume Next

        If Fso.FileExists(App.Path + "\" + "BG1.cdiu") = False Then MsgBox IIf(Language = 0, "找不到檔案 BG1.cdiu", "File Not Found BG1.cdiu"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
        If Fso.FileExists(App.Path + "\" + "BG2.cdiu") = False Then MsgBox IIf(Language = 0, "找不到檔案 BG2.cdiu", "File Not Found BG2.cdiu"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
        If Fso.FileExists(App.Path + "\" + "BG3.cdiu") = False Then MsgBox IIf(Language = 0, "找不到檔案 BG2.cdiu", "File Not Found BG3.cdiu"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
        If Fso.FileExists(App.Path + "\" + "Main.cdiu") = False Then MsgBox IIf(Language = 0, "找不到檔案 Main.cdiu", "File Not Found Main.cdiu"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
        
        Decrypt_12 App.Path + "\" + "BG1.cdiu", 0, App.Path + "\user\", True
        Decrypt_12 App.Path + "\" + "BG2.cdiu", 0, App.Path + "\user\", True
        Decrypt_12 App.Path + "\" + "BG3.cdiu", 0, App.Path + "\user\", True
        Decrypt_12 App.Path + "\" + "Main.cdiu", 0, App.Path + "\user\", True
        
       CheckAcvFile
End Sub

Public Sub CheckAcvFile()

Dim Check As Boolean, i As Long, Which As String, CheckFileA() As String, o As Long

AddBackArray CheckFileA, "BG1\1.dds"
AddBackArray CheckFileA, "BG2\2.dds"
AddBackArray CheckFileA, "BG3\3.dds"
AddBackArray CheckFileA, "BG4\4.dds"
AddBackArray CheckFileA, "BG5\5.dds"
AddBackArray CheckFileA, "BG6\6.dds"
AddBackArray CheckFileA, "BG7\7.dds"
AddBackArray CheckFileA, "BG8\8.dds"
AddBackArray CheckFileA, "BG9\9.dds"
AddBackArray CheckFileA, "BG11\11.dds"

AddBackArray CheckFileA, "BMP\ALINE.bmp"
AddBackArray CheckFileA, "BMP\L3.dds"
AddBackArray CheckFileA, "BMP\NORMALBACK.bmp"
AddBackArray CheckFileA, "BMP\NORMALRX.bmp"
AddBackArray CheckFileA, "BMP\P.jpg"
AddBackArray CheckFileA, "BMP\R3.dds"

AddBackArray CheckFileA, "SLK\LIST1.txt"
AddBackArray CheckFileA, "SLK\LIST2.txt"
AddBackArray CheckFileA, "SLK\SLK2HEADER.txt"
AddBackArray CheckFileA, "SLK\SLKHEADER.txt"

AddBackArray CheckFileA, "SOUND\BEAT.ogg"
AddBackArray CheckFileA, "SOUND\GREAT.ogg"
AddBackArray CheckFileA, "SOUND\MISS.ogg"
AddBackArray CheckFileA, "SOUND\READY.ogg"
AddBackArray CheckFileA, "SOUND\SPACE.ogg"
AddBackArray CheckFileA, "SOUND\START.ogg"

AddBackArray CheckFileA, "DDS\BAD.dds"
AddBackArray CheckFileA, "DDS\BCUP.dds"
AddBackArray CheckFileA, "DDS\BUP.dds"
AddBackArray CheckFileA, "DDS\BYBUPB.dds"
AddBackArray CheckFileA, "DDS\BYBUPF.dds"
AddBackArray CheckFileA, "DDS\BYBUPY.dds"
AddBackArray CheckFileA, "DDS\CA.dds"
AddBackArray CheckFileA, "DDS\CB.dds"
AddBackArray CheckFileA, "DDS\CBBACK.dds"
AddBackArray CheckFileA, "DDS\CE.dds"
AddBackArray CheckFileA, "DDS\COOL.dds"
AddBackArray CheckFileA, "DDS\CP.dds"
AddBackArray CheckFileA, "DDS\CT.dds"
AddBackArray CheckFileA, "DDS\CU.dds"
AddBackArray CheckFileA, "DDS\CYBACK.dds"
AddBackArray CheckFileA, "DDS\GREAT.dds"
AddBackArray CheckFileA, "DDS\K0.dds"
AddBackArray CheckFileA, "DDS\K1.dds"
AddBackArray CheckFileA, "DDS\K2.dds"
AddBackArray CheckFileA, "DDS\K3.dds"
AddBackArray CheckFileA, "DDS\K4.dds"
AddBackArray CheckFileA, "DDS\K5.dds"
AddBackArray CheckFileA, "DDS\K6.dds"
AddBackArray CheckFileA, "DDS\K7.dds"
AddBackArray CheckFileA, "DDS\L0.dds"
AddBackArray CheckFileA, "DDS\L1.dds"
AddBackArray CheckFileA, "DDS\L2.dds"
AddBackArray CheckFileA, "DDS\L3.dds"
AddBackArray CheckFileA, "DDS\L4.dds"
AddBackArray CheckFileA, "DDS\L5.dds"
AddBackArray CheckFileA, "DDS\LOGO0.dds"
AddBackArray CheckFileA, "DDS\LOGO1.dds"
AddBackArray CheckFileA, "DDS\MISS.dds"
AddBackArray CheckFileA, "DDS\MP.dds"
AddBackArray CheckFileA, "DDS\PERFECT.dds"
AddBackArray CheckFileA, "DDS\POWER.dds"
AddBackArray CheckFileA, "DDS\READY0.dds"
AddBackArray CheckFileA, "DDS\READY1.dds"
AddBackArray CheckFileA, "DDS\READY2.dds"
AddBackArray CheckFileA, "DDS\READY3.dds"
AddBackArray CheckFileA, "DDS\READY4.dds"
AddBackArray CheckFileA, "DDS\RP.dds"
AddBackArray CheckFileA, "DDS\SLINE.dds"
AddBackArray CheckFileA, "DDS\SPACE.dds"
AddBackArray CheckFileA, "DDS\SPACEBAR.dds"
AddBackArray CheckFileA, "DDS\SPOWER.dds"
AddBackArray CheckFileA, "DDS\YCUP.dds"
AddBackArray CheckFileA, "DDS\YUP.dds"
AddBackArray CheckFileA, "DDS\CRAZY.dds"
AddBackArray CheckFileA, "DDS\RM2.dds"
AddBackArray CheckFileA, "DDS\RM1.dds"
AddBackArray CheckFileA, "DDS\LM2.dds"
AddBackArray CheckFileA, "DDS\LM1.dds"
AddBackArray CheckFileA, "DDS\RMISS.dds"
AddBackArray CheckFileA, "DDS\LMISS.dds"


    For i = 10 To 39
        AddBackArray CheckFileA, "COMBO\C" + CStr(i) + ".dds"
    Next i

If GetUBound(CheckFile) < 0 Then MsgBox IIf(Language = 0, "找不到檔案 " + CheckFileA(0), "File Not Found " + CheckFileA(0)), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
    For i = 0 To UBound(CheckFileA)
        Check = False
        For o = 0 To UBound(CheckFile)
            If CheckFileA(i) = CheckFile(o) Then Check = True
        Next o
        
        If Check = False Then MsgBox IIf(Language = 0, "找不到檔案 " + CheckFileA(i), "File Not Found " + CheckFileA(i)), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
        
    Next i

End Sub
