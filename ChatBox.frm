VERSION 5.00
Begin VB.Form ChatBox 
   BorderStyle     =   1  '單線固定
   Caption         =   "房間資訊"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14625
   Icon            =   "ChatBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   14625
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton ChooseSong 
      Caption         =   "選擇歌曲"
      Height          =   735
      Left            =   9240
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton BootUser 
      Caption         =   "踢走玩家"
      Height          =   735
      Left            =   12000
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton ExitRoom 
      Caption         =   "離開房間"
      Height          =   735
      Left            =   5760
      TabIndex        =   12
      Top             =   6240
      Width           =   2055
   End
   Begin VB.ListBox People_List 
      Height          =   5280
      ItemData        =   "ChatBox.frx":17D2A
      Left            =   11640
      List            =   "ChatBox.frx":17D2C
      TabIndex        =   10
      Top             =   720
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "開始/準備"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Timer UpdateText 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4560
      Top             =   1080
   End
   Begin VB.CommandButton SendMsg 
      Caption         =   "送出訊息"
      Height          =   375
      Left            =   10200
      TabIndex        =   6
      Top             =   5720
      Width           =   1215
   End
   Begin VB.TextBox InPutBox 
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   9855
   End
   Begin VB.TextBox ShowTextBox 
      Height          =   2895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   4
      Top             =   2640
      Width           =   11175
   End
   Begin VB.Label ShowWhoMake 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   18
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label ShowRoomSong 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   17
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label WhoMake_Label 
      AutoSize        =   -1  'True
      Caption         =   "歌曲作者"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Song_Label 
      AutoSize        =   -1  'True
      Caption         =   "遊戲歌曲"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label People_Label 
      AutoSize        =   -1  'True
      Caption         =   "在線玩家"
      Height          =   180
      Left            =   11760
      TabIndex        =   11
      Top             =   360
      Width           =   720
   End
   Begin VB.Label ShowRoomHost 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label HostName_Label 
      AutoSize        =   -1  'True
      Caption         =   "房主名稱"
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label ShowRoomName 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   45
   End
   Begin VB.Label ShowRoomName_Label 
      AutoSize        =   -1  'True
      Caption         =   "房間名稱"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   720
   End
   Begin VB.Label ShowRoomNumber 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label ShowRoomNumber_Label 
      AutoSize        =   -1  'True
      Caption         =   "房間編號"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "ChatBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LastMsg As String

Dim Fso As New FileSystemObject

Public Function FindFile(GData() As String, ByRef Path As String, Optional TPath As String)

Dim Folder As Object, file As Object, TFolder As Object, Number As Boolean

Set TFolder = Fso.GetFolder(Path)

        For Each Folder In TFolder.SubFolders
            FindFile GData, Path + "\" + Folder.Name, TPath + Folder.Name + "\"
            Number = True
        Next Folder

        For Each file In TFolder.Files
            AddBackArray GData, TPath + file.Name
            Number = True
        Next file

        If Number = False Then
            AddBackArray GData, TPath
        End If

Set TFolder = Nothing

End Function

Public Sub AddBackArray(GData As Variant, ByVal AddWord As Variant)

Dim Number As Integer

Number = GetUBound(GData) + 1
ReDim Preserve GData(Number)
GData(Number) = AddWord

End Sub

Public Function GetUBound(GData As Variant) As Integer

On Error Resume Next

GetUBound = -1
GetUBound = UBound(GData)

End Function

Private Sub ExitRoomDo(Optional Boot As Boolean)

Dim Result As String

If Boot = False Then Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F55485944425F5F5D1E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId))

ChatRoom = False
ChatBox.Enabled = False
ChatBox.Timer1.Enabled = False
ChatBox.Hide
OpenRoom.Enabled = True
OpenRoom.UpdateRoomDo
OpenRoom.UpdateRoom.Enabled = True
UpdateText.Enabled = False
OpenRoom.Show
code = 0
ChatBox.ShowRoomSong.Caption = ""
ChatBox.ShowWhoMake.Caption = ""
NetWork = False

End Sub

Private Sub BootUser_Click()

Dim Result As String, People As String

If ChatBox.People_List.Text = "" Then Exit Sub

ChatBox.BootUser.Enabled = False

People = ChatBox.People_List.Text
People = Replace(People, " - 已準備", "")
People = Replace(People, " - Ready", "")

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F525F5F441E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16525F5F440D") + GetCode(People))

ChatBox.BootUser.Enabled = True

UpdateTextShow

End Sub

Private Sub ChooseSong_Click()
Dim OpenSong As String, Result As String

OpenSong = OpenFile("cbg", "Open")
If OpenSong = "" Then Exit Sub

LoadCbe OpenSong, True

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F435F5E571E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16535F54550D") + GetCode(CStr(CodeA)) + GetLink("1656595C555E515D550D") + GetCode(FindFileName(OpenSong)) + GetLink("16435F5E575E515D550D") + GetCode(GameMelody) + GetLink("1643595E5755420D") + GetCode(GameSinger) + GetLink("1647585F5D515B550D") + GetCode(GameAuthor))

UpdateTextShow

End Sub

Private Sub ExitRoom_Click()

ExitRoomDo

End Sub

Private Sub Form_Load()

SetLC
Timer1.Enabled = True

End Sub

Private Sub InPutBox_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then SendMsg_Click

End Sub

Private Sub SendMsg_Click()

If LastMsg = ChatBox.InPutBox.Text Then Exit Sub

ChatBox.SendMsg.Enabled = False

If ChatBox.InPutBox.Text <> "" Then UpdateTextShow GetLink("165D43570D") + GetCode(ChatBox.InPutBox.Text)
InPutBox.Text = ""

ChatBox.SendMsg.Enabled = True

LastMsg = ChatBox.InPutBox.Text

End Sub

Private Sub ShowTextBox_Change()

ChatBox.ShowTextBox.SelStart = Len(ChatBox.ShowTextBox.Text)

End Sub

Private Sub StartButton_Click()

Dim Result As String, CanStart As Boolean, FileData() As String, i As Long

If code = 0 Then Exit Sub

FindFile FileData, App.Path + "\Game\"

    For i = 0 To UBound(FileData)
        If Right(FileData(i), 4) = ".cbg" Then
            LoadCbe App.Path + "\Game\" + FileData(i), True
            If CodeA = code Then
                CanStart = True
                DeleteDir App.Path + "\Game\cma\"
                Exit For
            End If
        End If
    Next i

ChatBox.StartButton.Enabled = False

        If CanStart = True Then
            Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F42555154491E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16535F54550D"))
            UpdateTextShow
        End If

ChatBox.StartButton.Enabled = True

End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False

If ChatRoom = True Then InPutBox.SetFocus

End Sub

Private Sub UpdateText_Timer()

UpdateText.Enabled = False

UpdateTextShow

ChatBox.UpdateText.Enabled = True

End Sub

Public Sub UpdateTextShow(Optional Add As String)

Dim Result As String, Show() As String, ShowA() As String, ShowB() As String, People As String, Number As Long, GNumber As Long, FileData() As String, q As Long, OpenData As String, i As Long, CanStart As Boolean

Number = ChatBox.People_List.ListIndex

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F535851441E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("165859540D") + GetCode(GetHardId) + GetLink("16425F5F5D59540D") + GetCode(CStr(Roomid)) + Add)

    Show = Split(Result, "999bbb999bbb")
    For i = 0 To UBound(Show) - 1
        ShowA = Split(Show(i), "000bbb999bbb")
        Roomid = ShowA(2)

        If Len(ChatBox.ShowTextBox.Text) > 7000 Then ChatBox.ShowTextBox.Text = ""

        If InStr(ShowA(1), "3930373738373837353434353431") > 1 Then
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(Replace(ShowA(1), "3930373738373837353434353431", "")) + IIf(Language = 0, " 已進入房間", " Comes This Room.") + vbCrLf
        ElseIf InStr(ShowA(1), "1576497674679485") > 1 Then
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(Replace(ShowA(1), "1576497674679485", "")) + IIf(Language = 0, " 已離開房間", " Leaves This Room.") + vbCrLf
        ElseIf InStr(ShowA(1), "9465764512459586") > 1 Then
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(Replace(ShowA(1), "9465764512459586", "")) + IIf(Language = 0, " 已被踢離開房間", " Leaves This Room.") + vbCrLf
            If User = GetDCode(Replace(ShowA(1), "9465764512459586", "")) Then ExitRoomDo True
        ElseIf InStr(ShowA(1), "1474147585946715") > 1 Then
            ShowB = Split(ShowA(1), "1474147585946715")
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(ShowB(0)) + IIf(Language = 0, " 已更換歌曲為 ", " Changes Song To ") + GetDCode(ShowB(2)) + " - " + GetDCode(ShowB(3)) + " - " + GetDCode(ShowB(4)) + IIf(Language = 0, " 編號為 ", " SongCode ") + ShowB(1) + vbCrLf
            ChatBox.ShowRoomSong.Caption = GetDCode(ShowB(2)) + " - " + GetDCode(ShowB(3))
            ChatBox.ShowWhoMake.Caption = GetDCode(ShowB(4))
            code = Val(ShowB(1))
        ElseIf InStr(ShowA(1), "3164975894746315") > 1 Then
            ShowB = Split(ShowA(1), "3164975894746315")
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(ShowB(0)) + IIf(Language = 0, " 的成積為 ", " 's Result ") + GetDCode(ShowB(1)) + "P " + GetDCode(ShowB(2)) + "G " + GetDCode(ShowB(3)) + "C " + GetDCode(ShowB(4)) + "B " + GetDCode(ShowB(5)) + "M - " + GetDCode(ShowB(6)) + "Score" + vbCrLf
        ElseIf InStr(ShowA(1), "9641365214521469") > 1 Then
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(Replace(ShowA(1), "9641365214521469", "")) + IIf(Language = 0, " 的遊戲已結束了", " 's Game Already End.") + vbCrLf
        ElseIf InStr(ShowA(1), "7469741361465214") > 1 Then
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + IIf(Language = 0, " 系統訊息: ", " System Info: ") + GetDCode(Replace(ShowA(1), "7469741361465214", "")) + IIf(Language = 0, " 開始了遊戲", " Start The Game.") + vbCrLf
            ChatBox.Enabled = False
            ChatBox.Hide
            UpdateText.Enabled = False

            cmt.AMenu.Visible = False
            cmt.BMenu.Visible = False
            cmt.CMenu.Visible = False
            cmt.dmenu.Visible = False
            cmt.emenu.Visible = False
            cmt.fmenu.Visible = False
            cmt.gmenu.Visible = False
            cmt.MapMenu.Visible = False
            cmt.HighAdmin.Visible = False
            cmt.SetLanguage.Visible = False
            cmt.NetWork.Visible = False
            cmt.Frame2.Visible = False
            cmt.Frame5.Visible = False
            cmt.Frame1.Visible = False
            cmt.Frame3.Visible = False
            cmt.Frame4.Visible = False
            cma6.CloseAll True
            OpenRoom.Enabled = False
            OpenRoom.UpdateRoom.Enabled = False
            OpenRoom.Hide
            Inited = False
            InitedA = False
            cma4.Initialise cmt.MainPicture
            cma6.InitDI
            cma1.LoadSound
            cma5.LoadAcv
            Room = False
            cma3.NewFileDo
            UseMode = "game"
            NetWork = True
            
            FindFile FileData, App.Path + "\Game\"

                For q = 0 To UBound(FileData)
                    If Right(FileData(q), 4) = ".cbg" Then
                        LoadCbe App.Path + "\Game\" + FileData(q), True
                        If CodeA = code Then
                            CanStart = True
                            OpenData = App.Path + "\Game\" + FileData(q)
                            DeleteDir App.Path + "\Game\cma\"
                            Exit For
                        End If
                    End If
                Next q

            cma1.OpenSoundByte OpenData
            cmt.Height = 11520
            cmt.Enabled = True
            cmt.Show
            cma1.DoPlayOrStop

        Else
            ChatBox.ShowTextBox.Text = ChatBox.ShowTextBox.Text + ShowA(0) + IIf(Language = 0, " 說: ", " Say: ") + GetDCode(ShowA(1)) + vbCrLf
        End If
    Next i
    
Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F40555F405C551E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId))

    People_List.Clear
    Show = Split(Result, "000bbb999bbb")
    
    For i = 0 To UBound(Show) - 1
        If InStr(Show(i), "9146761996314526") > 1 Then
            People = GetDCode(Replace(Show(i), "9146761996314526", "")) + IIf(Language = 0, " - 已準備", " - Ready")
        Else
            People = GetDCode(Show(i))
            If People = ChatBox.ShowRoomHost.Caption Then People = People + IIf(Language = 0, " - 房主", " - Room Host")
        End If
        
        People_List.AddItem People
    Next i
    
    If ChatRoom = True Then
        If UBound(Show) - 1 > Number Then
             ChatBox.People_List.ListIndex = Number
        Else
             ChatBox.People_List.ListIndex = UBound(Show) - 1
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

ExitRoomDo

End Sub


Public Sub SetLC()

If Admin = False Then On Error Resume Next

        If Language = 0 Then
            ChatBox.SendMsg.Caption = "發送訊息"
            ChatBox.ShowRoomNumber_Label.Caption = "房間號碼"
            ChatBox.ShowRoomName_Label.Caption = "房間名稱"
            ChatBox.HostName_Label.Caption = "房主名稱"
            ChatBox.People_Label.Caption = "在線玩家"
            ChatBox.Caption = "房間資訊"
            ChatBox.ExitRoom.Caption = "離開房間"
            ChatBox.BootUser.Caption = "踢走玩家"
            ChatBox.StartButton.Caption = "開始/準備"
            ChatBox.ChooseSong.Caption = "選擇歌曲"
            ChatBox.Song_Label.Caption = "遊戲歌曲"
            ChatBox.WhoMake_Label.Caption = "歌曲作者"
            
        ElseIf Language = 1 Then
            ChatBox.SendMsg.Caption = "Send Message"
            ChatBox.ShowRoomNumber_Label.Caption = "RoomNumber"
            ChatBox.ShowRoomName_Label.Caption = "Room Name"
            ChatBox.HostName_Label.Caption = "Room Host"
            ChatBox.People_Label.Caption = "Online List"
            ChatBox.Caption = "Room Info"
            ChatBox.ExitRoom.Caption = "Leave Room"
            ChatBox.BootUser.Caption = "Boot User"
            ChatBox.StartButton.Caption = "Start/Ready"
            ChatBox.ChooseSong.Caption = "Choose Song"
            ChatBox.Song_Label.Caption = "Game Song"
            ChatBox.WhoMake_Label.Caption = "Author"
            
        End If
End Sub



