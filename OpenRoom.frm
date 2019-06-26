VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form OpenRoom 
   BorderStyle     =   1  '單線固定
   Caption         =   "建立房間"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   Icon            =   "OpenRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   14445
   StartUpPosition =   3  '系統預設值
   Visible         =   0   'False
   Begin VB.Frame GoRoom_Frame 
      Caption         =   "進入房間"
      Height          =   1455
      Left            =   8040
      TabIndex        =   8
      Top             =   4440
      Width           =   6255
      Begin VB.TextBox GoRoom_RoomName 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton GoRoom_OpenRoomButton 
         Caption         =   "進入房間"
         Height          =   1095
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox GoRoom_RoomPw 
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label GoRoom_RoomPW_Label 
         AutoSize        =   -1  'True
         Caption         =   "房間密碼"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   720
      End
      Begin VB.Label GoRoom_RoomName_Label 
         AutoSize        =   -1  'True
         Caption         =   "房間編號"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.Timer UpdateRoom 
      Interval        =   5000
      Left            =   5040
      Top             =   5160
   End
   Begin VB.CommandButton BackOut 
      Caption         =   "退出網路連線"
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      RowHeightMin    =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame OpenRoom_Frame 
      Caption         =   "建立房間"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   6255
      Begin VB.TextBox RoomPw 
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton OpenRoomButton 
         Caption         =   "建立房間"
         Height          =   1215
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox RoomName 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Text            =   "測試房間"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label RoomName_Label 
         AutoSize        =   -1  'True
         Caption         =   "房間名稱"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   450
         Width           =   720
      End
      Begin VB.Label RoomPW_Label 
         AutoSize        =   -1  'True
         Caption         =   "房間密碼"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   720
      End
   End
End
Attribute VB_Name = "OpenRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fso As New FileSystemObject

Private Sub ExitList()

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
Unload OpenRoom
cmt.Enabled = True
cmt.Show
NetWork = False
            cmt.AMenu.Visible = True
            cmt.BMenu.Visible = True
            cmt.CMenu.Visible = True
            cmt.dmenu.Visible = True
            cmt.emenu.Visible = True
            cmt.fmenu.Visible = True
            cmt.gmenu.Visible = True

End Sub

Private Sub BackOut_Click()

ExitList

End Sub

Private Sub Form_Load()

setBoard
UpdateRoomDo
OpenRoom.UpdateRoom.Enabled = True
SetLO
If Fso.FolderExists(App.Path + "\Game\") = False Then Fso.CreateFolder (App.Path + "\Game\")

End Sub

Private Sub setBoard()

MSFlexGrid1.cols = 7
MSFlexGrid1.rows = 1

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1500
MSFlexGrid1.ColWidth(6) = 4560

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.Text = IIf(Language = 0, "房間編號", "Number")
MSFlexGrid1.Col = 1
MSFlexGrid1.Text = IIf(Language = 0, "房間名稱", "Room Name")
MSFlexGrid1.Col = 2
MSFlexGrid1.Text = IIf(Language = 0, "房間密碼", "PassWord")
MSFlexGrid1.Col = 3
MSFlexGrid1.Text = IIf(Language = 0, "房主名稱", "Host Name")
MSFlexGrid1.Col = 4
MSFlexGrid1.Text = IIf(Language = 0, "房內人數", "PeoPle")
MSFlexGrid1.Col = 5
MSFlexGrid1.Text = IIf(Language = 0, "房間狀態", "Playing")
MSFlexGrid1.Col = 6
MSFlexGrid1.Text = IIf(Language = 0, "房間歌曲", "Song")

End Sub

Private Sub Form_Unload(Cancel As Integer)

ExitList

End Sub

Private Sub GoRoom_OpenRoomButton_Click()

Dim Result As String, Show() As String, ShowA() As String, Tmp As String

    OpenRoom.GoRoom_OpenRoomButton.Enabled = False
    OpenRoom.GoRoom_RoomName.Enabled = False
    OpenRoom.GoRoom_RoomPw.Enabled = False

RoomPassword = IIf(OpenRoom.GoRoom_RoomPw.Text <> "", OpenRoom.GoRoom_RoomPw.Text, "9898x9898")
RoomNumber = IIf(OpenRoom.GoRoom_RoomName.Text <> "", OpenRoom.GoRoom_RoomName.Text, "0")

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F575F425F5F5D1E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId))

If Result <> "0" And Result <> "1" And Result <> "2" Then
    Room = True

    RoomNumber = Val(OpenRoom.GoRoom_RoomName.Text)
    ChatBox.ShowTextBox.Text = ""
    ChatBox.ShowRoomSong.Caption = ""
    ChatBox.ShowWhoMake.Caption = ""
    ChatBox.ShowRoomNumber.Caption = OpenRoom.GoRoom_RoomName.Text
    ChatBox.ShowRoomName.Caption = Split(Result, "999bbb999")(0)
    ChatBox.ShowRoomName.Caption = GetDCode(ChatBox.ShowRoomName.Caption)
    If ChatBox.ShowRoomName.Caption = "9898x9898" Then ChatBox.ShowRoomName.Caption = ""
    ChatBox.ShowRoomHost.Caption = Split(Result, "999bbb999")(1)
    Roomid = Split(Result, "999bbb999")(2) - 1
    Tmp = Split(Result, "999bbb999")(3)
    If Tmp <> "" Then ChatBox.ShowRoomSong.Caption = GetDCode(Tmp)
    Tmp = Split(Result, "999bbb999")(4)
    If Tmp <> "" Then ChatBox.ShowWhoMake.Caption = GetDCode(Tmp)
    If ChatBox.ShowRoomHost.Caption = User Then ChatBox.StartButton.Visible = True
    OpenRoom.Hide
    OpenRoom.Enabled = False
    ChatBox.Show
    ChatBox.Enabled = True
    ChatBox.UpdateTextShow
    cmt.Enabled = False
    cmt.Hide
    ChatBox.UpdateText.Enabled = True
    OpenRoom.RoomName.Text = IIf(Language = 0, "測試房間", "Test Room")
    OpenRoom.RoomPw.Text = ""
    OpenRoom.GoRoom_RoomName.Text = "0"
    OpenRoom.GoRoom_RoomPw.Text = ""
    OpenRoom.UpdateRoom.Enabled = False
    ChatRoom = True
    
    If ChatBox.ShowRoomHost.Caption = User Then
        ChatBox.BootUser.Visible = True
        ChatBox.ChooseSong.Visible = True
    Else
        ChatBox.BootUser.Visible = False
        ChatBox.ChooseSong.Visible = False
    End If
    Tmp = Split(Result, "999bbb999")(5)
    code = Val(Tmp)
ElseIf Result = "1" Then
    MsgBox IIf(Language = 0, "遊戲已開始", "The Game Already Start"), vbYes, IIf(Language = 0, "系統訊息", "System Info")
ElseIf Result = "2" Then
    MsgBox IIf(Language = 0, "房間不存在", "Have Not Room"), vbYes, IIf(Language = 0, "系統訊息", "System Info")
Else
    MsgBox IIf(Language = 0, "密碼錯誤", "Password Wrong"), vbYes, IIf(Language = 0, "系統訊息", "System Info")
End If

If OpenRoom.GoRoom_RoomName.Text = "" Then OpenRoom.GoRoom_RoomName.Text = 0

    OpenRoom.GoRoom_OpenRoomButton.Enabled = True
    OpenRoom.GoRoom_RoomName.Enabled = True
    OpenRoom.GoRoom_RoomPw.Enabled = True

End Sub

Private Sub GoRoom_RoomName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then OpenRoom.GoRoom_RoomPw.SetFocus

End Sub

Private Sub GoRoom_RoomPw_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then GoRoom_OpenRoomButton_Click

End Sub

Private Sub MSFlexGrid1_LostFocus()

PressA = False

End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

PressA = False

End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 And PressA = True And timeGetTime - PressTime < 2000 Then
    
            MSFlexGrid1.Row = MSFlexGrid1.RowSel
            MSFlexGrid1.Col = 0
            
            If MSFlexGrid1.Text <> "房間編號" And MSFlexGrid1.Text <> "Number" Then OpenRoom.GoRoom_RoomName.Text = MSFlexGrid1.Text
            
            MSFlexGrid1.Col = 2
            
            If MSFlexGrid1.Text = "無密碼" Or MSFlexGrid1.Text = "No PassWord" Then
                MSFlexGrid1.Col = 5
                If MSFlexGrid1.Text = "未開始" Or MSFlexGrid1.Text = "Ready" Then GoRoom_OpenRoomButton_Click
            Else
                OpenRoom.GoRoom_RoomPw.SetFocus
            End If
    
    End If

PressA = True
PressTime = timeGetTime

End Sub

Private Sub OpenRoomButton_Click()

Dim Result As String

    OpenRoom.OpenRoomButton.Enabled = False
    OpenRoom.RoomName.Enabled = False
    OpenRoom.RoomPw.Enabled = False

RoomName = IIf(OpenRoom.RoomName.Text <> "", OpenRoom.RoomName.Text, "9898x9898")
RoomPassword = IIf(OpenRoom.RoomPw.Text <> "", OpenRoom.RoomPw.Text, "9898x9898")

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F5D515B55425F5F5D1E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D5E515D550D") + GetCode(RoomName) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId))

RoomNumber = Split(Result, "919194649794")(0)
Roomid = Split(Result, "919194649794")(1) - 1

If RoomNumber > 0 Then

    Room = True

    ChatBox.ShowTextBox.Text = ""
    ChatBox.ShowRoomNumber.Caption = RoomNumber
    ChatBox.ShowRoomHost.Caption = User
    ChatBox.ShowRoomName.Caption = OpenRoom.RoomName.Text
    
    ChatBox.ShowRoomSong.Caption = IIf(Language = 0, "未選擇歌曲", "Have Not Choose The Song")
    ChatBox.ShowWhoMake.Caption = IIf(Language = 0, "未選擇歌曲", "Have Not Choose The Song")
    
    OpenRoom.Hide
    OpenRoom.Enabled = False
    ChatBox.Show
    ChatBox.Enabled = True
    ChatBox.UpdateText.Enabled = True
    
    If ChatBox.ShowRoomHost.Caption = User Then
        ChatBox.BootUser.Visible = True
        ChatBox.ChooseSong.Visible = True
    Else
        ChatBox.BootUser.Visible = False
        ChatBox.ChooseSong.Visible = False
    End If
    
    ChatBox.StartButton.Visible = True
    ChatBox.UpdateTextShow
    ChatBox.UpdateText.Enabled = True
    OpenRoom.RoomName.Text = IIf(Language = 0, "測試房間", "Test Room")
    OpenRoom.RoomPw.Text = ""
    OpenRoom.GoRoom_RoomName.Text = "0"
    OpenRoom.GoRoom_RoomPw.Text = ""
    OpenRoom.UpdateRoom.Enabled = False
    ChatRoom = True
    ChatBox.ShowRoomSong.Caption = ""
    ChatBox.ShowWhoMake.Caption = ""
Else
    MsgBox IIf(Language = 0, "密碼錯誤", "Password Wrong"), vbYes, IIf(Language = 0, "系統訊息", "System Info")
End If

    OpenRoom.RoomName.Enabled = True
    OpenRoom.RoomPw.Enabled = True
    OpenRoom.OpenRoomButton.Enabled = True

End Sub

Private Sub RoomName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then RoomPw.SetFocus

End Sub

Private Sub RoomPw_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then OpenRoomButton_Click

End Sub

Public Sub SetLO()

If Admin = False Then On Error Resume Next

        If Language = 0 Then
            OpenRoom.Caption = "建立房間"
            OpenRoom.RoomName_Label.Caption = "房間名稱"
            OpenRoom.RoomPW_Label.Caption = "房間密碼"
            OpenRoom.RoomName.Text = "測試房間"
            OpenRoom.OpenRoomButton.Caption = "建立房間"
            
            OpenRoom.Caption = "進入房間"
            OpenRoom.GoRoom_RoomName_Label.Caption = "房間編號"
            OpenRoom.GoRoom_RoomPW_Label.Caption = "房間密碼"
            OpenRoom.GoRoom_OpenRoomButton.Caption = "進入房間"
            OpenRoom.BackOut.Caption = "退出網路連線"
            OpenRoom.OpenRoom_Frame.Caption = "建立房間"
            OpenRoom.GoRoom_Frame.Caption = "進入房間"
        ElseIf Language = 1 Then
            OpenRoom.Caption = "Open Room"
            OpenRoom.RoomName_Label.Caption = "Room Name"
            OpenRoom.RoomPW_Label.Caption = "Room Password"
            OpenRoom.RoomName.Text = "Test Room"
            OpenRoom.OpenRoomButton.Caption = "Open Room"
            
            OpenRoom.Caption = "Go Room"
            OpenRoom.GoRoom_RoomName_Label.Caption = "Room Number"
            OpenRoom.GoRoom_RoomPW_Label.Caption = "Room Password"
            OpenRoom.GoRoom_OpenRoomButton.Caption = "Go Room"
            OpenRoom.BackOut.Caption = "Leave NetWork"
            OpenRoom.OpenRoom_Frame.Caption = "Make Room"
            OpenRoom.GoRoom_Frame.Caption = "Go Room"

        End If
End Sub

Public Sub UpdateRoomDo()

Dim Result As String, Show() As String, ShowA() As String, i As Long, u As Long, a As Long

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F575544425F5F5D1E4058400F454355420D") + GetCode(User) + GetLink("165859540D") + GetCode(GetHardId))

Show = Split(Result, "3737191928286464")

If UBound(Show) < 0 Then Exit Sub

MSFlexGrid1.cols = 7
MSFlexGrid1.rows = UBound(Show) + 1

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.Text = IIf(Language = 0, "房間編號", "Number")
MSFlexGrid1.Col = 1
MSFlexGrid1.Text = IIf(Language = 0, "房間名稱", "Room Name")
MSFlexGrid1.Col = 2
MSFlexGrid1.Text = IIf(Language = 0, "房間密碼", "PassWord")
MSFlexGrid1.Col = 3
MSFlexGrid1.Text = IIf(Language = 0, "房主名稱", "Host Name")
MSFlexGrid1.Col = 4
MSFlexGrid1.Text = IIf(Language = 0, "房內人數", "People")
MSFlexGrid1.Col = 5
MSFlexGrid1.Text = IIf(Language = 0, "房間狀態", "Playing")
MSFlexGrid1.Col = 6
MSFlexGrid1.Text = IIf(Language = 0, "房間歌曲", "Song")

For i = LBound(Show) To UBound(Show) - 1

    ShowA = Split(Show(i), "917397346284628")
    
    MSFlexGrid1.Row = i + 1
    
    For u = 0 To MSFlexGrid1.cols - 1
        MSFlexGrid1.Col = u

        If u = 0 Then
    
            MSFlexGrid1.Text = ShowA(u)
            
        ElseIf u = 1 Then
    
            MSFlexGrid1.Text = GetDCode(ShowA(u))

        ElseIf u = 2 Then
        
            If ShowA(u) = "9898x9898" Then
                MSFlexGrid1.Text = IIf(Language = 0, "無密碼", "No PassWord")
            Else
                MSFlexGrid1.Text = IIf(Language = 0, "有密碼", "Need PassWord")
            End If
            
        ElseIf u = 3 Then
        
            MSFlexGrid1.Text = ShowA(u)
        ElseIf u = 4 Then
        
            MSFlexGrid1.Text = ShowA(u)

        ElseIf u = 5 Then
        
            If ShowA(u) = "0" Then
                MSFlexGrid1.Text = IIf(Language = 0, "未開始", "Ready")
            Else
                MSFlexGrid1.Text = IIf(Language = 0, "正在遊戲中", "Playing")
            End If
        ElseIf u = 6 Then
            If ShowA(u) <> "" And ShowA(u + 1) <> "" Then
                MSFlexGrid1.Text = GetDCode(ShowA(u)) + " - " + GetDCode(ShowA(u + 1))
            Else
                MSFlexGrid1.Text = IIf(Language = 0, "未選歌曲", "No Song")
            End If
        End If
    Next u
Next i

End Sub

Private Sub UpdateRoom_Timer()

UpdateRoomDo

End Sub
