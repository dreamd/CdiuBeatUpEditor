VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form cmt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   ClientHeight    =   11715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15450
   Enabled         =   0   'False
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   781
   ScaleMode       =   3  '像素
   ScaleWidth      =   1030
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "播放器"
      Height          =   1095
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
      Begin VB.CommandButton b4Space 
         Caption         =   "S<"
         Height          =   495
         Left            =   2400
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton AfterSpace 
         Caption         =   ">S"
         Height          =   495
         Left            =   7920
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton RightOne 
         Caption         =   ">"
         Height          =   495
         Left            =   7560
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton LeftOne 
         Caption         =   "<"
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin MSComctlLib.Slider Times 
         Height          =   495
         Left            =   3240
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   873
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider noteSpeed 
         Height          =   975
         Left            =   8400
         TabIndex        =   63
         ToolTipText     =   "箭頭音量"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   75
         SmallChange     =   50
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   5
      End
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   975
         Left            =   8760
         TabIndex        =   64
         ToolTipText     =   "歌曲音量"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   75
         SmallChange     =   50
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   5
      End
      Begin VB.Image Button 
         Height          =   735
         Index           =   3
         Left            =   1560
         Picture         =   "test.frx":628A
         ToolTipText     =   "停止"
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Button 
         Height          =   735
         Index           =   2
         Left            =   840
         Picture         =   "test.frx":7F20
         ToolTipText     =   "停止"
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Button 
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "test.frx":9BB6
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Button 
         Height          =   735
         Index           =   1
         Left            =   120
         Picture         =   "test.frx":B84C
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "普通設定"
      Height          =   6255
      Left            =   8760
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   5520
         Tag             =   "此程式為免費版本，版權持有人為藍翎，嚴禁擅自販賣。發怖網址為 https://github.com/dreamd/CdiuBeatUpEditor/releases"
         Top             =   4080
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   8
         Left            =   2400
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "136"
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   7
         Left            =   2400
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "135"
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   6
         Left            =   2400
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "134"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   5
         Left            =   2400
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "133"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   4
         Left            =   2400
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "132"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   3
         Left            =   2400
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "131"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Timer TimerRS 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5280
         Top             =   2040
      End
      Begin VB.Timer TimerLS 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4560
         Top             =   1920
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   2
         Left            =   2400
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "130"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "75"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4680
         Tag             =   $"test.frx":D4E2
         Top             =   1200
      End
      Begin VB.TextBox OK_Bpm 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Index           =   0
         Left            =   2400
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "150"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox OK_Offset 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Left            =   2400
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0.7"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Timer TimerL 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3360
         Top             =   1200
      End
      Begin VB.Timer TimerR 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4080
         Top             =   1200
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm9:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   6
         Left            =   240
         TabIndex        =   61
         Top             =   5640
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm8:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   5
         Left            =   240
         TabIndex        =   59
         Top             =   5040
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm7:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   4
         Left            =   240
         TabIndex        =   57
         Top             =   4440
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm6:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm5:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm4:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label_bpm3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm3:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label_bpm2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Bpm2:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label_Load 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   " Load Temp  "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   480
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   2250
      End
      Begin VB.Label Label_save 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         Caption         =   "  Save  "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   480
         Left            =   4200
         TabIndex        =   4
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label_offset 
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Offset:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label_bpm 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Main Bpm:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.PictureBox MainPicture 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   3  '像素
      ScaleWidth      =   1024
      TabIndex        =   5
      Top             =   0
      Width           =   15360
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "內部管理"
         Height          =   1695
         Left            =   1440
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox AddUser_text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label CancelAddUser 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Cancel  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   2640
            TabIndex        =   43
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label AddUser 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Save  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   600
            TabIndex        =   38
            Top             =   1080
            Width           =   1290
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "段落工具"
         Height          =   4815
         Left            =   8640
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   6735
         Begin VB.ListBox Team_List 
            Height          =   3480
            ItemData        =   "test.frx":D580
            Left            =   240
            List            =   "test.frx":D582
            TabIndex        =   33
            Top             =   480
            Width           =   6255
         End
         Begin VB.Label HideTeam 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Hide  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   4680
            TabIndex        =   32
            Top             =   4200
            Width           =   1245
         End
         Begin VB.Label UseTeam 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Paste  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   2760
            TabIndex        =   31
            Top             =   4200
            Width           =   1425
         End
         Begin VB.Label AddTeam 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Add  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   1080
            TabIndex        =   30
            Top             =   4200
            Width           =   1110
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "高級設定"
         ForeColor       =   &H00000000&
         Height          =   6255
         Left            =   7320
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox BeatUpList_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "c:\beatup.slk"
            Top             =   5040
            Width           =   4935
         End
         Begin VB.TextBox AllSongList_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "c:\擠學.slk"
            Top             =   4440
            Width           =   4935
         End
         Begin VB.TextBox ScriptFolder_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "c:\"
            Top             =   3840
            Width           =   4935
         End
         Begin VB.TextBox OggFolder_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "c:\"
            Top             =   3240
            Width           =   4935
         End
         Begin VB.TextBox MusicCode_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "km001"
            Top             =   2640
            Width           =   4935
         End
         Begin VB.TextBox Level_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   2040
            Width           =   4935
         End
         Begin VB.TextBox Author_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "Author"
            Top             =   1440
            Width           =   4935
         End
         Begin VB.TextBox Melody_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Melody"
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox Single_Text 
            Appearance      =   0  '平面
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   510
            Left            =   3000
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Singer"
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label BeatUpList_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "BeatUp List:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   42
            Top             =   5040
            Width           =   2130
         End
         Begin VB.Label AllSongList_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "All Song List:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   39
            Top             =   4440
            Width           =   2265
         End
         Begin VB.Label User_Text 
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   495
            Left            =   2520
            TabIndex        =   35
            Top             =   5640
            Visible         =   0   'False
            Width           =   4935
         End
         Begin VB.Label User_Set 
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "User:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   5640
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label CancelHighSetting 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Cancel  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   4440
            TabIndex        =   27
            Top             =   5640
            Width           =   1635
         End
         Begin VB.Label ScriptFolder_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Script Folder:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   20
            Top             =   3840
            Width           =   2310
         End
         Begin VB.Label OggFolder_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Ogg Folder:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   1980
         End
         Begin VB.Label MusicCode_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Music Code:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   18
            Top             =   2640
            Width           =   2100
         End
         Begin VB.Label Level_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Level:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   1035
         End
         Begin VB.Label Autho_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Author:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label Melody_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Melody:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label SaveHighSetting 
            Appearance      =   0  '平面
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            BorderStyle     =   1  '單線固定
            Caption         =   "  Save  "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   480
            Left            =   2280
            TabIndex        =   14
            Top             =   5640
            Width           =   1290
         End
         Begin VB.Label Singer_Set 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "Singer:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   450
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   3840
         Left            =   0
         Picture         =   "test.frx":D584
         ScaleHeight     =   256
         ScaleMode       =   3  '像素
         ScaleWidth      =   256
         TabIndex        =   6
         Top             =   18000
         Width           =   3840
      End
   End
   Begin VB.Menu AMenu 
      Caption         =   "檔案"
      Begin VB.Menu NewFile 
         Caption         =   "開新檔案"
         Enabled         =   0   'False
      End
      Begin VB.Menu openmusic 
         Caption         =   "開啟音樂"
      End
      Begin VB.Menu OpenSlk 
         Caption         =   "導入slk文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SaveSlkButton 
         Caption         =   "導出slk文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu OpenKbe 
         Caption         =   "導入kbe文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SaveKbe 
         Caption         =   "導出kbe文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu OpenDdr 
         Caption         =   "導入ddr文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SaveDdr 
         Caption         =   "導出ddr文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu OpenCbe 
         Caption         =   "導入cbe文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SaveCbe 
         Caption         =   "導出cbe文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu HighSave 
         Caption         =   "高級儲存"
         Enabled         =   0   'False
      End
      Begin VB.Menu CheckError 
         Caption         =   "檢查Slk中的錯誤"
      End
      Begin VB.Menu SaveAsCbg 
         Caption         =   "導出遊戲檔"
         Enabled         =   0   'False
      End
      Begin VB.Menu LoadAutoSave 
         Caption         =   "開啟臨時儲存"
      End
   End
   Begin VB.Menu BMenu 
      Caption         =   "設定"
      Enabled         =   0   'False
      Begin VB.Menu setting 
         Caption         =   "普通設定"
         Enabled         =   0   'False
      End
      Begin VB.Menu ProSetting 
         Caption         =   "高級設定"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu CMenu 
      Caption         =   "播放器"
      Enabled         =   0   'False
      Begin VB.Menu ShowOrHide 
         Caption         =   "顯示/隱藏"
         Enabled         =   0   'False
      End
      Begin VB.Menu PlayOrStop 
         Caption         =   "播放"
      End
      Begin VB.Menu EndSong 
         Caption         =   "停止"
      End
      Begin VB.Menu PlaySpace 
         Caption         =   "播放空白鍵部分"
      End
   End
   Begin VB.Menu dmenu 
      Caption         =   "模式"
      Enabled         =   0   'False
      Begin VB.Menu SetNormalMode 
         Caption         =   "正常模式"
      End
      Begin VB.Menu SetSeeMode 
         Caption         =   "觀戰模式"
         Enabled         =   0   'False
      End
      Begin VB.Menu SetGameMode 
         Caption         =   "遊戲模式"
      End
   End
   Begin VB.Menu emenu 
      Caption         =   "功能"
      Enabled         =   0   'False
      Begin VB.Menu AutoSpace16 
         Caption         =   "自動16空白鍵"
      End
      Begin VB.Menu AutoSpace24 
         Caption         =   "自動24空白鍵"
      End
      Begin VB.Menu AutoSpace32 
         Caption         =   "自動32空白鍵"
      End
      Begin VB.Menu AutoSpace48 
         Caption         =   "自動48空白鍵"
      End
      Begin VB.Menu DelSpace 
         Caption         =   "後面空白鍵全刪"
      End
      Begin VB.Menu SpaceOut16 
         Caption         =   "後面空白鍵退後16"
      End
      Begin VB.Menu AllRandomKey 
         Caption         =   "全部箭頭隨機"
      End
      Begin VB.Menu LRRandomKey 
         Caption         =   "全部箭頭左右隨機"
      End
      Begin VB.Menu KeyLeft 
         Caption         =   "全部箭頭左移一格"
      End
      Begin VB.Menu KeyRight 
         Caption         =   "全部箭頭右移一格"
      End
      Begin VB.Menu AllKeyOne 
         Caption         =   "全部箭頭變為1"
      End
      Begin VB.Menu DelAllNote 
         Caption         =   "刪除全部箭頭"
      End
      Begin VB.Menu RxMove 
         Caption         =   "直線移動"
      End
   End
   Begin VB.Menu fmenu 
      Caption         =   "編輯"
      Enabled         =   0   'False
      Begin VB.Menu CutButton 
         Caption         =   "剪下"
      End
      Begin VB.Menu CopyButton 
         Caption         =   "復制"
      End
      Begin VB.Menu PushUpButton 
         Caption         =   "貼上"
      End
      Begin VB.Menu DelButton 
         Caption         =   "刪除"
      End
      Begin VB.Menu UnDo 
         Caption         =   "復原"
      End
      Begin VB.Menu AutoFillAll 
         Caption         =   "自動填滿"
      End
      Begin VB.Menu SomeKeyRandom 
         Caption         =   "部分箭頭隨機"
      End
      Begin VB.Menu SomeKeyLRRandom 
         Caption         =   "部分箭頭左右隨機"
      End
      Begin VB.Menu SomeKeyLeft 
         Caption         =   "部分箭頭左移一格"
      End
      Begin VB.Menu SomeKeyRight 
         Caption         =   "部分箭頭右移一格"
      End
      Begin VB.Menu SomeKeyBeOne 
         Caption         =   "部分箭頭變為1"
      End
   End
   Begin VB.Menu gmenu 
      Caption         =   "段落"
      Enabled         =   0   'False
      Begin VB.Menu gshow 
         Caption         =   "顯示/隱藏"
      End
   End
   Begin VB.Menu MapMenu 
      Caption         =   "地圖"
      Begin VB.Menu ChangeBack1 
         Caption         =   "籃球場"
      End
      Begin VB.Menu ChangeBack2 
         Caption         =   "Lafesta"
      End
      Begin VB.Menu ChangeBack3 
         Caption         =   "漢江公園"
      End
      Begin VB.Menu ChangeBack4 
         Caption         =   "E251C"
      End
      Begin VB.Menu ChangeBack5 
         Caption         =   "光華門"
      End
      Begin VB.Menu ChangeBack6 
         Caption         =   "紫禁場"
      End
      Begin VB.Menu ChangeBack7 
         Caption         =   "萬聖節"
      End
      Begin VB.Menu ChangeBack8 
         Caption         =   "結他場"
      End
      Begin VB.Menu ChangeBack9 
         Caption         =   "滑雪場"
      End
      Begin VB.Menu ChangeBack11 
         Caption         =   "Basic House"
      End
      Begin VB.Menu SetByUser 
         Caption         =   "使用者定義"
      End
   End
   Begin VB.Menu HighAdmin 
      Caption         =   "管理"
      Visible         =   0   'False
   End
   Begin VB.Menu SetLanguage 
      Caption         =   "Language"
      Begin VB.Menu SetChinese 
         Caption         =   "Chinese"
      End
      Begin VB.Menu SetEnglish 
         Caption         =   "English"
      End
   End
   Begin VB.Menu NetWork 
      Caption         =   "網路連線"
      Visible         =   0   'False
   End
   Begin VB.Menu MakeItUp 
      Caption         =   "合成"
      Visible         =   0   'False
   End
   Begin VB.Menu MakeItOut 
      Caption         =   "拆開"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "cmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MouseStartX As Long
Public MouseStartY As Long

Public MouseMoveX As Long
Public MouseMoveY As Long

Public MouseDown As Boolean
Public MouseMove As Boolean

Private Sub AddTeam_Click()

Me.MousePointer = 15

End Sub
Private Sub Form_Resize()
    Select Case Me.WindowState
        Case 2  '?中最大化
            Me.WindowState = 0
    End Select
    
End Sub
Sub ReSize(Width As Long, Height As Long)
    cmt.MainPicture.Width = cmt.MainPicture.Width * Width / cmt.MainPicture.ScaleWidth
    cmt.MainPicture.Height = cmt.MainPicture.Height * Height / cmt.MainPicture.ScaleHeight
    Me.Height = Me.Height * Height / Me.ScaleHeight
    Me.Width = Me.Width * Width / Me.ScaleWidth
End Sub

Private Sub AfterSpace_Click()

Dim ToNowBeat As Long, ToOffset As Single, i As Long, Check As Boolean

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

    For i = ToNowBeat + 1 To TotalBeat - 1
        If GData(i * 8 + 6) = True Then Check = True: Exit For
    Next i

If Check = True Then cmt.Times.value = PData(i - 1) + OffSet

End Sub

Private Sub AfterSpace_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerRS.Enabled = True

End Sub

Private Sub AfterSpace_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerRS.Enabled = False

End Sub

Private Sub b4Space_Click()

Dim ToNowBeat As Long, ToOffset As Single, i As Long, Check As Boolean

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

    For i = ToNowBeat - 1 To 0 Step -1
        If GData(i * 8 + 6) = True Then Check = True: Exit For
    Next i

If Check = True Then
    cmt.Times.value = PData(i - 1) + OffSet
End If

End Sub

Private Sub b4Space_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerLS.Enabled = True

End Sub

Private Sub b4Space_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerLS.Enabled = False

End Sub

Private Sub CancelAddUser_Click()

cmt.Frame5.Visible = False

End Sub

Private Sub ChangeBack11_Click()

ChangeMapByUser 10

End Sub

Private Sub CheckError_Click()

Dim NowFolder As String, FileName As String, LoadFile As String

If Admin = False Then On Error Resume Next

LoadFile = cma2.OpenFile("Slk", "Open")
If LoadFile <> "" Then LoadFile = ClearName(LoadFile)

NowFolder = Replace(cma1.SongPath, cma3.FindFileName(cma1.SongPath), "")
FileName = cma3.FindFileName(cma1.SongPath)

If LoadFile <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then cma7.CheckSlk LoadFile


End Sub

Private Sub DelAllNote_Click()

cma3.AutoSave
cma3.DelANote

End Sub

Private Sub AddUser_Click()

cma3.DelFile App.Path + "\Newuser\User.cdiu"
cma7.Saveuser cmt.AddUser_text.Text
cmt.Frame5.Visible = False

End Sub

Private Sub CancelHighSetting_Click()

If Admin = False Then On Error Resume Next

cmt.Single_Text.Text = Singer
cmt.Melody_Text.Text = Melody
cmt.Author_Text.Text = Author
cmt.Level_Text.Text = level
cmt.MusicCode_Text.Text = MusicCode
cmt.OggFolder_Text.Text = OggF
cmt.ScriptFolder_Text.Text = ScrF

cmt.Frame3.Visible = False
cmt.Frame2.Visible = True
cmt.Frame1.Visible = False
cmt.Frame4.Visible = False
cmt.Frame5.Visible = False

End Sub

Private Sub ChangeBack1_Click()

ChangeMapByUser 0

End Sub

Private Sub ChangeBack2_Click()

ChangeMapByUser 1

End Sub

Private Sub ChangeBack3_Click()

ChangeMapByUser 2

End Sub

Private Sub ChangeBack4_Click()

ChangeMapByUser 3

End Sub

Private Sub ChangeBack5_Click()

ChangeMapByUser 4

End Sub

Private Sub ChangeBack6_Click()

ChangeMapByUser 5

End Sub

Private Sub ChangeBack7_Click()

ChangeMapByUser 6

End Sub

Private Sub ChangeBack8_Click()

ChangeMapByUser 7

End Sub

Private Sub ChangeBack9_Click()

ChangeMapByUser 8

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Check As VbMsgBoxResult, SaveFile As String, Result As String

If Admin = False Then On Error Resume Next

If NetWork = False Then

        If TotalBeat > 0 Then
            If cma2.CheckCombo(TotalBeat) > 0 Then
                Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
            Else
            cma6.CloseAll
            End If
        Else
            cma6.CloseAll
        End If
        
        
        If Check = vbYes Then
        
            SaveFile = cma2.OpenFile("cdiu", "Save")
        
            If SaveFile <> "" Then cma3.CbeOut SaveFile
                cma6.CloseAll
        ElseIf Check = vbNo Then
            cma6.CloseAll
        ElseIf Check = vbCancel Then
            Cancel = -1
        End If
Else

Result = GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F55485944425F5F5D1E4058400F454355420D") + GetCode(User) + GetLink("16425F5F5D0D") + GetCode(CStr(RoomNumber)) + GetLink("16425F5F5D40470D") + GetCode(RoomPassword) + GetLink("165859540D") + GetCode(GetHardId))

End If

End Sub

Private Sub gshow_Click()

cma2.Frame4Show

End Sub

Private Sub HideTeam_Click()

cma2.Frame4Show

End Sub

'Private Sub JoinRoom_Click()
'GoRoom.Show: GoRoom.GoRoomName.SetFocus
'End Sub

Private Sub Label_Load_Click()

cma3.NewFileDo
cma3.AutoLoadDo

End Sub

Private Sub LoadAutoSave_Click()

cma3.AutoLoadDo

End Sub

'Private Sub MakeRoom_Click()
'OpenRoom.Show: OpenRoom.RoomName.SetFocus
'End Sub

Private Sub ShowBox()

    cmt.Hide
    cmt.Enabled = False
    Room = True
    cma1.CloseSound
    cma4.UnloadD3D
    cma6.UnloadDI
    
    OpenRoom.Show
    OpenRoom.Enabled = True

End Sub

Private Sub NetWork_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If TotalBeat > 0 Then
    If cma2.CheckCombo(TotalBeat) > 0 Then
        Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
    Else
        ShowBox
    End If
Else
    ShowBox
End If

If Check = vbYes Then

    SaveFile = cma2.OpenFile("cdiu", "Save")

    If SaveFile <> "" Then cma3.CbeOut SaveFile

    ShowBox
ElseIf Check = vbNo Then
    ShowBox
ElseIf Check = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub OK_Bpm_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

If Admin = False Then On Error Resume Next

If KeyCode = 13 And index = 0 Then cmt.OK_Offset.SetFocus

If KeyCode = 13 And index = 1 Then cmt.OK_Bpm.Item(2).SetFocus

If KeyCode = 13 And index = 2 Then cmt.OK_Bpm.Item(3).SetFocus

If KeyCode = 13 And index = 3 Then cmt.OK_Bpm.Item(4).SetFocus

If KeyCode = 13 And index = 4 Then cmt.OK_Bpm.Item(5).SetFocus

If KeyCode = 13 And index = 5 Then cmt.OK_Bpm.Item(6).SetFocus

If KeyCode = 13 And index = 6 Then cmt.OK_Bpm.Item(7).SetFocus

If KeyCode = 13 And index = 7 Then cmt.OK_Bpm.Item(8).SetFocus

If KeyCode = 13 And index = 8 Then Label_save_Click

End Sub


Private Sub OK_Offset_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then cmt.OK_Bpm.Item(1).SetFocus

End Sub

Private Sub Button_Click(index As Integer)

If Admin = False Then On Error Resume Next

Select Case index
    Case 0
        cma1.DoPlayOrStop
    Case 1
        cma1.DoPlayOrStop
    Case 2
        cma2.EndTheSong
    Case 3
        Button.Item(0).Visible = True
        Frame1.Visible = True
        OK_Bpm.Item(0).SetFocus
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        If Mode = "playing" Then cma1.DoPlayOrStop
End Select

End Sub

Private Sub RxMove_Click()

SetRx = SetRx + 1
cma3.AutoSave

End Sub

Private Sub LRRandomKey_Click()

cma5.GoLeftRightRandom
cma3.AutoSave

End Sub

Private Sub AllRandomKey_Click()

cma5.GoAllRandom
cma3.AutoSave

End Sub

Private Sub KeyLeft_Click()

cma5.AllKeyLeft
cma3.AutoSave

End Sub

Private Sub KeyRight_Click()

cma5.AllKeyRight
cma3.AutoSave

End Sub

Private Sub SaveAsCbg_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("cbg", "Save")

If SaveFile <> "" Then cma3.GameFile SaveFile

End Sub

Private Sub SetByUser_Click()

ChangeMapByUser 9

End Sub

Private Sub SetChinese_Click()

Language = 0
cma5.SetL
cma5.SaveSetting False, True

End Sub

Private Sub SetEnglish_Click()

Language = 1
cma5.SetL
cma5.SaveSetting False, True

End Sub

Private Sub SomeKeyRandom_Click()

cma5.DoSomeRandom
cma3.AutoSave

End Sub

Private Sub SomeKeyLRRandom_Click()

cma5.DoSomeLRRandom
cma3.AutoSave

End Sub

Private Sub SomeKeyBeOne_Click()

cma5.DoSomeOneKey
cma3.AutoSave

End Sub

Private Sub AllKeyOne_Click()

cma5.OneKey
cma3.AutoSave

End Sub

Private Sub AutoFillAll_Click()

cma5.GoAutoDo
cma3.AutoSave

End Sub

Private Sub AutoSpace16_Click()

Me.MousePointer = 9

End Sub

Private Sub AutoSpace24_Click()

Me.MousePointer = 8

End Sub


Private Sub AutoSpace32_Click()

Me.MousePointer = 7

End Sub


Private Sub AutoSpace48_Click()

Me.MousePointer = 6

End Sub

Private Sub CopyButton_Click()

cma5.GoCopy

End Sub

Private Sub SomeKeyRight_Click()

cma5.DoSomeKeyRight
cma3.AutoSave

End Sub

Private Sub SomeKeyLeft_Click()

cma5.DoSomeKeyLeft
cma3.AutoSave

End Sub

Private Sub CutButton_Click()

cma5.GoCopy True
cma5.GoDelete

End Sub

Private Sub DelButton_Click()

cma5.GoDelete
cma3.AutoSave

End Sub

Private Sub DelSpace_Click()

Me.MousePointer = 5

End Sub

Private Sub PushUpButton_Click()

cma5.PushOutKey
cma3.AutoSave

End Sub

Private Sub SaveHighSetting_Click()

cma5.SaveSetting
cma3.AutoSave

End Sub

Private Sub SpaceOut16_Click()

Me.MousePointer = 10

End Sub

Private Sub PlaySpace_Click()

Me.MousePointer = 12

End Sub

Private Sub EndSong_Click()

cma2.EndTheSong

End Sub

Private Sub Form_Unload(Cancel As Integer)

cma6.CloseAll

End Sub

Private Sub Form_Load()

Dim i As Long

Randomize
Code1 = Fix(9999 * Rnd) Xor 12341234
'ReDim MData(1000000)

'If App.PrevInstance = True Then cma7.ExitExe
Me.Top = 0: Me.Left = 0: UseMode = "see"
cmt.User_Text = User
cmt.Caption = "Cdiu BeatUp Editor v" + VerionA + "." + VerionB
OggF = cmt.OggFolder_Set
ScrF = cmt.ScriptFolder_Set
ASL = cmt.AllSongList_Set
BUL = cmt.BeatUpList_Set
ReDim SaveSelect(0)
BpmSet(0) = cmt.OK_Bpm(0).Text
BpmSet(1) = cmt.OK_Bpm(1).Text
BpmSet(2) = cmt.OK_Bpm(2).Text
BpmSet(3) = cmt.OK_Bpm(3).Text
BpmSet(4) = cmt.OK_Bpm(4).Text
BpmSet(5) = cmt.OK_Bpm(5).Text
BpmSet(6) = cmt.OK_Bpm(6).Text
BpmSet(7) = cmt.OK_Bpm(7).Text
BpmSet(8) = cmt.OK_Bpm(8).Text

OffsetSet = cmt.OK_Offset.Text
Mode = "close"
SetRx = 1
TotalBeat = 0
ChooseBackGround = 1
Language = 0

ReDim GData(7)
ReDim SData(7)

    For i = 0 To 6
        SData(i) = 4
    Next i
    
cma2.FreeLibrary hLibR
cma2.FreeLibrary hLibL
cma2.MakeTemp "fmod.dll", TempPath, "DLL", "FMOD"
cma2.MakeTemp "zlib.dll", TempPath, "DLL", "ZLIB"
hLibR = cma2.LoadLibrary(CStr(TempPath + "fmod.dll"))
hLibL = cma2.LoadLibrary(CStr(TempPath + "zlib.dll"))


cma7.LoadUserFile
'cma7.LoadUser App.Path + "\User.cdiu"
cma7.CheckUser

'If Code1 <> (Code2 Xor 12341234) Then End
If Code1 = (Code2 Xor 12341234) Then cma4.Initialise cmt.MainPicture
cma1.LoadSound
cma5.LoadAcv


cma6.InitDI
cma2.SaveData

ReDim PData(1)

FindBeatTime PData

ExitMsg = False
UseMode = "see"
SaveFileINI = "Setting.ini"
cma5.LoadFast
cma5.LoadINI
cma5.SetL
'cma7.CheckAdmin
cma4.ChangeMapByUser ChooseBackGround

If Admin = False Then
    Me.Height = 12300
    Me.Width = 15450
Else
    cmt.MakeItUp.Visible = True
    cmt.MakeItOut.Visible = True
End If

cmt.Timer1.Enabled = True
ReSize 1024, 768
Timer2.Enabled = True
End Sub


Private Sub Label_save_Click()

Dim i As Long

If Admin = False Then On Error Resume Next
    
    For i = 0 To 2
        If IsNumeric(cmt.OK_Bpm.Item(i).Text) = False Then Exit Sub
        If cmt.OK_Bpm.Item(i).Text > 250 Then Exit Sub
        If cmt.OK_Bpm.Item(i).Text < 1 Then Exit Sub
    Next i
    
If IsNumeric(cmt.OK_Offset.Text) = False Then Exit Sub
If cmt.OK_Offset.Text > 1.7 Then Exit Sub
If cmt.OK_Offset.Text <= 0 Then Exit Sub

cma2.SaveData True

cma3.CheckChangeBpm

TotalBeat = SoundL * CBT
cma3.AutoSave
cmt.OpenSlk.Enabled = True
cmt.OpenKbe.Enabled = True
cmt.OpenDdr.Enabled = True
cmt.OpenCbe.Enabled = True

End Sub

Private Sub LeftOne_Click()

Dim ToNowBeat As Long

If Admin = False Then On Error Resume Next

FindBeatTime PData

ToNowBeat = FindWhichBeat(PData, cmt.Times.value)

cmt.Times.value = PData(ToNowBeat - 1) + OffSet

End Sub

Private Sub RightOne_Click()

Dim ToNowBeat As Long

If Admin = False Then On Error Resume Next

FindBeatTime PData

ToNowBeat = FindWhichBeat(PData, cmt.Times.value)

cmt.Times.value = PData(ToNowBeat + 1) + OffSet

End Sub

Private Sub RightOne_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerR.Enabled = True

End Sub

Private Sub RightOne_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerR.Enabled = False

End Sub

Private Sub LeftOne_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerL.Enabled = True

End Sub

Private Sub LeftOne_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

TimerL.Enabled = False

End Sub

Private Sub MainPicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

        If Mode = "playing" Then Exit Sub

            If Button = 1 Then
                MouseMove = False
                MouseDown = True
                MouseStartX = x
                MouseStartY = y
            End If

End Sub

Private Sub MainPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

        If Mode = "playing" Then Exit Sub

                If MouseDown = True Then
                        MouseMove = True
                        MouseMoveX = x
                        MouseMoveY = y
                End If
        
End Sub

Private Sub MainPicture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim cY  As Long, NowBeat As Long, cX As Long, j As Long, OffsetK As Single, ToNowBeat As Long, ToOffset As Single

        If Mode = "playing" Then Exit Sub
        
If Admin = False Then On Error Resume Next

'MsgBox CStr(y)

If Button = 2 Then PopupMenu fmenu: Exit Sub

MouseDown = False

                If MouseMove Then
                        If cma4.SelectedKeys > 0 Then
                        
                            ReDim SaveSelect(cma4.SelectedKeys)
                                For j = 1 To UBound(SaveSelect)
                                        SaveSelect(j) = cma4.GetKeyS(CLng(j))
                                Next j
                        Else
                            ReDim SaveSelect(0)
                        End If
                    Exit Sub
                End If

cma6.CheckTime ToNowBeat, ToOffset

        If UseMode = "normal" Then
            y = y - 108
            x = x - 10
            If y < 0 Then Exit Sub
            If y > 616 Then Exit Sub
            cY = CLng(Mid(CStr(y / 64), 1, 1))
            cX = ((x + 1 - 20) / 32) + ToOffset + ToNowBeat

            If cX < 47 Then Exit Sub
            If y > 416 Then cY = 7
            
            If y > 448 Then cY = 8
            If y > 469 Then cY = 9
            If y > 490 Then cY = 10
            If y > 511 Then cY = 11
            If y > 532 Then cY = 12
            If y > 553 Then cY = 13
            If y > 574 Then cY = 14
            If y > 595 Then cY = 15
            
            If Me.MousePointer = 9 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.AutoDoSpace cX, 16
                cma3.AutoSave
                
            ElseIf Me.MousePointer = 8 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.AutoDoSpace cX, 24
                cma3.AutoSave
                
            ElseIf Me.MousePointer = 7 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.AutoDoSpace cX, 32
                cma3.AutoSave
                
            ElseIf Me.MousePointer = 6 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.AutoDoSpace cX, 48
                cma3.AutoSave

            ElseIf Me.MousePointer = 5 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.AutoDelSpace cX
                cma3.AutoSave
                
            ElseIf Me.MousePointer = 10 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.SpaceBack16 cX
                cma3.AutoSave

            ElseIf Me.MousePointer = 12 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma1.DoPlayOrStop cX, (cma5.NextSpace(cX)) - cX
                
            ElseIf Me.MousePointer = 15 Then
                Me.MousePointer = 0
                If cY <> 6 Then Exit Sub
                cma5.SaveUnDo
                cma5.SaveThisSpace cX
            ElseIf cY >= 8 Then
                cma5.SaveUnDo
                SetSData cY - 8, cX
                cma3.AutoSave
            Else
                cma5.SaveUnDo
                GData(cX * 8 + cY) = IIf(GData(cX * 8 + cY) = True, False, True)
                cma3.AutoSave
            End If
        Else
            Exit Sub
        End If

cmt.SaveSlkButton.Enabled = True
cmt.SaveKbe.Enabled = True
cmt.SaveDdr.Enabled = True
cmt.SaveCbe.Enabled = True
cmt.HighSave.Enabled = True
cmt.SaveAsCbg.Enabled = True

End Sub

Private Sub HighSave_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(SoundL) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("slk", "Save")

If SaveFile <> "" Then cma3.HighSaveDo SaveFile

End Sub


Private Sub openmusic_Click()

Dim NowFolder As String, FileName As String

If Admin = False Then On Error Resume Next

cma1.SongPath = cma2.OpenFile("Ogg", "Open", "Abm")
If cma1.SongPath <> "" Then cma1.SongPath = ClearName(cma1.SongPath)

NowFolder = Replace(cma1.SongPath, cma3.FindFileName(cma1.SongPath), "")
FileName = cma3.FindFileName(cma1.SongPath)

If cma1.SongPath <> "" And cma2.Cdiu_File("check", NowFolder, FileName) = False Then cma1.OpenSound cma1.SongPath

End Sub

Private Sub NewFile_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) > 0 Then
    Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
Else
    cma3.NewFileDo
End If

If Check = vbYes Then

SaveFile = cma2.OpenFile("cdiu", "Save")

If SaveFile <> "" Then cma3.CbeOut SaveFile

    cma3.NewFileDo
ElseIf Check = vbNo Then
    cma3.NewFileDo
End If

End Sub

Private Sub openslk_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If TotalBeat > 0 Then
    If cma2.CheckCombo(TotalBeat) > 0 Then
        Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
    Else
        cma3.OpenSlkDo
    End If
Else
    cma3.OpenSlkDo
End If

If Check = vbYes Then

    SaveFile = cma2.OpenFile("cdiu", "Save")

    If SaveFile <> "" Then cma3.CbeOut SaveFile

    End
ElseIf Check = vbNo Then
    cma3.OpenSlkDo
ElseIf Check = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub opencbe_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If TotalBeat > 0 Then
    If cma2.CheckCombo(TotalBeat) > 0 Then
        Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
    Else
        cma3.OpenCbeDo
    End If
Else
    cma3.OpenCbeDo
End If

If Check = vbYes Then

    SaveFile = cma2.OpenFile("cdiu", "Save")

    If SaveFile <> "" Then cma3.CbeOut SaveFile

    End
ElseIf Check = vbNo Then
    cma3.OpenCbeDo
ElseIf Check = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub openddr_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If TotalBeat > 0 Then
    If cma2.CheckCombo(TotalBeat) > 0 Then
        Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
    Else
        cma3.OpenDdrDo
    End If
Else
    cma3.OpenDdrDo
End If

If Check = vbYes Then

    SaveFile = cma2.OpenFile("cdiu", "Save")

    If SaveFile <> "" Then cma3.CbeOut SaveFile

    End
ElseIf Check = vbNo Then
    cma3.OpenDdrDo
ElseIf Check = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub openkbe_Click()

Dim Check As VbMsgBoxResult, SaveFile As String

If Admin = False Then On Error Resume Next

If TotalBeat > 0 Then
    If cma2.CheckCombo(TotalBeat) > 0 Then
        Check = MsgBox(IIf(Language = 0, "檔案是否需要儲存?", "Do You Need To Save?"), vbYesNoCancel, IIf(Language = 0, "系統訊息", "System Info"))
    Else
        cma3.OpenKbeDo
    End If
Else
    cma3.OpenKbeDo
End If

If Check = vbYes Then

    SaveFile = cma2.OpenFile("cdiu", "Save")

    If SaveFile <> "" Then cma3.CbeOut SaveFile

    End
ElseIf Check = vbNo Then
    cma3.OpenKbeDo
ElseIf Check = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub PlayOrStop_Click()

cma1.DoPlayOrStop

End Sub

Private Sub SaveSlkButton_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("Slk", "Save")

If SaveFile <> "" Then cma3.CbeToSlk SaveFile

End Sub

Private Sub Savekbe_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("kbe", "Save")

If SaveFile <> "" Then cma3.CbeToKbe SaveFile

End Sub

Private Sub Saveddr_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("ddr", "Save")

If SaveFile <> "" Then cma3.CbeToDdr SaveFile

End Sub

Private Sub Savecbe_Click()

Dim SaveFile As String

If Admin = False Then On Error Resume Next

If cma2.CheckCombo(TotalBeat) = 0 Then Exit Sub

SaveFile = cma2.OpenFile("cdiu", "Save")

If SaveFile <> "" Then cma3.CbeOut SaveFile

End Sub

Private Sub SetNormalMode_Click()

UseMode = "normal"
cma2.UseModeChange

End Sub

Private Sub MakeItUp_Click()

cma5.MakeItUpDo

End Sub


Private Sub MakeItOut_Click()

cma5.MakeItOutDo

End Sub

Private Sub SetSeeMode_Click()

UseMode = "see"
cma2.UseModeChange

End Sub


Private Sub SetGameMode_Click()

        UseMode = "game"
        cma2.UseModeChange
        

End Sub

Private Sub Setting_Click()

If Admin = False Then On Error Resume Next

        cmt.Button.Item(0).Visible = True
        cmt.Frame1.Visible = True
        cmt.OK_Bpm.Item(0).SetFocus
        cmt.Frame2.Visible = False
        cmt.Frame3.Visible = False
        cmt.Frame4.Visible = False
        cmt.Frame5.Visible = False
        If Mode = "playing" Then cma1.DoPlayOrStop

End Sub

Private Sub ProSetting_Click()

If Admin = False Then On Error Resume Next

        cmt.Button.Item(0).Visible = True
        cmt.Frame1.Visible = False
        cmt.Frame2.Visible = False
        cmt.Frame3.Visible = True
        cmt.Frame4.Visible = False
        cmt.Frame5.Visible = False
        If Mode = "playing" Then cma1.DoPlayOrStop

End Sub

Private Sub ShowOrHide_Click()

cma2.CheckDFrame2

End Sub

Private Sub Timer1_Timer()

cmt.Timer1.Enabled = False
cma4.Render
cma4.LoopRender

End Sub

Private Sub Timer2_Timer()
    Me.Timer2.Enabled = False
    MsgBox IIf(Language = 0, Me.Timer2.Tag, Me.Timer1.Tag), 0, IIf(Language = 0, "系統訊息", "System Info")
End Sub

Private Sub UnDo_Click()

cma5.UnDoIt

End Sub

Private Sub TimerR_Timer()

RightOne_Click

End Sub

Private Sub TimerL_Timer()

LeftOne_Click

End Sub

Private Sub TimerLS_Timer()

b4Space_Click

End Sub

Private Sub TimerRS_Timer()

AfterSpace_Click

End Sub

Private Sub Times_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Mouse = True

End Sub

Private Sub Times_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Admin = False Then On Error Resume Next

cma1.ChangeTime Times.value
Mouse = False

KeyTime(0) = 0
KeyTime(1) = 0
KeyTime(2) = 0
KeyTime(3) = 0
KeyTime(4) = 0
KeyTime(5) = 0
KeyTime(6) = 0

End Sub

Public Sub Times_Scroll()

MouseMove = False

End Sub

Private Sub Times_change()

CurrPos = cmt.Times.value
Times_Scroll

End Sub


Private Sub UseTeam_Click()

Dim ChooseNumber As Long

If Admin = False Then On Error Resume Next

ChooseNumber = cmt.Team_List.ListIndex

Clipboard.SetText FastTeam(ChooseNumber)
cma5.PushOutKey
cma3.AutoSave

End Sub

Private Sub HighAdmin_Click()

cmt.Frame5.Visible = IIf(cmt.Frame5.Visible = True, False, True)

End Sub
