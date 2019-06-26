Attribute VB_Name = "cma4"
Option Explicit

Public D3DX As D3DX8
Public dx  As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8

Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Current As Integer

Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont

Dim BackGround2(10) As Direct3DTexture8, SpaceBar As Direct3DTexture8

Dim Crazy As Direct3DTexture8

Dim TRM1() As Byte, TRM2() As Byte

Dim TLM1() As Byte, TLM2() As Byte

Dim RM1 As Direct3DTexture8, RM2 As Direct3DTexture8

Dim LM1 As Direct3DTexture8, LM2 As Direct3DTexture8

Dim LMiss As Direct3DTexture8, RMiss As Direct3DTexture8

Dim Combo(10 To 39) As Direct3DTexture8, Perfect(4) As Direct3DTexture8

Dim KeyStrip(3) As TLVERTEX, KeyImageS As Direct3DTexture8, KeyImage(7) As Direct3DTexture8, KeySImage(5) As Direct3DTexture8

Dim KeyPress(5) As Direct3DTexture8, Power As Direct3DTexture8, KeyLine(5) As Direct3DTexture8

Dim Spb As Direct3DTexture8, Sp As Direct3DTexture8

Dim Yup As Direct3DTexture8, Ycup As Direct3DTexture8

Dim Bup As Direct3DTexture8, Bcup As Direct3DTexture8

Dim CB(5) As Direct3DTexture8, CYBACK(1) As Direct3DTexture8

Dim BYBUPY(2) As Direct3DTexture8, READY(4) As Direct3DTexture8

Dim NormalBack As Direct3DTexture8, ShowBeatupLogo(1) As Direct3DTexture8

Dim NormalBackP As Direct3DTexture8, ALine As Direct3DTexture8

Dim NORMALRX As Direct3DTexture8

Dim L3 As Direct3DTexture8, R3 As Direct3DTexture8
Dim TL3() As Byte, TR3() As Byte

Dim SelectedKey(800) As Integer

Public SelectedKeys As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Dim Fso As New FileSystemObject

Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    Color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Public Sub UnloadD3D()
    
    Dim i As Integer, u As Integer

    Set D3DX = Nothing
    Set dx = Nothing
    Set D3D = Nothing
    Set D3DDevice = Nothing
    
    Set MainFont = Nothing
    Set MainFontDesc = Nothing
    Set fnt = Nothing
    Set Crazy = Nothing
    
    Set RM1 = Nothing
    Set RM2 = Nothing
    Set LM1 = Nothing
    Set LM2 = Nothing
    
    Set LMiss = Nothing
    Set RMiss = Nothing
    
    For i = 0 To 10
        Set BackGround2(i) = Nothing
    Next i
    
    For i = 0 To 4
        Set Perfect(i) = Nothing
    Next i
        
    For i = 10 To 39
        Set Combo(i) = Nothing
    Next i
    
    For i = 0 To 7
        Set KeyImage(i) = Nothing
    Next i
    
    For i = 0 To 5
        Set KeySImage(i) = Nothing
    Next i
    
    For i = 0 To 5
        Set KeyPress(i) = Nothing
    Next i

    For i = 0 To 5
        Set KeyLine(i) = Nothing
    Next i

    For i = 0 To 5
        Set CB(i) = Nothing
    Next i
    
    For i = 0 To 4
        Set READY(i) = Nothing
    Next i

    Set KeyImageS = Nothing

    For i = 0 To 1
        Set CYBACK(i) = Nothing
    Next i
    
    For i = 0 To 2
        Set BYBUPY(i) = Nothing
    Next i
    
    For i = 0 To 1
        Set ShowBeatupLogo(i) = Nothing
    Next i
    
    Set Power = Nothing
    Set SpaceBar = Nothing

    Set Spb = Nothing
    Set Sp = Nothing
    
    Set Yup = Nothing
    Set Ycup = Nothing
    
    Set Bup = Nothing
    Set Bcup = Nothing
    
    Set NormalBack = Nothing
    Set NormalBackP = Nothing
    
    Set ALine = Nothing
    
    Set NORMALRX = Nothing
    
    Set L3 = Nothing
    Set R3 = Nothing
    
End Sub

Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, Color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX

    CreateTLVertex.x = x
    CreateTLVertex.y = y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.specular = specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
    
End Function

Public Function Initialise(DisplayA As PictureBox) As Boolean

If Inited = True Then Exit Function
Inited = True

If Admin = False Then On Error Resume Next

Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS

'Set dx = New DirectX8
Set dx = CreateObject("DIRECT.DirectX8.0")
Set D3D = dx.Direct3DCreate()
Set D3DX = New D3DX8


D3D.GetAdapterDisplayMode 0, DispMode

D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP

D3DWindow.BackBufferCount = 1
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.BackBufferHeight = DisplayA.ScaleHeight
D3DWindow.BackBufferWidth = DisplayA.ScaleWidth
D3DWindow.hDeviceWindow = DisplayA.hwnd
D3DWindow.Windowed = 1

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DisplayA.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)

D3DDevice.SetVertexShader FVF
D3DDevice.SetRenderState D3DRS_LIGHTING, False
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

fnt.Name = "Tahoma"
fnt.Size = 12

Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

End Function

Public Sub SaveDDS(Name As String, Add() As Byte)

                cma5.CreateDir App.Path + "\user\"
                
                If Fso.FileExists(App.Path + "\user\" + Name) = False Then
                    Open App.Path + "\user\" + Name For Binary As #1
                        Put #1, , Add
                    Close #1
                End If

End Sub

Public Sub ChangeMapByUser(ChangeMap As Long)

If ChangeMap = ChooseBackGround Then Exit Sub
If Fso.FileExists(App.Path + "\user\BG.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\LM1.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\LM2.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\L3.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\RM1.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\RM2.dds") = False Then Exit Sub
If Fso.FileExists(App.Path + "\user\R3.dds") = False Then Exit Sub

    If ChangeMap = 9 Then
        Set BackGround2(9) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\BG.dds")
        
        Set LM1 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\LM1.dds")
        Set LM2 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\LM2.dds")
        Set L3 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\L3.dds")
        
        Set RM1 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\RM1.dds")
        Set RM2 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\RM2.dds")
        Set R3 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\user\R3.dds")
    End If

    If ChooseBackGround = 9 Then
        Set LM1 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TLM1(0), CLng(UBound(TLM1) + 1))
        Set LM2 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TLM2(0), CLng(UBound(TLM2) + 1))
        Set RM1 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TRM1(0), CLng(UBound(TRM1) + 1))
        Set RM2 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TRM2(0), CLng(UBound(TRM2) + 1))
        Set L3 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TL3(0), CLng(UBound(TL3) + 1))
        Set R3 = D3DX.CreateTextureFromFileInMemory(D3DDevice, TR3(0), CLng(UBound(TR3) + 1))
    End If

ChooseBackGround = ChangeMap

End Sub

Public Sub CheckFileIn(Name As String, Add() As Byte)

Dim i As Long, u As Long, Number As Long, Which As String

AddBackArray CheckFile, Name

    For i = 0 To 10
        If Name = "BG" + CStr(i + 1) + "\" + CStr(i + 1) + ".dds" Then
            If i = 1 Then SaveDDS "BG.dds", Add
            Set BackGround2(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1))
            Exit Sub
        End If
    Next i

Number = 10

        For i = 1 To 3
                For u = 0 To 9
                        If Name = "COMBO\C" + CStr(i) + CStr(u) + ".dds" Then Set Combo(Number) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
                        Number = Number + 1
                Next u
        Next i

        For i = LBound(KeyImage) To UBound(KeyImage)
            If Name = "DDS\K" + CStr(i) + ".dds" Then Set KeyImage(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i

        For i = LBound(KeySImage) To UBound(KeySImage)
            If Name = "DDS\S" + CStr(i) + ".dds" Then Set KeySImage(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i
        
        For i = LBound(KeyPress) To UBound(KeyPress)
            If Name = "DDS\L" + CStr(i) + ".dds" Then Set KeyPress(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i
        
        For i = 0 To 4
            Select Case i
                Case 0: Which = "PERFECT"
                Case 1: Which = "GREAT"
                Case 2: Which = "COOL"
                Case 3: Which = "BAD"
                Case 4: Which = "MISS"
            End Select
            If Name = "DDS\" + Which + ".dds" Then Set Perfect(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i
    
        For i = 0 To 5
            Select Case i
                Case 0: Which = "CB"
                Case 1: Which = "CE"
                Case 2: Which = "CA"
                Case 3: Which = "CT"
                Case 4: Which = "CU"
                Case 5: Which = "CP"
            End Select
            If Name = "DDS\" + Which + ".dds" Then Set CB(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i

        For i = 0 To 1
            Select Case i
                Case 0: Which = "CYBACK"
                Case 1: Which = "CBBACK"
            End Select
            If Name = "DDS\" + Which + ".dds" Then Set CYBACK(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i
        
        For i = 0 To 2
            Select Case i
                Case 0: Which = "BYBUPY"
                Case 1: Which = "BYBUPB"
                Case 2: Which = "BYBUPF"
            End Select
            If Name = "DDS\" + Which + ".dds" Then Set BYBUPY(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i
        
        For i = 0 To 4
            If Name = "DDS\READY" + CStr(i) + ".dds" Then Set READY(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i

        For i = 0 To 1
            If Name = "DDS\LOGO" + CStr(i) + ".dds" Then Set ShowBeatupLogo(i) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        Next i

        If Name = "DDS\RP.dds" Then
            Set KeyLine(0) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1))
            Set KeyLine(2) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1))
            Set KeyLine(3) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1))
            Set KeyLine(5) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        End If

        If Name = "DDS\MP.dds" Then
            Set KeyLine(1) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1))
            Set KeyLine(4) = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
        End If
        
        Select Case Name
            Case "BMP\ALINE.bmp":        Set ALine = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "BMP\L3.dds":           Set L3 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "L3.dds", Add: TL3 = Add: Exit Sub
            Case "BMP\NORMALBACK.bmp":   Set NormalBack = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "BMP\NORMALRX.bmp":     Set NORMALRX = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "BMP\P.jpg":            Set NormalBackP = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "BMP\R3.dds":           Set R3 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "R3.dds", Add: TR3 = Add: Exit Sub
            
            Case "DDS\SPACEBAR.dds":     Set SpaceBar = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\SPACE.dds":        Set KeyImageS = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\SLINE.dds":        Set Spb = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\SPOWER.dds":       Set Sp = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            
            Case "DDS\POWER.dds":        Set Power = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\YUP.dds":          Set Yup = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\YCUP.dds":         Set Ycup = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\BUP.dds":          Set Bup = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\BCUP.dds":         Set Bcup = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\CRAZY.dds":        Set Crazy = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            
            Case "DDS\RM1.dds":          Set RM1 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "RM1.dds", Add: TRM1 = Add: Exit Sub
            Case "DDS\RM2.dds":          Set RM2 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "RM2.dds", Add: TRM2 = Add: Exit Sub
            
            Case "DDS\LM1.dds":          Set LM1 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "LM1.dds", Add: TLM1 = Add: Exit Sub
            Case "DDS\LM2.dds":          Set LM2 = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): SaveDDS "LM2.dds", Add: TLM2 = Add: Exit Sub
            
            Case "DDS\RMISS.dds":        Set RMiss = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            Case "DDS\LMISS.dds":        Set LMiss = D3DX.CreateTextureFromFileInMemory(D3DDevice, Add(0), CLng(UBound(Add) + 1)): Exit Sub
            
            
            Case "SOUND\BEAT.ogg": cma1.LoadEffect Name, Add
            Case "SOUND\GREAT.ogg": cma1.LoadEffect Name, Add
            Case "SOUND\READY.ogg": cma1.LoadEffect Name, Add
            Case "SOUND\SPACE.ogg": cma1.LoadEffect Name, Add
            Case "SOUND\START.ogg": cma1.LoadEffect Name, Add
            Case "SOUND\MISS.ogg": cma1.LoadEffect Name, Add
            
            Case "SLK\LIST1.txt": ReDim Slk1(UBound(Add)): cq.IcyCopyMemory ByVal VarPtr(Slk1(0)), ByVal VarPtr(Add(0)), UBound(Add) + 1
            Case "SLK\LIST2.txt": ReDim Slk2(UBound(Add)): cq.IcyCopyMemory ByVal VarPtr(Slk2(0)), ByVal VarPtr(Add(0)), UBound(Add) + 1
            Case "SLK\SLK2HEADER.txt": ReDim Slk3(UBound(Add)): cq.IcyCopyMemory ByVal VarPtr(Slk3(0)), ByVal VarPtr(Add(0)), UBound(Add) + 1
            Case "SLK\SLKHEADER.txt": ReDim Slk4(UBound(Add)): cq.IcyCopyMemory ByVal VarPtr(Slk4(0)), ByVal VarPtr(Add(0)), UBound(Add) + 1
        End Select

End Sub

Public Sub Render()

Dim ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset

Initialise cmt.MainPicture

RenderBeginScene

If UseMode = "see" Then
    ShowBackGround
    ShowCrazy 287, 653
    CheckStart
    CheckComboAdd
    RenderCombo 92, 354
    CheckMoreCombo
    RenderFinish
    If ToNowBeat > 1 Then CheckPress 1
    RenderLeftKey 423, 281
    RenderRightKey 423, 672
    RenderL 419, 0
    RenderR 419, 910
    RenderSpaceKey 644, 371, 280
    If ToNowBeat > 1 Then CheckPress 0
ElseIf UseMode = "game" Then
    ShowBackGround
    ShowCrazy 287, 653
    CheckStart
    CheckComboAdd
    RenderCombo 92, 354
    CheckMoreCombo
    RenderFinish
    If ToNowBeat > 1 Then CheckPress 1
    RenderLeftKey 423, 281
    RenderRightKey 423, 672
    RenderL 419, 0
    RenderR 419, 910
    RenderSpaceKey 644, 371, 280
    If ToNowBeat > 1 Then CheckPress 0
ElseIf UseMode = "normal" Then
    ShowNormalBack 0, 108
    If Mode <> "playing" Then ShowNormalRX 108
    If ToNowBeat > 1 Then CheckPress 1
    ShowNormalNote 108
    If ToNowBeat > 1 Then CheckPress 0
End If

ShowText
RenderEndScene
cma6.DXKeyboard
UpdateFps

End Sub

Public Sub ShowMiss()

If Admin = False Then On Error Resume Next

Dim StartX As Single, StartY As Single

StartX = 30
StartY = 425

D3DDevice.SetTexture 0, LMiss
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(100, 100, 100), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 512, StartY, 0, 1, RGB(100, 100, 100), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 256, 0, 1, RGB(100, 100, 100), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 512, StartY + 256, 0, 1, RGB(100, 100, 100), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

StartX = 674
StartY = 426

D3DDevice.SetTexture 0, RMiss
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(100, 100, 100), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 512, StartY, 0, 1, RGB(100, 100, 100), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 256, 0, 1, RGB(100, 100, 100), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 512, StartY + 256, 0, 1, RGB(100, 100, 100), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub ShowCrazy(StartX As Single, StartY As Single)

If Admin = False Then On Error Resume Next

D3DDevice.SetTexture 0, Crazy
KeyStrip(0) = CreateTLVertex(StartX - 32, StartY - 32, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 32, StartY - 32, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX - 32, StartY + 32, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 32, StartY + 32, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

StartX = 615
StartY = 422

D3DDevice.SetTexture 0, RM2
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 128, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 128, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

StartX = StartX + 128

D3DDevice.SetTexture 0, RM1
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 255, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 255, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

StartX = 285
StartY = 421

D3DDevice.SetTexture 0, LM2
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 128, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 128, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

StartX = 30

D3DDevice.SetTexture 0, LM1
KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 255, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 255, StartY + 255, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub CheckPress(UMode As Integer)

Dim ToNowBeat As Long, ToOffset As Single, cX As Single, clX As Single, cY As Single, k As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then Exit Sub
CheckGData ToNowBeat + 1

For k = 0 To 5

If UseMode = "see" Or UseMode = "game" Then

    If k < 3 Then
        cY = (k) * 64 + 423
        cX = 667
        clX = 672
    ElseIf k > 2 And k < 6 Then
        cY = (k - 3) * 64 + 423
        cX = 281
        clX = 0
    End If
    
ElseIf UseMode = "normal" Then
        cY = (k) * 64 + 108
        cX = -10
End If

        If (KeyTime(k) <= ToNowBeat + ToOffset And KeyTime(k) + 2 > ToNowBeat + ToOffset) Or (GData((ToNowBeat - 1) * 8 + k) = True And UseMode <> "game") Then
        
            Select Case UMode
                Case 0:
                    If UseMode = "game" Then
                        If ((GameCheck((ToNowBeat - 1) * 8 + k) = 1) Or (GameCheck((ToNowBeat - 1) * 8 + k) = 2)) And GData((ToNowBeat - 1) * 8 + k) = True Then KeyPower k, cX, cY, ToNowBeat - 1, KeyTime(k)
                    ElseIf UseMode = "game" Then
                        KeyPower k, cX, cY, ToNowBeat - 1
                    Else
                        If GData((ToNowBeat - 1) * 8 + k) = True Then KeyPower k, cX, cY, ToNowBeat - 1
                    End If
                Case 1
            End Select
        End If
Next k

For k = 0 To 5

If UseMode = "see" Or UseMode = "game" Then

    If k < 3 Then
        cY = (k) * 64 + 423
        cX = 672
        clX = 672
    ElseIf k > 2 And k < 6 Then
        cY = (k - 3) * 64 + 423
        cX = 281
        clX = 0
    End If
    
ElseIf UseMode = "normal" Then
        cY = (k) * 64 + 108
        cX = -10
End If

        If KeyTime(k) + 1 <= ToNowBeat + ToOffset And KeyTime(k) + 2 > ToNowBeat + ToOffset And UMode = 0 Then KeyPressShow k, cX, cY, True
        If UBound(GData) >= (ToNowBeat - 1) * 8 + k Then If GData((ToNowBeat - 1) * 8 + k) = True And UseMode <> "game" And UMode = 0 Then KeyPressShow k, cX, cY
        
        If (KeyTime(k) <= ToNowBeat + ToOffset And KeyTime(k) + 2 > ToNowBeat + ToOffset) Or (GData(ToNowBeat * 8 + k) = True And UseMode <> "game") Then
        If k < 3 Then cX = cX - 5
            Select Case UMode
                Case 0
                    If UseMode = "game" Then
                        If ((GameCheck(ToNowBeat * 8 + k) = 1) Or (GameCheck(ToNowBeat * 8 + k) = 2)) And GData(ToNowBeat * 8 + k) = True Then KeyPower k, cX, cY, ToNowBeat, KeyTime(k)
                    ElseIf UseMode = "see" Then
                        KeyPower k, cX, cY, ToNowBeat
                    Else
                        If GData(ToNowBeat * 8 + k) = True Then KeyPower k, cX, cY, ToNowBeat
                    End If
                Case 1
                    If UseMode = "game" Then
                        KeyLineShow k, clX, cY, IIf(UseMode = "normal", 2000, 555), ToNowBeat, KeyTime(k)
                    Else
                        KeyLineShow k, clX, cY, IIf(UseMode = "normal", 2000, 555), ToNowBeat
                    End If
            End Select
        End If
Next k

If UseMode = "see" Or UseMode = "game" Then
    cX = 371
    cY = 644
    clX = 375
ElseIf UseMode = "normal" Then
    cX = -106
    cY = 493
    clX = -50
End If

        If (KeyTime(6) < ToNowBeat + ToOffset And KeyTime(6) + 1 > ToNowBeat + ToOffset) Or (GData(ToNowBeat * 8 + 6) = True And UseMode <> "game") Then
            Select Case UMode
                Case 0:
                    If UseMode = "game" Then
                        If ((GameCheck(ToNowBeat * 8 + 6) = 1) Or (GameCheck(ToNowBeat * 8 + 6) = 2)) Then SpacePower cX, cY
                    Else
                        SpacePower cX, cY
                    End If
                Case 1: SpacePowerBack clX, cY, IIf(UseMode = "normal", 2100, 0)
            End Select
            
        End If

End Sub

Public Sub SpacePower(ByVal StartX As Single, ByVal StartY As Single)

Dim Sx As Single, tx As Single, sY As Single, ty As Single, SelectedT As Boolean, ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then Exit Sub

        If cmt.MouseMove = True Then
                If cmt.MouseStartX > cmt.MouseMoveX Then
                    tx = cmt.MouseStartX
                    Sx = cmt.MouseMoveX
                Else
                    Sx = cmt.MouseStartX
                    tx = cmt.MouseMoveX
                End If
                
                If cmt.MouseStartY > cmt.MouseMoveY Then
                    ty = cmt.MouseStartY
                    sY = cmt.MouseMoveY
                Else
                    sY = cmt.MouseStartY
                    ty = cmt.MouseMoveY
                End If
            End If
            
            SelectedT = 31 > Sx And 31 < tx And 560 + 31 > sY And 560 + 31 < ty And cmt.MouseMove And UseMode = "normal"
            
            
            StartX = StartX + 139.5
            StartY = StartY + 6
            
            D3DDevice.SetTexture 0, Sp
            KeyStrip(0) = CreateTLVertex(StartX - 128 + ToOffset * 32, StartY - 128 + ToOffset * 32, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
            KeyStrip(1) = CreateTLVertex(StartX + 128 - ToOffset * 32, StartY - 128 + ToOffset * 32, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
            KeyStrip(2) = CreateTLVertex(StartX - 128 + ToOffset * 32, StartY + 128 - ToOffset * 32, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
            KeyStrip(3) = CreateTLVertex(StartX + 128 - ToOffset * 32, StartY + 128 - ToOffset * 32, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
            
                        If SelectedT = True Then
                            SelectedKeys = SelectedKeys + 1
                            SelectedKey(SelectedKeys) = ToNowBeat * 8 + 6
                        End If
            
End Sub

Public Sub SpacePowerBack(StartX As Single, StartY As Single, More As Single)

If Admin = False Then On Error Resume Next

D3DDevice.SetTexture 0, Spb
KeyStrip(0) = CreateTLVertex(StartX + 11, StartY - 5, 0, 1, RGB(255, 255, 255), 0, 0, 0)
KeyStrip(1) = CreateTLVertex(StartX + 279 - 12 + More, StartY - 5, 0, 1, RGB(255, 255, 255), 0, 1, 0)
KeyStrip(2) = CreateTLVertex(StartX + 11, StartY + 22 + 7, 0, 1, RGB(255, 255, 255), 0, 0, 1)
KeyStrip(3) = CreateTLVertex(StartX + 279 - 12 + More, StartY + 22 + 7, 0, 1, RGB(255, 255, 255), 0, 1, 1)
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub KeyPressShow(PressKey As Long, ByVal StartX As Single, ByVal StartY As Single, Optional Press As Boolean)

Dim ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then Exit Sub

If UseMode = "normal" Then
    If PressKey = 1 Or PressKey = 4 Then StartX = StartX + 1.5
ElseIf UseMode = "see" Or UseMode = "game" Then
    If PressKey = 1 Then
        StartX = StartX - 5
    ElseIf PressKey = 0 Or PressKey = 2 Then
        StartX = StartX - 2
    ElseIf PressKey = 4 Then
        StartX = StartX + 3
    ElseIf PressKey = 3 Or PressKey = 5 Then
        StartX = StartX - 2
    End If
End If

If Press = True Then ToOffset = ToNowBeat + ToOffset - KeyTime(PressKey) - 1

If ToOffset > 1 Or ToOffset < 0 Then Exit Sub

StartX = StartX + 38
StartY = StartY + 32

                D3DDevice.SetTexture 0, KeyPress(PressKey)
                KeyStrip(0) = CreateTLVertex(StartX - 30 - ToOffset * 15, StartY - 30 - ToOffset * 15, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 30 + ToOffset * 15, StartY - 30 - ToOffset * 15, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX - 30 - ToOffset * 15, StartY + 30 + ToOffset * 15, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 30 + ToOffset * 15, StartY + 30 + ToOffset * 15, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub ShowSpaceBar(StartX As Single, StartY As Single)

If Admin = False Then On Error Resume Next

                    D3DDevice.SetTexture 0, SpaceBar
                    KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(StartX + 512, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(StartX, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(StartX + 512, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub


Public Sub KeyLineShow(PressKey As Long, StartX As Single, cY As Single, Width As Single, Optional Beat As Long, Optional PressBeat As Single)

Dim ToNowBeat As Long, ToOffset As Single, cYM As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset

If PressBeat <> 0 Then
    ToOffset = ToNowBeat + IIf(ToOffset < 0, (1 + ToOffset), ToOffset) - PressBeat
Else
    ToOffset = ToNowBeat - Beat + ToOffset
End If

If ToOffset > 0.8 Then Exit Sub

If UseMode = "normal" Then cYM = 4

                    D3DDevice.SetTexture 0, KeyLine(PressKey)
                    KeyStrip(0) = CreateTLVertex(StartX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(StartX + Width, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(StartX, cY + 64 + cYM, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(StartX + Width, cY + 64 + cYM, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub KeyPower(Which As Long, ByVal cX As Single, ByVal cY As Single, Beat As Long, Optional PressBeat As Single)

Dim Sx As Single, tx As Single, sY As Single, ty As Single, SelectedT As Boolean, ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset

If PressBeat <> 0 Then
    ToOffset = ToNowBeat + IIf(ToOffset < 0, (1 + ToOffset), ToOffset) - PressBeat
Else
    ToOffset = ToNowBeat - Beat + ToOffset
End If

If ToOffset > 1.2 Then Exit Sub

        If cmt.MouseMove = True Then
                If cmt.MouseStartX > cmt.MouseMoveX Then
                    tx = cmt.MouseStartX
                    Sx = cmt.MouseMoveX
                Else
                    Sx = cmt.MouseStartX
                    tx = cmt.MouseMoveX
                End If
                
                If cmt.MouseStartY > cmt.MouseMoveY Then
                    ty = cmt.MouseStartY
                    sY = cmt.MouseMoveY
                Else
                    sY = cmt.MouseStartY
                    ty = cmt.MouseMoveY
                End If
            End If
            
            
            cX = cX + 35
            cY = cY + 30
            
            SelectedT = cX + 31 > Sx And cX + 31 < tx And cY + 31 > sY And cY + 31 < ty And cmt.MouseMove And UseMode = "normal"
            If MData((ToNowBeat - 1) * 8 + Which) = True Then ToOffset = 0
                    D3DDevice.SetTexture 0, Power
                    KeyStrip(0) = CreateTLVertex(cX - (2.5 - ToOffset) * 70, cY - (2.5 - ToOffset) * 70, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(245, 245, 245)), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(cX + (2.5 - ToOffset) * 70, cY - (2.5 - ToOffset) * 70, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(245, 245, 245)), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(cX - (2.5 - ToOffset) * 70, cY + (2.5 - ToOffset) * 70, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(245, 245, 245)), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(cX + (2.5 - ToOffset) * 70, cY + (2.5 - ToOffset) * 70, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(245, 245, 245)), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                    
                        If SelectedT = True Then
                            SelectedKeys = SelectedKeys + 1
                            SelectedKey(SelectedKeys) = ToNowBeat * 8 + Which
                        End If

End Sub

Public Sub RenderL(StartY As Single, StartX As Single)

If Admin = False Then On Error Resume Next

                                    D3DDevice.SetTexture 0, L3
                                    KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(StartX + 113, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(StartX, StartY + 197, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(StartX + 113, StartY + 197, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub RenderR(StartY As Single, StartX As Single)

If Admin = False Then On Error Resume Next

                                    D3DDevice.SetTexture 0, R3
                                    KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(StartX + 114, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(StartX, StartY + 197, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(StartX + 114, StartY + 197, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Sub

Public Sub ShowNormalRX(StartY As Single)

Dim i As Long, cX As Single, ToOffset As Single, ToNowBeat As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat + 33

        For i = 0 To 33
                            cX = (i - ToOffset) * 32 - 4
                                    
                                    If (i + ToNowBeat) Mod 2 = (SetRx + 1) Mod 2 And cX > 70 Then
                                    
                                    D3DDevice.SetTexture 0, NORMALRX
                                    
                                    KeyStrip(0) = CreateTLVertex(cX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 1, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX, StartY + 620, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 1, StartY + 620, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                    
                                    TextRect.Left = cX + 16
                                    TextRect.Right = cX + 16 + 64
                                    TextRect.Top = 89
                                    D3DX.DrawText MainFont, &HFFFFFFFF, CStr(i + ToNowBeat), TextRect, 0
                                    
                        End If
        Next i
        
                                    TextRect.Left = 0
                                    TextRect.Right = 130
                                    TextRect.Top = 495
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "空白鍵", "Space Bar"), TextRect, 0
                                    TextRect.Top = 527
                                    D3DX.DrawText MainFont, &HFFFFFFFF, "Finish", TextRect, 0
                                    TextRect.Top = 557
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "9鍵速度", "Note 9 Speed"), TextRect, 0
                                    TextRect.Top = 578
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "6鍵速度", "Note 6 Speed"), TextRect, 0
                                    TextRect.Top = 599
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "3鍵速度", "Note 3 Speed"), TextRect, 0
                                    TextRect.Top = 620
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "7鍵速度", "Note 7 Speed"), TextRect, 0
                                    TextRect.Top = 641
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "4鍵速度", "Note 4 Speed"), TextRect, 0
                                    TextRect.Top = 662
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "1鍵速度", "Note 1 Speed"), TextRect, 0
                                    TextRect.Top = 683
                                    D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "S鍵速度", "Note Space Speed"), TextRect, 0
                                    TextRect.Top = 706
                                    D3DX.DrawText MainFont, &HFFFFFFFF, "Bpm", TextRect, 0

                                    

End Sub

Public Sub ShowNormalNote(StartY As Single)

Dim i As Long, cX As Single, cY As Single, k As Integer, ToNowBeat As Long, ToOffset As Single, KeySizeX As Single, KeySizeY As Single, Sx As Single, tx As Single, sY As Single, ty As Single, SelectedT As Boolean, WhichKey As String, o As Integer

If Admin = False Then On Error Resume Next

SelectedKeys = 0

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 2 Then ToNowBeat = 2
CheckGData ToNowBeat + 33

        If cmt.MouseMove = True Then
                If cmt.MouseStartX > cmt.MouseMoveX Then
                    tx = cmt.MouseStartX
                    Sx = cmt.MouseMoveX
                Else
                    Sx = cmt.MouseStartX
                    tx = cmt.MouseMoveX
                End If
                
                If cmt.MouseStartY > cmt.MouseMoveY Then
                    ty = cmt.MouseStartY
                    sY = cmt.MouseMoveY
                Else
                    sY = cmt.MouseStartY
                    ty = cmt.MouseMoveY
                End If
                
                If cmt.MouseDown = True Then
                    D3DDevice.SetTexture 0, ALine
                    KeyStrip(0) = CreateTLVertex(Sx, sY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(Sx + 1, sY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(Sx, ty, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(Sx + 1, ty, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                    
                    KeyStrip(0) = CreateTLVertex(tx - 1, sY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(tx, sY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(tx - 1, ty, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(tx, ty, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
       
       
                    KeyStrip(0) = CreateTLVertex(Sx, sY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(tx, sY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(Sx, sY + 1, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(tx, sY + 1, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                    
                    KeyStrip(0) = CreateTLVertex(Sx, ty - 1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                    KeyStrip(1) = CreateTLVertex(tx, ty - 1, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                    KeyStrip(2) = CreateTLVertex(Sx, ty, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                    KeyStrip(3) = CreateTLVertex(tx, ty, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
            End If
        End If
        
        For i = 32 To 1 Step -1
                For k = 0 To 7
                    cX = (i - ToOffset) * 32 - 2
                    cY = k * 64 + StartY
                    KeySizeX = 64
                    KeySizeY = 64
                    SelectedT = cX + 31 > Sx And cX + 31 < tx And cY + 31 > sY And cY + 31 < ty And cmt.MouseMove
                    
                    If k = 6 Then KeySizeY = 32
                    If k = 7 Then KeySizeY = 32: KeySizeX = 32: cY = cY - 32: cX = cX + 16
                    
                    If MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat - 1) * 8 + k) = True And MData((i + ToNowBeat + 1) * 8 + k) = True Then
                    
                        D3DDevice.SetTexture 0, KeySImage(k)
                        KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
                        KeyStrip(1) = CreateTLVertex(cX + KeySizeX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
                        KeyStrip(2) = CreateTLVertex(cX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
                        KeyStrip(3) = CreateTLVertex(cX + KeySizeX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        
                        If MData((i + ToNowBeat - 2) * 8 + k) = False Then
                            D3DDevice.SetTexture 0, KeySImage(k)
                            KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
                            KeyStrip(1) = CreateTLVertex(cX + (KeySizeX / 2), cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
                            KeyStrip(2) = CreateTLVertex(cX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
                            KeyStrip(3) = CreateTLVertex(cX + (KeySizeX / 2), cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        End If
                        
                        If MData((i + ToNowBeat + 2) * 8 + k) = False Then
                            D3DDevice.SetTexture 0, KeySImage(k)
                            KeyStrip(0) = CreateTLVertex(cX + (KeySizeX / 2), cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
                            KeyStrip(1) = CreateTLVertex(cX + KeySizeX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
                            KeyStrip(2) = CreateTLVertex(cX + (KeySizeX / 2), cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
                            KeyStrip(3) = CreateTLVertex(cX + KeySizeX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        End If
                        
                        If MData((i + ToNowBeat + 2) * 8 + k) = False Then
                            cX = (i + 1 - ToOffset) * 32 - 2
                            If k > 2 Then
                                o = k - 3
                            Else
                                o = k + 3
                            End If
                            
                            D3DDevice.SetTexture 0, KeyImage(o)
                            KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
                            KeyStrip(1) = CreateTLVertex(cX + KeySizeX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
                            KeyStrip(2) = CreateTLVertex(cX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
                            KeyStrip(3) = CreateTLVertex(cX + KeySizeX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        End If
                        
                        If SelectedT = True Then
                            SelectedKeys = SelectedKeys + 1
                            SelectedKey(SelectedKeys) = (ToNowBeat + i) * 8 + k
                        End If
                        
                    ElseIf GData((i + ToNowBeat) * 8 + k) = True And cX > -2.3 And MData((i + ToNowBeat - 2) * 8 + k) = False Then
                    
                        D3DDevice.SetTexture 0, KeyImage(k)
                        KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 0)
                        KeyStrip(1) = CreateTLVertex(cX + KeySizeX, cY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 0)
                        KeyStrip(2) = CreateTLVertex(cX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 0, 1)
                        KeyStrip(3) = CreateTLVertex(cX + KeySizeX, cY + KeySizeY, 0, 1, IIf(SelectedT, RGB(150, 150, 150), RGB(255, 255, 255)), 0, 1, 1)
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        
                        If SelectedT = True Then
                            SelectedKeys = SelectedKeys + 1
                            SelectedKey(SelectedKeys) = (ToNowBeat + i) * 8 + k
                        End If
                    End If
                            
                            If i + ToNowBeat > 0 And Mode <> "playing" And SData((i + ToNowBeat) * 8 + k) <> SData((i + ToNowBeat - 1) * 8 + k) Then
                                    TextRect.Left = IIf(k < 7, cX + 8, cX - 8)
                                    TextRect.Right = cX + 64
                                    TextRect.Top = 560 + k * 21
                                    
                                    Select Case k
                                        Case 0: WhichKey = IIf(Language = 0, "9鍵 X", "N9 X") + CStr(SData((i + ToNowBeat) * 8 + 0))
                                        Case 1: WhichKey = IIf(Language = 0, "6鍵 X", "N6 X") + CStr(SData((i + ToNowBeat) * 8 + 1))
                                        Case 2: WhichKey = IIf(Language = 0, "3鍵 X", "N3 X") + CStr(SData((i + ToNowBeat) * 8 + 2))
                                        Case 3: WhichKey = IIf(Language = 0, "7鍵 X", "N7 X") + CStr(SData((i + ToNowBeat) * 8 + 3))
                                        Case 4: WhichKey = IIf(Language = 0, "4鍵 X", "N4 X") + CStr(SData((i + ToNowBeat) * 8 + 4))
                                        Case 5: WhichKey = IIf(Language = 0, "1鍵 X", "N1 X") + CStr(SData((i + ToNowBeat) * 8 + 5))
                                        Case 6: WhichKey = IIf(Language = 0, "S鍵 X", "NS X") + CStr(SData((i + ToNowBeat) * 8 + 6))
                                        Case 7: WhichKey = "BPM " + CStr(BpmSet(SData((i + ToNowBeat) * 8 + 7)))
                                    End Select
                                    
                                    D3DX.DrawText MainFont, &HFFFFFFFF, WhichKey, TextRect, 0
                                    WhichKey = ""
                            End If
                    
                Next k

        Next i
        
                            
End Sub

Public Function CheckFirstCombo() As Long

Dim i As Long, j As Long

If Admin = False Then On Error Resume Next
        
        ReDim Preserve GData(TotalBeat * 8)
        For i = 0 To TotalBeat - 1
                For j = 0 To 7
                        If GData(i * 8 + j) = True Then CheckFirstCombo = i: Exit Function
                Next j
        Next i
        
End Function

Public Sub CheckStart()

Dim FirstCombo As Long, ToOffset As Single, ToNowBeat As Long, k As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat

If UseMode = "game" And ToNowBeat >= 2 Then
    For k = 0 To 5
        If (GameCheck((ToNowBeat - 2) * 8 + k) = 0 Or GameCheck((ToNowBeat - 2) * 8 + k) = 5) And GData((ToNowBeat - 2) * 8 + k) = True Then ShowMiss
    Next k
End If

FirstCombo = CheckFirstCombo

    If ToNowBeat >= FirstCombo - 64 And ToNowBeat < FirstCombo - 48 Then
        ShowBeatup 347, 128
    ElseIf ToNowBeat > FirstCombo - 48 And ToNowBeat < FirstCombo - 44 Then
        ShowBeatup 347, 128
    ElseIf ToNowBeat > FirstCombo - 44 And ToNowBeat < FirstCombo - 40 Then
        ShowBeatup 347, 128
    ElseIf ToNowBeat > FirstCombo - 40 And ToNowBeat < FirstCombo - 36 Then
        ShowBeatup 347, 128
    End If

    If ToNowBeat < FirstCombo - 32 Then
        RenderReady 0, 505, 364
    ElseIf ToNowBeat >= FirstCombo - 32 And ToNowBeat < FirstCombo - 16 Then
        RenderReady 1, 505, 364
    ElseIf ToNowBeat > FirstCombo - 16 And ToNowBeat < FirstCombo - 12 Then
        RenderReady 2, 505, 364
    ElseIf ToNowBeat > FirstCombo - 12 And ToNowBeat < FirstCombo - 8 Then
        RenderReady 3, 505, 364
    ElseIf ToNowBeat > FirstCombo - 8 And ToNowBeat < FirstCombo - 4 Then
        RenderReady 4, 505, 364
    End If

End Sub

Public Function ShowBeatup(StartX As Single, StartY As Single)

If Admin = False Then On Error Resume Next

                D3DDevice.SetTexture 0, ShowBeatupLogo(0)
                KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 256, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX, StartY + 256, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 256, StartY + 256, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                
                D3DDevice.SetTexture 0, ShowBeatupLogo(1)
                KeyStrip(0) = CreateTLVertex(StartX + 256 - 1, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 256 + 128, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX + 256 - 1, StartY + 256, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 256 + 128, StartY + 256, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Function ShowNormalBack(StartX As Single, StartY As Single)

If Admin = False Then On Error Resume Next

                D3DDevice.SetTexture 0, NormalBack
                KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 1024, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX, StartY + 621, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 1024, StartY + 621, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

                D3DDevice.SetTexture 0, NormalBackP
                KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 57, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX, StartY + 384, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 57, StartY + 384, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Function RenderReady(Number As Integer, StartX As Single, StartY As Single)

Dim cK As Long, cI As Long, cT As Long

If Admin = False Then On Error Resume Next

        If Number > 1 Then StartX = StartX - 96: cK = 3
        If Number = 4 Then cK = 3: cI = 32: cT = 16
        If Number = 1 And Mode = "playing" Then cma1.PlayReady
        If Number = 2 And Mode = "playing" Then cma1.PlayGo
        
                D3DDevice.SetTexture 0, READY(Number)
                KeyStrip(0) = CreateTLVertex(StartX + cT + 8 - 128 / (cK + 1), StartY + 8 - 32, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + cT + 8 + 128 / (cK + 1) - cI, StartY + 8 - 32, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX + cT + 8 - 128 / (cK + 1), StartY + 8 + 32, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + cT + 8 + 128 / (cK + 1) - cI, StartY + 8 + 32, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Function FindFinish(FArray() As Long) As Long

Dim i As Long, Number As Long

If Admin = False Then On Error Resume Next

ReDim FArray(0)
For i = 0 To TotalBeat - 1
    If GData(i * 8 + 7) = True Then
        ReDim Preserve FArray(Number)
        FArray(Number) = i
        Number = Number + 1
    End If
Next i

End Function

Public Sub RenderFinish()

Dim ToNowBeat As Long, ToOffset As Single, Combo As Long, cX As Single, FinsihCombo As Long, FCombo() As Long, i As Long

If Admin = False Then On Error Resume Next

Static StartBeat As Long, DoIt() As Boolean, Which As Integer

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 10 Then Exit Sub
CheckGData ToNowBeat

FindFinish FCombo
ReDim DoIt(UBound(FCombo))

    For i = 0 To UBound(FCombo)
        If UBound(FCombo) = 0 And FCombo(0) = 0 Then Exit Sub
        If ToNowBeat = FCombo(i) - 1 And DoIt(i) = False Then
            DoIt(i) = True
            StartBeat = ToNowBeat
        End If
    Next i
    
If ToNowBeat = StartBeat Then
    cX = ToOffset * 514
    ShowBS 2, cX, 236
ElseIf ToNowBeat = StartBeat + 1 Then
    ShowBS 2, 514, 236
ElseIf ToNowBeat = StartBeat + 3 Then
    ShowBS 2, 514, 236
ElseIf ToNowBeat = StartBeat + 5 Then
    ShowBS 2, 514, 236
ElseIf ToNowBeat = StartBeat + 7 Then
    cX = 514 + ToOffset * 514
    ShowBS 2, cX, 236
End If

End Sub

Public Function UpdateFps()

If Admin = False Then On Error Resume Next

    If GetTickCount() - FPS_LastCheck >= 1000 Then
        FPS_Current = FPS_Count
        FPS_Count = 0
        FPS_LastCheck = GetTickCount()
    End If
    FPS_Count = FPS_Count + 1

End Function

Public Sub ShowText()

Dim LData() As String, ToNowBeat As Long, ToOffset As Single, ShowWhichText As String

If Admin = False Then On Error Resume Next

If UseMode <> "game" Then
    cma6.CheckTime ToNowBeat, ToOffset
Else
    cma6.CheckTime ToNowBeat, ToOffset, False
End If

If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat

TextRect.Left = 0
TextRect.Right = 500

TextRect.Top = 1
D3DX.DrawText MainFont, &HFFFFFFFF, "FPS:" + CStr(FPS_Current), TextRect, 0

TextRect.Top = 18
D3DX.DrawText MainFont, &HFFFFFFFF, "Beat:" + CStr(ToNowBeat) + "/" + Right("000" + CStr(TotalBeat), 4), TextRect, 0

TextRect.Left = 100
TextRect.Top = 1
If Mode <> "playing" Then D3DX.DrawText MainFont, &HFFFFFFFF, "Offset:" + FormatNumber(ToOffset, 4), TextRect, 0

TextRect.Left = 0
TextRect.Top = 36
ShowWhichText = IIf(Language = 0, "時間:", "TIME:")
D3DX.DrawText MainFont, &HFFFFFFFF, ShowWhichText + cma2.MstoMin(cmt.Times.value) + "/" + cma2.MstoMin(SoundL), TextRect, 0

TextRect.Top = 54
ShowWhichText = IIf(Language = 0, "OFFSET 自動偵測:", "OFFSET DETECT:")
If CheckOffset > 1 Then CheckOffset = CheckOffset - 1
D3DX.DrawText MainFont, &HFFFFFFFF, ShowWhichText + CStr(FormatNumber(CheckOffset, 4)) + "/" + CStr(FormatNumber(1 + CheckOffset, 4)), TextRect, 0

TextRect.Top = 72
ShowWhichText = IIf(Language = 0, "剛按鍵的OFFSET:", "Key OFFSET:")
D3DX.DrawText MainFont, &HFFFFFFFF, ShowWhichText + CStr(PressOffset), TextRect, 0

TextRect.Top = 90
ShowWhichText = IIf(Language = 0, "Combo總數:", "ALlCOMBO:")
D3DX.DrawText MainFont, &HFFFFFFFF, ShowWhichText + CStr(cma2.CheckCombo(TotalBeat)), TextRect, 0


If UseMode = "game" Then
        TextRect.Top = 126
        D3DX.DrawText MainFont, &HFFFFFFFF, "Perfect:" + CStr(GameP(0)), TextRect, 0
        TextRect.Top = 144
        D3DX.DrawText MainFont, &HFFFFFFFF, "Great:" + CStr(GameP(1)), TextRect, 0
        TextRect.Top = 162
        D3DX.DrawText MainFont, &HFFFFFFFF, "Cool:" + CStr(GameP(2)), TextRect, 0
        TextRect.Top = 180
        D3DX.DrawText MainFont, &HFFFFFFFF, "Bad:" + CStr(GameP(3)), TextRect, 0
        TextRect.Top = 198
        D3DX.DrawText MainFont, &HFFFFFFFF, "Miss:" + CStr(GameP(4)), TextRect, 0
        TextRect.Top = 234
        D3DX.DrawText MainFont, &HFFFFFFFF, IIf(Language = 0, "分數:", "Score:") + CStr(Score), TextRect, 0
End If

TextRect.Top = 745
TextRect.Right = 1024
D3DX.DrawText MainFont, &HFFFFFFFF, " Music:" + Singer + " " + Melody + "(" + CStr(CLng(BpmSet(0))) + " bpm) (Time " + cma2.MstoMin(SoundL - cmt.Times.value) + ")", TextRect, 0

End Sub

Public Sub LoopRender()

If Admin = False Then On Error Resume Next

        Do
            If Mode <> "playing" And Room = False Then Render
            cma7.NeedToCheck
        DoEvents
        Loop Until ExitMsg = True
End Sub

Public Sub RenderBeginScene()

If Admin = False Then On Error Resume Next

        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        D3DDevice.BeginScene

End Sub

Public Sub RenderEndScene()

If Admin = False Then On Error Resume Next

        D3DDevice.EndScene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub ShowBackGround()

Dim cX As Single, cY As Single
        
If Admin = False Then On Error Resume Next

            D3DDevice.SetTexture 0, BackGround2(ChooseBackGround)
            KeyStrip(0) = CreateTLVertex(1, 1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
            KeyStrip(1) = CreateTLVertex(1024, 1, 0, 1, RGB(255, 255, 255), 0, 1, 0)
            KeyStrip(2) = CreateTLVertex(1, 768, 0, 1, RGB(255, 255, 255), 0, 0, 1)
            KeyStrip(3) = CreateTLVertex(1024, 768, 0, 1, RGB(255, 255, 255), 0, 1, 1)
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
        
End Sub


Public Sub RenderRightKey(StartY As Single, StartX As Single)

Dim i As Long, ToOffset As Single, cX As Single, cY As Single, k As Integer, ToNowBeat As Long, OffsetX As Long, Which As Long, One As Long, jj As Long, lcX As Single, fcX As Single, scX As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat + 31

If UseMode = "game" Then Which = -1
'
        For i = 30 To 0 Step -1
            If i + ToNowBeat < 1 Or ToNowBeat < 5 Then Exit For
                For k = 0 To 2
                
                        cX = StartX + (i - ToOffset) * (CSng(SData((ToNowBeat - 1) * 8 + k)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k)) - (CSng(SData((ToNowBeat) * 8 + k))))) * 9 + 2
                            cY = k * 64 + StartY
                            
                                If k = 1 Then
                                    OffsetX = 3
                                    cX = cX - OffsetX + 2
                                Else
                                    OffsetX = -3
                                    cX = cX + 2
                                End If
                            
                        If UseMode <> "game" Then
                                If MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat - 1) * 8 + k) = True And MData((i + ToNowBeat + 1) * 8 + k) = True Then
                                    
                                    For jj = 0 To 128
                                        If MData((i + ToNowBeat + jj) * 8 + k) = False Then
                                            One = jj
                                            Exit For
                                        End If
                                    Next jj
                                    lcX = StartX + (i - ToOffset + One - 1) * (CSng(SData((ToNowBeat - 1) * 8 + k)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k)) - (CSng(SData((ToNowBeat) * 8 + k))))) * 9 + 2
                                    For jj = 0 To 128
                                        If MData((i + ToNowBeat - jj) * 8 + k) = False Then
                                            One = jj
                                            Exit For
                                        End If
                                    Next jj
                                    fcX = StartX + (i - ToOffset - One + 1) * (CSng(SData((ToNowBeat - 1) * 8 + k)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k)) - (CSng(SData((ToNowBeat) * 8 + k))))) * 9 + 2
                                    
                                    scX = cX
                                    If scX < fcX + 48 Then scX = fcX + SData((ToNowBeat - 1) * 8 + k) * 9
                                    
                                    If cX < StartX - OffsetX + 1 Then cX = StartX - OffsetX + 1
                                    If fcX < StartX - OffsetX + 1 Then fcX = StartX - OffsetX + 1
                                    If lcX < StartX - OffsetX + 1 Then lcX = StartX - OffsetX + 1
                                    If scX < StartX - OffsetX + 1 Then scX = StartX - OffsetX + 1

                                    D3DDevice.SetTexture 0, KeySImage(k)
                                    KeyStrip(0) = CreateTLVertex(scX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(scX + SData((ToNowBeat - 1) * 8 + k) * 9, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(scX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(scX + SData((ToNowBeat - 1) * 8 + k) * 9, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

                                    If MData((i + ToNowBeat + 2) * 8 + k) = False Then
                                        If k = 1 Then
                                            lcX = lcX - OffsetX + 2
                                        Else
                                            lcX = lcX + 2
                                        End If
                                        D3DDevice.SetTexture 0, KeyImage(k + 3)
                                        KeyStrip(0) = CreateTLVertex(lcX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                        KeyStrip(1) = CreateTLVertex(lcX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                        KeyStrip(2) = CreateTLVertex(lcX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                        KeyStrip(3) = CreateTLVertex(lcX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                    End If
                        
                                ElseIf (GData((i + ToNowBeat) * 8 + k) = True And cX > StartX - OffsetX + 1 And cX < 1024 And MData((i + ToNowBeat - 2) * 8 + k) = False) Or (MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat + 1) * 8 + k) = False And cX > StartX - OffsetX + 1 And cX < 1024) Then
                                
                                        If MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat - 1) * 8 + k) = False Then
                                            D3DDevice.SetTexture 0, KeySImage(k)
                                            KeyStrip(0) = CreateTLVertex(cX + 48, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                            KeyStrip(1) = CreateTLVertex(cX + SData((ToNowBeat - 1) * 8 + k) * 9, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                            KeyStrip(2) = CreateTLVertex(cX + 48, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                            KeyStrip(3) = CreateTLVertex(cX + SData((ToNowBeat - 1) * 8 + k) * 9, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                        End If
                                    
                                    If MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat + 1) * 8 + k) = False Then
                                        D3DDevice.SetTexture 0, KeySImage(k)
                                        KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                        KeyStrip(1) = CreateTLVertex(cX + 16, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                        KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                        KeyStrip(3) = CreateTLVertex(cX + 16, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                    End If
                                
                                    If MData((i + ToNowBeat) * 8 + k) = True And MData((i + ToNowBeat + 1) * 8 + k) = False Then
                                        D3DDevice.SetTexture 0, KeyImage(k + 3)
                                    Else
                                        D3DDevice.SetTexture 0, KeyImage(k)
                                    End If
                                    
                                    KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                End If
                        Else
                            If GameCheck((i + ToNowBeat) * 8 + k) = 0 And GData((i + ToNowBeat) * 8 + k) = True Then
                                    D3DDevice.SetTexture 0, KeyImage(k)
                                    KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                            End If
                        End If
                Next
NextI:
        Next

End Sub

Public Sub RenderLeftKey(StartY As Single, LastX As Single)

Dim ToOffset As Single, cX As Single, cY As Single, k As Integer, i As Long, ToNowBeat As Long, OffsetX As Long, Which As Long, jj As Long, One As Long, scX As Single, lcX As Single, fcX As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat + 31
LastX = LastX + 4

If UseMode = "game" Then Which = -1

        For i = 30 To 0 Step -1
            If i + ToNowBeat < 1 Or ToNowBeat < 5 Then Exit For
                For k = 0 To 2

                            cX = LastX - (i - ToOffset) * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - CSng(SData((ToNowBeat) * 8 + k + 3)))) * 9
                            cY = k * 64 + StartY
                            
                                'If k = 1 Then cX = cX + 1
                            If UseMode <> "game" Then

                                If MData((i + ToNowBeat) * 8 + k + 3) = True And MData((i + ToNowBeat - 1) * 8 + k + 3) = True And MData((i + ToNowBeat + 1) * 8 + k + 3) = True Then
                                    
                                    For jj = 0 To 128
                                        If MData((i + ToNowBeat + jj) * 8 + k + 3) = False Then
                                            One = jj
                                            Exit For
                                        End If
                                    Next jj
                                    lcX = LastX - (i - ToOffset + One - 1) * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - CSng(SData((ToNowBeat) * 8 + k + 3)))) * 9 + 4.3
                                    For jj = 0 To 128
                                        If MData((i + ToNowBeat - jj) * 8 + k + 3) = False Then
                                            One = jj
                                            Exit For
                                        End If
                                    Next jj
                                    fcX = LastX - (i - ToOffset - One + 1) * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - ToOffset * (CSng(SData((ToNowBeat - 1) * 8 + k + 3)) - CSng(SData((ToNowBeat) * 8 + k + 3)))) * 9 + 4.3
                                    
                                    scX = cX
                                    If scX > fcX - SData((ToNowBeat - 1) * 8 + k + 3) * 9 Then scX = fcX - SData((ToNowBeat - 1) * 8 + k + 3) * 9
                                    
                                    If cX > LastX Then cX = LastX + SData((ToNowBeat - 1) * 8 + k + 3) * 9
                                    If fcX > LastX Then fcX = LastX + SData((ToNowBeat - 1) * 8 + k + 3) * 9
                                    If lcX > LastX Then lcX = LastX + SData((ToNowBeat - 1) * 8 + k + 3) * 9
                                    If scX > LastX Then scX = LastX + SData((ToNowBeat - 1) * 8 + k + 3) * 9

                                    D3DDevice.SetTexture 0, KeySImage(k + 3)
                                    KeyStrip(0) = CreateTLVertex(scX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(scX + SData((ToNowBeat - 1) * 8 + k + 3) * 9, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(scX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(scX + SData((ToNowBeat - 1) * 8 + k + 3) * 9, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

                                    If MData((i + ToNowBeat + 2) * 8 + k + 3) = False Then
                                        If k = 1 Then
                                            OffsetX = 8
                                            lcX = lcX + OffsetX - 3
                                        End If
                                        D3DDevice.SetTexture 0, KeyImage(k)
                                        KeyStrip(0) = CreateTLVertex(lcX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                        KeyStrip(1) = CreateTLVertex(lcX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                        KeyStrip(2) = CreateTLVertex(lcX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                        KeyStrip(3) = CreateTLVertex(lcX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                    End If
                        
                                ElseIf (GData((i + ToNowBeat) * 8 + k + 3) = True And cX < LastX And MData((i + ToNowBeat - 2) * 8 + k + 3) = False) Or (MData((i + ToNowBeat) * 8 + k + 3) = True And MData((i + ToNowBeat + 1) * 8 + k + 3) = False And cX < LastX) Then
                                
                                        If MData((i + ToNowBeat) * 8 + k + 3) = True And MData((i + ToNowBeat - 1) * 8 + k + 3) = False Then
                                            D3DDevice.SetTexture 0, KeySImage(k + 3)
                                            KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                            KeyStrip(1) = CreateTLVertex(cX + 16, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                            KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                            KeyStrip(3) = CreateTLVertex(cX + 16, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                        End If
                                    
                                    If MData((i + ToNowBeat) * 8 + k + 3) = True And MData((i + ToNowBeat + 1) * 8 + k + 3) = False Then
                                            D3DDevice.SetTexture 0, KeySImage(k + 3)
                                            KeyStrip(0) = CreateTLVertex(cX + 48, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                            KeyStrip(1) = CreateTLVertex(cX + SData((ToNowBeat - 1) * 8 + k + 3) * 9, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                            KeyStrip(2) = CreateTLVertex(cX + 48, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                            KeyStrip(3) = CreateTLVertex(cX + SData((ToNowBeat - 1) * 8 + k + 3) * 9, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                    End If
                                
                                    If MData((i + ToNowBeat) * 8 + k + 3) = True And MData((i + ToNowBeat + 1) * 8 + k + 3) = False Then
                                        D3DDevice.SetTexture 0, KeyImage(k)
                                    Else
                                        D3DDevice.SetTexture 0, KeyImage(k + 3)
                                    End If
                                If k = 1 Then cX = cX + 5
                                
                                    KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                End If


                            Else
                                If GameCheck((i + ToNowBeat) * 8 + k + 3) = 0 And GData((i + ToNowBeat) * 8 + k + 3) = True Then
                                    D3DDevice.SetTexture 0, KeyImage(k + 3)
                                    KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 64, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX, cY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 64, cY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                End If
                            End If
                Next
NextI:
        Next


End Sub

Public Sub RenderSpaceKey(StartY As Single, StartX As Single, LastX As Single)

Dim i As Long, ToOffset As Single, cX As Single, cY As Single, k As Integer, ToNowBeat As Long, j As Long, Number As Single, ShowNumber As Long ', ShowWhich As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

CheckGData ToNowBeat + 45

        For i = 1 To 44
                        If GData((i + ToNowBeat) * 8 + 6) = True Then
                        
                            For j = 0 To 1
                                Select Case j
                                    Case 0: Number = -1
                                    Case Else: Number = 1
                                End Select
                                cX = StartX + 129 + Number * (i - ToOffset) * CSng(SData(ToNowBeat * 8 + 6)) * 3
                                'cX = StartX + 129 + Number * (i - ToOffset) * ShowNumber * 3
                                
                                If cX > 360 And cX < 640 Then
                                    D3DDevice.SetTexture 0, KeyImageS
                                    KeyStrip(0) = CreateTLVertex(cX - 5, StartY - 5, 0, 1, RGB(240, 240, 240), 0, 0, 0)
                                    KeyStrip(1) = CreateTLVertex(cX + 32, StartY - 5, 0, 1, RGB(240, 240, 240), 0, 1, 0)
                                    KeyStrip(2) = CreateTLVertex(cX - 5, StartY + 27, 0, 1, RGB(240, 240, 240), 0, 0, 1)
                                    KeyStrip(3) = CreateTLVertex(cX + 32, StartY + 27, 0, 1, RGB(240, 240, 240), 0, 1, 1)
                                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                                End If
                            Next j
                        End If
        Next

End Sub

Public Function CheckGData(Number As Long)

Dim LNumber As Long, i As Long, u As Long

If Number * 8 > UBound(GData) Then
    ReDim Preserve GData(Number * 8)
    ReDim Preserve GameCheck(Number * 8)
End If

If Number * 8 > UBound(SData) Then
    LNumber = Fix(UBound(SData) / 8)
    ReDim Preserve SData(Number * 8)
    
    For i = LNumber To Fix(UBound(SData) / 8) - 1
        For u = 0 To 7
            SData(i * 8 + u) = SData((LNumber - 1) * 8 + u)
        Next u
    Next i
    
End If

End Function

Public Sub RenderCombo(StartY As Single, StartX As Single)

Dim CanBeat As Boolean, i As Integer, ToNowBeat As Long, ToOffset As Single, ComboNumber As Long

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat


                For i = 0 To 6
                        If ToNowBeat * 8 + 7 >= UBound(GData) Then ReDim Preserve GData((ToNowBeat + 1) * 8)
                    If GData(ToNowBeat * 8 + i) = True Then CanBeat = True: Exit For
                Next i
                
                Select Case CanBeat
                    Case True: RenderBig StartY, StartX
                    Case False: RenderSmall StartY, StartX
                End Select

End Sub

Public Sub RenderBig(StartY As Single, StartX As Single)

Dim WBeat As String, SNumber As Integer, i As Integer, NowShow As String, ChooseWhich As String, ComboNumber As Long, ToNowBeat As Long, ToOffset As Single, LoadR As Long, FirstCombo As Long, cX As Single, cToOffset As Single, cY As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat
ComboNumber = cma2.CheckCombo(ToNowBeat)
If UseMode = "game" Then ComboNumber = NowCombo
If ComboNumber = 0 Then GoTo EndRB

WBeat = CStr(ComboNumber): SNumber = Len(WBeat)

        If ComboNumber < 100 Then
            ChooseWhich = "C1"
        ElseIf ComboNumber < 200 Then
            ChooseWhich = "C2"
        Else
            ChooseWhich = "C3"
        End If

                cToOffset = ToOffset
        If ToOffset >= 0.5 Then cToOffset = 1 - ToOffset

        LoadR = GameR
        If UseMode <> "game" Then LoadR = 0
        

                StartX = 513
                StartY = 200
        
                If ToOffset >= 0.5 Then ToOffset = 1 - ToOffset

                D3DDevice.SetTexture 0, Perfect(LoadR)
                KeyStrip(0) = CreateTLVertex(StartX - 128 - ToOffset * 256, StartY - 64 - ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 128 + ToOffset * 256, StartY - 64 - ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX - 128 - ToOffset * 256, StartY + 64 + ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 128 + ToOffset * 256, StartY + 64 + ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                        
        cX = 500 - SNumber * (18 + cToOffset * 18)
        cY = 224
        
        For i = 1 To SNumber
            NowShow = ChooseWhich + Mid(WBeat, i, 1)
                D3DDevice.SetTexture 0, Combo(Val(Mid(NowShow, 2)))
            
                KeyStrip(0) = CreateTLVertex(cX, cY - cToOffset * 27.5, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(cX + 55 + cToOffset * 55, cY - cToOffset * 27.5, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(cX, cY + 55 + cToOffset * 27.5, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(cX + 55 + cToOffset * 55, cY + 55 + cToOffset * 27.5, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                
                cX = cX + 36 + cToOffset * 36
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
        
        Next i

Exit Sub
EndRB:

FirstCombo = CheckFirstCombo
If ToNowBeat < FirstCombo Then Exit Sub

    If UseMode = "game" Then
        LoadR = GameR
        

                StartX = 513
                StartY = 200
        
                If ToOffset >= 0.5 Then ToOffset = 1 - ToOffset

                        D3DDevice.SetTexture 0, Perfect(LoadR)
                        KeyStrip(0) = CreateTLVertex(StartX - 128 - ToOffset * 256, StartY - 64 - ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                        KeyStrip(1) = CreateTLVertex(StartX + 128 + ToOffset * 256, StartY - 64 - ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                        KeyStrip(2) = CreateTLVertex(StartX - 128 - ToOffset * 256, StartY + 64 + ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                        KeyStrip(3) = CreateTLVertex(StartX + 128 + ToOffset * 256, StartY + 64 + ToOffset * 128, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
    End If


End Sub

Public Sub CheckMoreCombo()

Dim Combo As Long, ToNowBeat As Long, ToOffset As Single, cX As Single

If Admin = False Then On Error Resume Next

Static StartBeat As Long, DoIt(1) As Boolean, Which As Integer

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 10 Then Exit Sub
CheckGData ToNowBeat
Combo = cma2.CheckCombo(ToNowBeat)
If UseMode = "game" Then Combo = NowCombo

If Combo = 100 And DoIt(0) <> True Then
    Which = 0
    DoIt(0) = True
    StartBeat = ToNowBeat
ElseIf Combo = 400 And DoIt(1) <> True Then
    Which = 1
    DoIt(1) = True
    StartBeat = ToNowBeat
End If

If ToNowBeat = StartBeat Then
    cX = 514 + (1 - ToOffset) * 500
    ShowBS Which, cX, 236
ElseIf ToNowBeat = StartBeat + 1 Then
    ShowBS Which, 514, 236
ElseIf ToNowBeat = StartBeat + 3 Then
    ShowBS Which, 514, 236
ElseIf ToNowBeat = StartBeat + 5 Then
    ShowBS Which, 514, 236
ElseIf ToNowBeat = StartBeat + 7 Then
    cX = 514 - ToOffset * 1000
    ShowBS Which, cX, 236
    DoIt(Which) = False
End If

End Sub

Public Sub CheckComboAdd()

Dim Combo As Long, ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat
Combo = cma2.CheckCombo(ToNowBeat)
If UseMode = "game" Then Combo = NowCombo

If Combo >= 100 Then ShowY 558, 260, 330, 594

If Combo >= 400 Then ShowB 558, 260, 254, 591

ShowSpaceBar 358, 622

    If Combo >= 150 Then
        ShowCB 0, 1, 390, 647
    ElseIf Combo >= 10 Then
        ShowCB 0, 0, 390, 647
    End If


    If Combo >= 200 Then
        ShowCB 1, 1, 390, 647
    ElseIf Combo >= 20 Then
        ShowCB 1, 0, 390, 647
    End If

    If Combo >= 250 Then
        ShowCB 2, 1, 390, 647
    ElseIf Combo >= 40 Then
        ShowCB 2, 0, 390, 647
    End If


    If Combo >= 300 Then
        ShowCB 3, 1, 390, 647
    ElseIf Combo >= 60 Then
        ShowCB 3, 0, 390, 647
    End If


    If Combo >= 350 Then
        ShowCB 4, 1, 390, 647
    ElseIf Combo >= 80 Then
        ShowCB 4, 0, 390, 647
    End If


    If Combo >= 400 Then
        ShowCB 5, 1, 390, 647
    ElseIf Combo >= 100 Then
        ShowCB 5, 0, 390, 647
    End If

End Sub

Public Function ShowBS(Number As Integer, StartX As Single, StartY As Single)

Dim ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

                D3DDevice.SetTexture 0, BYBUPY(Number)
                KeyStrip(0) = CreateTLVertex(StartX + 8 - 256, StartY + 8 - 64, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 8 + 256, StartY + 8 - 64, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX + 8 - 256, StartY + 8 + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 8 + 256, StartY + 8 + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Function ShowCB(Number As Integer, CNumber As Integer, StartX As Single, StartY As Single)

Dim cX As Single, ToNowBeat As Long, ToOffset As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0

cX = Number * 48

                D3DDevice.SetTexture 0, CYBACK(CNumber)
                KeyStrip(0) = CreateTLVertex(StartX + cX + 8 - 32, StartY + 8 - 16, 0, 1, RGB(225, 225, 225), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + cX + 8 + 32, StartY + 8 - 16, 0, 1, RGB(225, 225, 225), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX + cX + 8 - 32, StartY + 8 + 16, 0, 1, RGB(225, 225, 225), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + cX + 8 + 32, StartY + 8 + 16, 0, 1, RGB(225, 225, 225), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

                D3DDevice.SetTexture 0, CB(Number)
                KeyStrip(0) = CreateTLVertex(StartX + cX, StartY, 0, 1, RGB(225, 225, 225), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + cX + 16, StartY, 0, 1, RGB(225, 225, 225), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX + cX, StartY + 16, 0, 1, RGB(225, 225, 225), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + cX + 16, StartY + 16, 0, 1, RGB(225, 225, 225), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Sub RenderSmall(StartY As Single, StartX As Single)

Dim WBeat As String, SNumber As Integer, i As Integer, NowShow As String, ChooseWhich As String, ComboNumber As Long, ToNowBeat As Long, ToOffset As Single, LoadR As Long, FirstCombo As Long, cX As Single, cY As Single

If Admin = False Then On Error Resume Next

cma6.CheckTime ToNowBeat, ToOffset
If ToNowBeat < 1 Then ToNowBeat = 0
CheckGData ToNowBeat
ComboNumber = cma2.CheckCombo(ToNowBeat)
If UseMode = "game" Then ComboNumber = NowCombo

If ComboNumber = 0 Then GoTo EndRS

WBeat = CStr(ComboNumber): SNumber = Len(WBeat)

        If ComboNumber < 100 Then
            ChooseWhich = "C1"
        ElseIf ComboNumber < 200 Then
            ChooseWhich = "C2"
        Else
            ChooseWhich = "C3"
        End If

        cX = 500 - SNumber * 18
        cY = 224
        
        For i = 1 To SNumber
            NowShow = ChooseWhich + Mid(WBeat, i, 1)
                D3DDevice.SetTexture 0, Combo(Val(Mid(NowShow, 2)))
            
                KeyStrip(0) = CreateTLVertex(cX, cY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(cX + 55, cY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(cX, cY + 55, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(cX + 55, cY + 55, 0, 1, RGB(255, 255, 255), 0, 1, 1)

                cX = cX + 36
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
            
        Next i
        
        LoadR = GameR
        If UseMode <> "game" Then LoadR = 0
                
                StartX = 513
                StartY = 200
    
                D3DDevice.SetTexture 0, Perfect(LoadR)
                KeyStrip(0) = CreateTLVertex(StartX - 128, StartY - 64, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 128, StartY - 64, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX - 128, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 128, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

          
Exit Sub
EndRS:

FirstCombo = CheckFirstCombo
If ToNowBeat < FirstCombo Then Exit Sub

    If UseMode = "game" Then
        LoadR = GameR
        
                StartX = 513
                StartY = 200
    
                D3DDevice.SetTexture 0, Perfect(LoadR)
                KeyStrip(0) = CreateTLVertex(StartX - 128, StartY - 64, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 128, StartY - 64, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX - 128, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 128, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
    End If


End Sub

Public Function ShowB(StartX As Single, StartY As Single, BStartX As Single, BStartY As Single)

If Admin = False Then On Error Resume Next

                D3DDevice.SetTexture 0, Bup
                KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 64, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 64, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                
                D3DDevice.SetTexture 0, Bcup
                KeyStrip(0) = CreateTLVertex(BStartX, BStartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(BStartX + 512, BStartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(BStartX, BStartY + 128, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(BStartX + 512, BStartY + 128, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Function ShowY(StartX As Single, StartY As Single, BStartX As Single, BStartY As Single)

If Admin = False Then On Error Resume Next

                D3DDevice.SetTexture 0, Yup
                KeyStrip(0) = CreateTLVertex(StartX, StartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(StartX + 64, StartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(StartX, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(StartX + 64, StartY + 64, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))
                
                D3DDevice.SetTexture 0, Ycup
                KeyStrip(0) = CreateTLVertex(BStartX, BStartY, 0, 1, RGB(255, 255, 255), 0, 0, 0)
                KeyStrip(1) = CreateTLVertex(BStartX + 512, BStartY, 0, 1, RGB(255, 255, 255), 0, 1, 0)
                KeyStrip(2) = CreateTLVertex(BStartX, BStartY + 128, 0, 1, RGB(255, 255, 255), 0, 0, 1)
                KeyStrip(3) = CreateTLVertex(BStartX + 512, BStartY + 128, 0, 1, RGB(255, 255, 255), 0, 1, 1)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, KeyStrip(0), Len(KeyStrip(0))

End Function

Public Function GetKeyS(index As Integer) As Integer

If Admin = False Then On Error Resume Next

GetKeyS = SelectedKey(index)

End Function
