Attribute VB_Name = "cma7"
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFileByte Lib "wininet" Alias "InternetReadFile" (ByVal hFile As Long, sBuffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Dim Fso As New FileSystemObject

Public Sub ExitExe()

If Admin = False Then
    ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
Else
    End
End If

End Sub

Public Function GetLink(Url As String) As String

Dim EnCmd() As Byte

   EnCmd = JData(Url)

   For i = 0 To UBound(EnCmd)
        EnCmd(i) = EnCmd(i) Xor &H30
   Next

GetLink = StrConv(EnCmd, vbUnicode)

End Function

Public Function GetCode(Url As String) As String

Dim Tmp() As Byte, Tmp2() As Byte

Tmp = StrConv(Url, vbFromUnicode)
MCPU_PRoc.LargeByteToHex Tmp, Tmp2, 0

GetCode = StrConv(Tmp2, vbUnicode)

End Function

Public Function GetDCode(Url As String) As String

Dim Tmp() As Byte, Tmp2() As Byte

Tmp = StrConv(Url, vbFromUnicode)
MCPU_PRoc.LargeStrToBin Tmp, Tmp2, -1, 0

GetDCode = StrConv(Tmp2, vbUnicode)

End Function

Public Sub NeedToCheck()

'If Code1 <> (Code2 Xor 12341234) Then cma6.CloseAll

End Sub

Public Sub CheckUser()

'If GetData2(GetLink("584444400A1F1F4747471E445245401E5E55441F535459451F5352551E4058400F454355420D") + GetCode(User) + GetLink("165859540D") + GetCode(GetHardId)) <> GetPw(GetCode(User)) Then
    'MsgBox IIf(Language = 0, "該帳號未能使用", "This Account Not Allow To Use"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll
'End If
    cmt.Enabled = True
    Code2 = Code1 Xor 12341234


End Sub

Public Function GetPw(Which As String) As String

Dim o As Long, Tmp() As Byte, Tmp2() As Byte, Number As Long, Tmp3() As Byte

Tmp = StrConv(Which, vbFromUnicode)
MCPU_PRoc.LargeStrToBin Tmp, Tmp2, -1, 0

Number = UBound(Tmp2)
ReDim Tmp3(Number)

    For o = 0 To Number
        Tmp3(o) = Tmp2(o) Xor Tmp2(Number - o)
    Next o

MCPU_PRoc.LargeByteToHex Tmp3, Tmp, 0

GetPw = UCase(StrConv(Tmp, vbUnicode))

End Function

'Function GetData2(sUrl As String) As String
'Dim hOpen As Long, hFile As Long, sBuffer() As Byte, Ret As Long
'
'ReDim sBuffer(2000)
'
'hOpen = InternetOpen("Get Id Programme", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'hFile = InternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&)
'InternetReadFileByte hFile, sBuffer(0), 2000, Ret
'InternetCloseHandle hFile
'InternetCloseHandle hOpen
'If Ret > 2000 Then Ret = 2000
'If Ret > 0 Then
'    ReDim Preserve sBuffer(Ret - 1)
'    GetData2 = StrConv(sBuffer, vbUnicode)
'End If
'End Function

Function GetData2(sUrl As String) As String
Dim hOpen As Long, hFile As Long, sBuffer() As Byte, Ret As Long
Dim RSize As Long

hOpen = InternetOpen("Get Id Programme", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
hFile = InternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&)

Do
    DoEvents
    ReDim Preserve sBuffer(RSize + 200)
    InternetReadFileByte hFile, sBuffer(RSize), 200, Ret
    RSize = RSize + Ret
    
Loop While Ret > 0

InternetCloseHandle hFile
InternetCloseHandle hOpen
If RSize > 0 Then
    ReDim Preserve sBuffer(RSize - 1)
    GetData2 = StrConv(sBuffer, vbUnicode)
End If
End Function



Public Function GetData(sourceAddr) As String

Dim MyWebBrowser As New SHDocVw.InternetExplorer
Dim HTMLdoc As MSHTML.HTMLDocument

MyWebBrowser.Visible = False
MyWebBrowser.navigate (sourceAddr)

Do
Loop Until Not MyWebBrowser.Busy
Set HTMLdoc = MyWebBrowser.document

With HTMLdoc
    With .body
    
        GetData = .innerText
    End With
End With

Set HTMLdoc = Nothing
Set MyWebBrowser = Nothing

End Function

Function GetHardId() As String
Dim TmpA As String * 261
Dim TmpB As String * 261
Dim OutString As String
Dim OutStringB  As String
Dim CID As Long
GetVolumeInformation "C:\", TmpA, 261, CID, ByVal 0, ByVal 0, TmpB, 261
OutString = Right("00000000" + Hex(CID), 8)
For i = 0 To 2
OutStringB = OutStringB + Mid(OutString, i * 2 + 1, 2) + "-"
Next
OutStringB = OutStringB + Mid(OutString, 7, 2)
GetHardId = OutStringB
End Function


Public Sub LoadUser(LoadFile As String)

Dim NowFolder As String, SigT2 As String * 22, NewUser As String * 512

NowFolder = Replace(LoadFile, FindFileName(LoadFile), "")

    Decrypt_12 LoadFile, 0
    
    Open NowFolder + "\cma\" + "user.cbe" For Binary As #1
    Get #1, 1, SigT2
    If SigT2 = "CdiuBeatEditorFileUser" Then
        Get #1, , NewUser
        User = Trim(NewUser)
    End If
    Close #1
        cmt.User_Text.Caption = User
        User = cmt.User_Text.Caption
  
        DeleteDir NowFolder + "\cma\"

End Sub

Public Sub Saveuser(Saveuser As String)

Dim AByte() As Byte

AByte = StrConv(Saveuser, vbFromUnicode)
ReDim Preserve AByte(511)

    cma5.CreateDir TempPath + "\cma\"
    
    Open TempPath + "\cma\" + "user.cbe" For Binary As #1
    Put #1, 1, "CdiuBeatEditorFileUser"
    Put #1, , AByte
    Close #1
    
    
    cma5.CreateDir App.Path + "\Newuser\"
    Enrypt_12 TempPath + "cma\", "User", App.Path + "\Newuser\"

    DeleteDir TempPath + "\cma\"

End Sub

Public Sub LoadUserFile()

'If Fso.FileExists(App.Path + "\User.cdiu") = False Then MsgBox IIf(Language = 0, "該帳號未能使用", "This Account Not Allow To Use"), vbYes, IIf(Language = 0, "系統訊息", "System Info"): cma6.CloseAll

End Sub

Public Sub CheckAdmin()

Dim EnCmd() As Byte

   EnCmd = JData("1F51545D595E0A59515D4359454959")

   For i = 0 To UBound(EnCmd)
        EnCmd(i) = EnCmd(i) Xor &H30
   Next

If InStr(1, Command, StrConv(EnCmd, vbUnicode)) > 0 Then
    cmt.HighAdmin.Visible = True
    Admin = True
End If

End Sub

Public Function JData(DataA As String)
Dim y() As Byte, LenBC As Long, Data As String
Data = UCase(DataA)
LenBC = Len(Data) / 2 - 1
ReDim y(LenBC)
For i = 0 To LenBC
y(i) = "&H" + Mid(Data, i * 2 + 1, 2)
Next
JData = y()
End Function


Public Sub CheckSlk(LoadFile As String)

Dim YNumber As Integer, XNumber As Integer, WData() As String, ErrorS As String, PNumber As Long, i As Long, o As Long, u As Long, TData() As String, Number As Long, Finish As Long, count As Long, AData() As Integer

If Admin = False Then On Error Resume Next

cma3.LoadSlk LoadFile, YNumber, XNumber, WData
        
ReDim AData(0): AData(0) = 0

ErrorS = ErrorS + "    正在讀取列表內容..." + vbCrLf + vbCrLf

        If YNumber < 2 Then
            ErrorS = ErrorS + "    頭2行不能有空白 否則不能繼續檢查" + vbCrLf
            PNumber = PNumber + 1
        End If

        For i = 3 To YNumber
                
                If (WData(5, i) = "f") Or (WData(5, i) = "s,f") Or (WData(5, i) = "f,s") Or (WData(5, i) = "n,f") Or (WData(5, i) = "f,n") Or (WData(5, i) = "s,n,f") Or (WData(5, i) = "s,f,n") Or (WData(5, i) = "n,s,f") Or (WData(5, i) = "n,f,s") Or (WData(5, i) = "f,n,s") Or (WData(5, i) = "f,s,n") Then Finish = Finish + 1
                
                If (WData(5, i) = "f") Then
                    ErrorS = ErrorS + "    第" + CStr(i + 1) + "行設定為 Finish 動作 但沒有Note 或 Space" + vbCrLf
                    PNumber = PNumber + 1
                End If

                If (WData(4, i) <> "") And (WData(5, i) = "") And (WData(6, i) <> "") Then
                    ErrorS = ErrorS + "    第" + CStr(i) + "行沒有設定Type" + vbCrLf
                    PNumber = PNumber + 1
                End If
                
                If (WData(4, i) = "") And (WData(5, i) <> "") And (WData(6, i) <> "") Then
                    ErrorS = ErrorS + "    第" + CStr(i) + "行沒有設定Beat" + vbCrLf
                    PNumber = PNumber + 1
                End If
                
                If (i > 8) Then
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 6)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 5)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 4)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 3)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 2)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i - 1)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                End If
                
                If (i < YNumber - 6) Then
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 1)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 2)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 3)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 4)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 5)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                        
                        If (WData(5, i) = "f") And (WData(4, i) = WData(4, i + 6)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行設定為 Finish 動作時 如果需要有Note或space 就不可以分開 一定要 n,f 或 s,f 或 f" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                End If
                
                Select Case WData(6, i)
                    Case 1:
                    Case "1":
                    Case 3:
                    Case "3":
                    Case 4:
                    Case "4":
                    Case 6:
                    Case "6":
                    Case 7:
                    Case "7":
                    Case 9:
                    Case "9":
                    Case Else
                            If (WData(6, i) <> "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行的Note 設定為 ( " + WData(6, i) + " ) 可能有問題" + vbCrLf
                                PNumber = PNumber + 1
                            End If
                            
                            If (WData(5, i) = "n") And (WData(6, i) = "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行還沒有設定Note" + vbCrLf
                                PNumber = PNumber + 1
                            End If
                            
                            If (WData(5, i) = "n,f") And (WData(6, i) = "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行還沒有設定Note" + vbCrLf
                                PNumber = PNumber + 1
                            End If
                            
                            If (WData(5, i) = "f,n") And (WData(6, i) = "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行還沒有設定Note" + vbCrLf
                                PNumber = PNumber + 1
                            End If
                
                            If (WData(5, i) = "s,n") And (WData(6, i) = "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行還沒有設定Note" + vbCrLf
                                PNumber = PNumber + 1
                            End If
                            
                            If (WData(5, i) = "n,s") And (WData(6, i) = "") Then
                                ErrorS = ErrorS + "    第" + CStr(i) + "行還沒有設定Note" + vbCrLf
                                PNumber = PNumber + 1
                            End If

                End Select


                For o = 0 To UBound(AData)
                        If i = AData(o) Then GoTo NoCheckB
                Next o


                For u = 1 To YNumber
                        If u = i Then GoTo NoCheckA
                
                        If (WData(4, i) = WData(4, u)) Then
                            AddBackArray AData, u
                            ErrorS = ErrorS + "    第" + CStr(i) + "行和第" + CStr(u) + "行 是同1個Beat 請檢查是否雙鍵問題" + vbCrLf
                            PNumber = PNumber + 1
                        End If

NoCheckA:
                Next u
NoCheckB:

                If (WData(5, i) = "s") And (WData(6, i) <> "") Then
                    ErrorS = ErrorS + "    第" + CStr(i) + "行的Type是S 但有設定Note 可能有問題" + vbCrLf
                    PNumber = PNumber + 1
                End If

                If (i < YNumber) Then
                        If CInt(WData(4, i + 1)) < CInt(WData(4, i)) Then
                            ErrorS = ErrorS + "    第" + CStr(i) + "行的Beat數值比第" + CStr(i + 1) + "行的Beat數值大 可能有問題" + vbCrLf
                            PNumber = PNumber + 1
                        End If
                End If

                If ((i = YNumber) And (WData(5, i) = "s")) Then
                    ErrorS = ErrorS + "    最後一行設定為空白鍵 可能是錯誤" + vbCrLf
                    PNumber = PNumber + 1
                End If

Next i

        If (Finish = 0) Then
            ErrorS = ErrorS + "    此文件 沒有設定Finish 會有機會不能使用" + vbCrLf
            PNumber = PNumber + 1
        End If

        Open LoadFile For Input As #1
            Do
            ReDim Preserve TData(Number)
            Line Input #1, TData(Number)
            If TData(Number) = "E" Then count = count + 1
            Number = Number + 1
            Loop Until EOF(1)
        Close #1

        If (count > 1) Then
            ErrorS = ErrorS + "    這文件可能曾經存取錯誤過 有機會出面 爆箭頭情況" + vbCrLf
            PNumber = PNumber + 1
        End If

ErrorS = ErrorS + "    請檢查文件名的大小寫 以免文件名影響有問題..." + vbCrLf


ErrorS = ErrorS + vbCrLf + IIf(PNumber = 0, "    此文件 找不到任何問題", "    此文件共找到" + CStr(PNumber) + "個問題") + vbCrLf

systemRead.ShowText.Text = ErrorS
systemRead.Show

End Sub

