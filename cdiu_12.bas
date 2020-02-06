Attribute VB_Name = "cdiu_12"
Option Explicit

Dim Fso As New FileSystemObject
Dim objCompress As New clsCryptoAPIandCompression

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Type Cdiu_Header
     Sign As String * 22
     NumberOfFile As Long
     FooterPos As Long
     FileName As String * 60
End Type

Type FileBlog
     FileName As String * 512
     OSize As Long
     DataSize As Long
     DataAddr As Long
End Type

Public Function FindFile(Data() As String, TFolder As Object, Path As String)
Dim Folder As Object
Dim file As Object

    For Each Folder In TFolder.SubFolders
        FindFile Data, Folder, Path + Folder.Name + "\"
    Next

    For Each file In TFolder.Files
        Data(UBound(Data)) = Path + file.Name
        ReDim Preserve Data(UBound(Data) + 1)
    Next
    
End Function


Public Function Enrypt_12(Path As String, NewFile As String, SaveCdiuPath As String)


Dim TTime As Long, LoadFileTime As Long, SaveTime As Long, GetFileTime As Long, SaveFileTime As Long
Dim ZlibTime As Long, XorTime As Long, StrToHexTime As Long, StrToHexTimeB As Long
Dim HeaderA As Cdiu_Header, FooterB() As FileBlog
Dim FileByteOfName() As String, RData() As Byte, FileHex() As String
Dim PathA As String, SaveFileName As String
Dim i As Long, FileNumber As Integer
Dim Folder As Object

PathA = Mid(Path, 1, InStrRev(Path, "\"))

Set Folder = Fso.GetFolder(Path)
ReDim FileByteOfName(0)
FindFile FileByteOfName, Folder, ""
Set Folder = Nothing
FileNumber = UBound(FileByteOfName)
ReDim Preserve FileByteOfName(UBound(FileByteOfName) - 1)
ReDim FooterB(FileNumber - 1)
ReDim FileHex(FileNumber - 1)

SaveFileName = NewFile + ".cdiu"
SaveFileName = Replace(SaveFileName, ".cdiu.cdiu", ".cdiu")

HeaderA.Sign = "Cdiu_Encrypt_File_1.2"
HeaderA.NumberOfFile = FileNumber
HeaderA.FileName = StrToHex(Dir(Path, vbDirectory), StrToHexTimeB)

DeleteFile PathA + SaveFileName

Open SaveCdiuPath + SaveFileName For Binary Access Write As #3
Put #3, 1, HeaderA

For i = 0 To FileNumber - 1
    FileHex(i) = StrToHex(FileByteOfName(i), StrToHexTime)
    FooterB(i).FileName = FileHex(i)

    Open Path + "\" + FileByteOfName(i) For Binary As #1
    If LOF(1) = 0 Then
        ReDim RData(0)
    Else
        ReDim RData(LOF(1) - 1)
    End If
    FooterB(i).OSize = UBound(RData) + 1
    Get #1, 1, RData
    Close #1
    LoadFileTime = timeGetTime() - LoadFileTime
    
    ZlibTime = timeGetTime()
    objCompress.CompressByteArray RData, 9
    ZlibTime = timeGetTime() - ZlibTime
    
    XorTime = timeGetTime()
    objCompress.EncryptDecryptB VarPtr(RData(0)), UBound(RData) + 1, "it_is_done_by_cdiu_hahaha", True
    XorTime = timeGetTime() - XorTime

    FooterB(i).DataSize = UBound(RData) + 1
    FooterB(i).DataAddr = Loc(3)

    SaveTime = timeGetTime()
    Put #3, , RData
    SaveTime = timeGetTime() - SaveTime
Next

ReDim RData(0)
HeaderA.FooterPos = Loc(3)
Put #3, , FooterB
Put #3, 1, HeaderA
Close #3

End Function


Public Function Enrypt_12File(Path As String)

Dim TTime As Long, LoadFileTime As Long, SaveTime As Long, GetFileTime As Long, SaveFileTime As Long
Dim ZlibTime As Long, XorTime As Long, StrToHexTime As Long, StrToHexTimeB As Long
Dim HeaderA As Cdiu_Header, FooterB As FileBlog
Dim FileByteOfName As String, RData() As Byte, FileHex As String
Dim PathA As String, SaveFileName As String
Dim i As Long, FileNumber As Integer
Dim Folder As Object

TTime = timeGetTime()
GetFileTime = timeGetTime()
PathA = Mid(Path, 1, InStrRev(Path, "\"))

FileNumber = 1

GetFileTime = timeGetTime() - GetFileTime
SaveFileTime = timeGetTime()

HeaderA.Sign = "Cdiu_Encrypt_File_1.2"
HeaderA.NumberOfFile = FileNumber
HeaderA.FileName = StrToHex(SaveFileName, StrToHexTimeB)

DeleteFile PathA + SaveFileName

Open PathA + SaveFileName For Binary Access Write As #3
Put #3, 1, HeaderA

    FileHex = StrToHex(FileByteOfName, StrToHexTime)
    FooterB.FileName = FileHex
    LoadFileTime = timeGetTime()

    Open Path For Binary As #1
    If LOF(1) = 0 Then
        ReDim RData(0)
    Else
        ReDim RData(LOF(1) - 1)
    End If
    FooterB.OSize = UBound(RData) + 1
    Get #1, 1, RData
    Close #1
    LoadFileTime = timeGetTime() - LoadFileTime
    
    ZlibTime = timeGetTime()
    objCompress.CompressByteArray RData, 9
    ZlibTime = timeGetTime() - ZlibTime
    
    XorTime = timeGetTime()
    objCompress.EncryptDecryptB VarPtr(RData(0)), UBound(RData) + 1, "it_is_done_by_cdiu_hahaha", True
    XorTime = timeGetTime() - XorTime

    FooterB.DataSize = UBound(RData) + 1
    FooterB.DataAddr = Loc(3)

    SaveTime = timeGetTime()
    Put #3, , RData
    SaveTime = timeGetTime() - SaveTime

ReDim RData(0)
HeaderA.FooterPos = Loc(3)
Put #3, , FooterB
Put #3, 1, HeaderA
Close #3

SaveFileTime = timeGetTime() - SaveFileTime
TTime = timeGetTime() - TTime

End Function



Public Function Decrypt_12(DcodeFileName As String, TTime As Long, Optional SavePath As String, Optional SaveInTmp As Boolean)

Dim StrToHexTime As Long, XorTime As Long, StrToHexTimeB As Long, ZlibTime As Long, StrToHexTimeC As Long
Dim LoadFileTime As Long, SaveFileTime As Long, SaveTime As Long
Dim HeaderA As Cdiu_Header, FooterB() As FileBlog
Dim SaveFileName As String, NewFolder As String
Dim FileByteOfName() As String, RData() As Byte
Dim i As Long

Open DcodeFileName For Binary Access Read As #3
    Get #3, 1, HeaderA
    ReDim FooterB(HeaderA.NumberOfFile - 1)
    Get #3, HeaderA.FooterPos + 1, FooterB

i = InStrRev(DcodeFileName, "\")
If i > 0 Then
    SaveFileName = Mid(DcodeFileName, i + 1)
Else
    SaveFileName = DcodeFileName
End If

If HeaderA.NumberOfFile = 1 Then
    NewFolder = Mid(DcodeFileName, 1, InStrRev(DcodeFileName, "\"))
Else
    NewFolder = Mid(DcodeFileName, 1, InStrRev(DcodeFileName, "\")) + StrConv(StrToBin(Trim(HeaderA.FileName), StrToHexTime), vbUnicode) + "\"
    If Not (Fso.FolderExists(NewFolder)) And SaveInTmp = False Then Fso.CreateFolder NewFolder
End If

ReDim FileByteOfName(HeaderA.NumberOfFile - 1)


For i = 0 To HeaderA.NumberOfFile - 1
    FileByteOfName(i) = StrConv(StrToBin(Trim(FooterB(i).FileName), StrToHexTime), vbUnicode)
    ReDim RData(FooterB(i).DataSize - 1)
    Get #3, FooterB(i).DataAddr + 1, RData

    XorTime = timeGetTime()
    objCompress.EncryptDecryptB VarPtr(RData(0)), UBound(RData) + 1, "it_is_done_by_cdiu_hahaha", False
    XorTime = timeGetTime() - XorTime

    ZlibTime = timeGetTime()
    objCompress.DecompressByteArray RData, FooterB(i).OSize
    ZlibTime = timeGetTime() - ZlibTime

    SaveTime = timeGetTime()
    DeleteFile NewFolder + FileByteOfName(i)
    If SaveInTmp = False Then
        GenFolder NewFolder + FileByteOfName(i)
    
        If SavePath = "" Then cma5.CreateDir NewFolder + "cma\"
        
        If SavePath = "" Then SavePath = NewFolder + "cma"
        
        Open SavePath + "\" + FileByteOfName(i) For Binary As #2
        Put #2, 1, RData
        Close #2
        cma4.CheckFileIn FileByteOfName(i), RData
    Else
        If Fso.FileExists(SavePath + "\" + FileByteOfName(i)) Then
            Open SavePath + "\" + FileByteOfName(i) For Binary As #4
            ReDim RData(FileLen(SavePath + "\" + FileByteOfName(i)) - 1)
            Get #4, , RData
            Close #4
        End If
        cma4.CheckFileIn FileByteOfName(i), RData
    End If
Next i
Close #3

End Function

Public Function GenFolder(file As String)
    Dim Pos As Long
    Pos = 0
    Do
        Pos = InStr(Pos + 1, file, "\")
        If Pos > 0 Then
            If Not (Fso.FolderExists(Mid(file, 1, Pos - 1))) Then Fso.CreateFolder Mid(file, 1, Pos - 1)
        End If
    Loop While Pos > 0
End Function
