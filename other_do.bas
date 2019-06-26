Attribute VB_Name = "other_do"
Sub DeleteFile(file As String)
On Error GoTo EndDelete
Kill file
EndDelete:
End Sub

Public Function StrToBin(data As String, LStrToBinTime As Long)
Dim OData() As Byte, i As Long
Dim IData() As Byte
Dim Tmp As Integer, ByteUBound As Long
IData = data
ByteUBound = (UBound(IData) / 4) - 1
ReDim OData(ByteUBound)
For i = 0 To ByteUBound
    OData(i) = CByte(CharToBin(IData(i * 4)) * &H10 + CharToBin(IData(i * 4 + 2)))
Next
StrToBin = OData
End Function

Function CharToBin(CharByte As Byte) As Integer
        If CharByte >= &H30 And CharByte <= &H39 Then
            CharToBin = CharByte - &H30
        ElseIf CharByte >= &H41 And CharByte <= &H46 Then
            CharToBin = CharByte - &H41 + 10
        ElseIf CharByte >= &H61 And CharByte <= &H66 Then
            CharToBin = CharByte - &H61 + 10
        End If
End Function

Public Function StrToHex(AData As String, StrToHexTime)
Dim StrData As String, i As Long, AByte() As Byte
StrToHexTime = timeGetTime()
AByte = StrConv(AData, vbFromUnicode)

'把轉換的16進制 轉換成10進制的字串
For i = 0 To UBound(AByte)
    StrData = StrData + Hex(AByte(i))
Next
StrToHexTime = timeGetTime - StrToHexTime
'傳回值
StrToHex = StrData
End Function


Public Function ReserveStr(value As String) As String

Dim i As Long, PValue() As Byte
Dim TLen As Long, Tmp(1) As Byte
PValue = value
TLen = UBound(PValue) / 2

For i = 0 To TLen Step 2
    Tmp(o) = PValue(i)
    Tmp(1) = PValue(i + 1)
     PValue(i) = PValue(UBound(PValue) - i - 1)
     PValue(i + 1) = PValue(UBound(PValue) - i)
     PValue(UBound(PValue) - i - 1) = Tmp(o)
     PValue(UBound(PValue) - i) = Tmp(1)
Next
ReserveStr = PValue
End Function
