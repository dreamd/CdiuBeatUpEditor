VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCryptoAPIandCompression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'the following are for encryption/decryption
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptEncryptB Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecryptB Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, ByRef pdwDataLen As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByVal dwFlags As Long) As Long

'the following are for compression/decompression
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'the following are for encryption/decryption
Private Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0" & vbNullChar
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const USE_THIS_CSP As String = MS_ENHANCED_PROV 'can set this to other providers.  See Microsoft's website for full listing
Private Const KEY_CONTAINER As String = "CryptoAPIExample" & vbNullChar 'this is usually set to your program's name
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_DELETEKEYSET As Long = 16
Private Const CRYPT_CREATE_SALT As Long = 4
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const KP_SALT As Long = 2
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_SID_SHA As Long = 4
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const CALG_SHA As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Const AT_KEYEXCHANGE As Long = 1
Private Const AT_SIGNATURE As Long = 2
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB = 1
Private Const PUBLICKEYBLOB As Long = 6
Private Const PRIVATEKEYBLOB As Long = 7
Private Const CRYPT_NO_SALT = 16
Private Const NTE_BAD_SIGNATURE As Long = -2146893818
Private Const KEY_LENGTH_512 = &H2000000 'upper half = "200" hex = 512 decimal
Private Const KEY_LENGTH_1024 = &H4000000 'upper half = "400" hex = 1024 decimal
Private Const KEY_LENGTH_2048 = &H8000000 'etc....
Private Const KEY_LENGTH_4096 = &H10000000
Private Const KEY_LENGTH_8192 = &H20000000
Private Const KEY_LENGTH_16384 = &H40000000

'the following are for encryption/decryption
Dim hCryptProv As Long 'the handle to the CSP
Dim strSALT As String 'the session key SALT
Dim strHash As String 'the value of the Hash
Dim hSessionKey As Long 'the handle to the current session key
Dim hKeyPair As Long 'the handle to the current key pair
Dim strPublicPrivateBlob As String 'the value of the private key in BLOB format.  Note that the public key is also put in here by the CryptoAPI
Dim strPublicBlob As String 'the value of the public key in BLOB format.  This is what you can send to other people.
Dim strSessionBlob As String 'the encrypted session key used during key pair encryption/decryption
Dim lngType As Long 'type of key in use (Export or Signature)
Dim strSig As String 'the value of the signature
Dim lngTheKeyLength As Long 'key length

'the following are for compression/decompression
Dim lngCompressedSize As Long
Dim lngDecompressedSize As Long

Enum CZErrors 'for compression/decompression
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Enum CompressionLevels 'for compression/decompression
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    'note that levels 2-8 exist, too
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = -1
End Enum

Public Sub SessionStart()
'this could be placed at the beginning of EncryptDecrypt, but if you are doing
'multiple encryptions/decryptions, calling this once speeds things up
Dim lngReturnValue As Long

'Get handle to CSP
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, USE_THIS_CSP, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
If lngReturnValue = 0 Then
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, USE_THIS_CSP, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above.  Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists."
End If

End Sub

Public Sub SessionEnd()

'Release any session key handle
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

'Release any key pair handle
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair

'Release provider handle
If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0

End Sub


Public Function EncryptDecryptB(ByVal Text As Long, DataSize As Long, Key As String, Encrypt As Boolean)
'the code in this function encrypts/decrypts the data using a single key
Dim lngLength As Long
Dim lngSALTLen As Long
Dim lngReturnValue As Long

SessionKeyFromPassword Key 'get a session key derived from the password


strSALT = "1234143" + vbNullChar
lngReturnValue = CryptSetKeyParam(hSessionKey, KP_SALT, strSALT, 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not set the SALT."

'Encrypt or decrypt depending on the Encrypt parameter
lngLength = DataSize
If Encrypt Then
    lngReturnValue = CryptEncryptB(hSessionKey, 0, 1, 0, Text, lngLength, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Error during CryptEncrypt."
Else
    lngReturnValue = CryptDecryptB(hSessionKey, 0, 1, 0, Text, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Error during CryptDecrypt."
End If

'Destroy the session key
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

End Function


Private Sub SessionKeyFromPassword(ByVal Key As String)
'This sub takes a string key as input and sets the module-level
'hSessionKey variable to a new session key handle.
'This sub is used by EncryptDecrypt, Export_KeyPair and Import_KeyPair.
Dim lngParams As Long
Dim lngReturnValue As Long
Dim strHash As String
Dim lngHashLen As Long
Dim hHash As Long 'the handle to the hash object

'Create a hash object to calculate a session
'key from the Password (instead of encrypting
'with the actual key)
lngReturnValue = CryptCreateHash(hCryptProv, CALG_SHA, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not create a Hash Object (CryptCreateHash API)"
'can use CALG_MD5 to get a 128-bit hash.  CALG_SHA returns a 160-bit hash (more secure).

'Hash the Password
lngReturnValue = CryptHashData(hHash, Key, Len(Key), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not calculate a Hash Value (CryptHashData API)"
'A hash is a 'fingerprint' of any string.
'Hashes are extremely useful for determining whether a
'transmission or file has been altered.  This code can use
'one of two algorithms (see note above).  No matter what the
'length of input data, the hash will be a fixed length and
'will be unique for that string of data.  The same hash is produced for
'the same input data every time.  This is useful here to
'produce a fixed-length, unique password for any length password entered.

'Get the actual hash value
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0) 'get the hash length
strHash = String(lngHashLen + 1, vbNullChar)
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0) 'get the hash value
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not lngReturnValuerieve the hash value"

'Set certain values to add more flexibility and security.
'Make the key exportable. (I don't export the key in this sample code)
lngParams = CRYPT_EXPORTABLE 'use this when you generate your own SALT, which is recommended (see 8 lines below)
'Make the key exportable and add a system-generated SALT.

'release old session key handle if one exists
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey

'Derive a session key from the hash object
lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, hSessionKey)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not create a session key (CryptDeriveKey API)"

'Destroy the hash object
If hHash <> 0 Then CryptDestroyHash hHash

End Sub


Public Function CompressByteArray(TheData() As Byte, CompressionLevel As Integer) As Long
'compress a byte array
Dim lngResult As Long
Dim lngBufferSize As Long
Dim arrByteArray() As Byte

lngDecompressedSize = UBound(TheData) + 1



'Allocate memory for byte array
lngBufferSize = UBound(TheData) + 1
lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
ReDim arrByteArray(lngBufferSize)

'Compress byte array (data)
lngResult = compress2(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)

'Truncate to compressed size
ReDim Preserve TheData(lngBufferSize - 1)
CopyMemory TheData(0), arrByteArray(0), lngBufferSize

'Set property
lngCompressedSize = UBound(TheData) + 1

'return error code (if any)
CompressByteArray = lngResult

End Function

Public Function DecompressByteArray(TheData() As Byte, OriginalSize As Long) As Long
'decompress a byte array
Dim lngResult As Long
Dim lngBufferSize As Long
Dim arrByteArray() As Byte

lngDecompressedSize = OriginalSize
lngCompressedSize = UBound(TheData) + 1

'Allocate memory for byte array
lngBufferSize = OriginalSize
lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
ReDim arrByteArray(lngBufferSize)

'Decompress data
lngResult = uncompress(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1)

'Truncate buffer to compressed size
ReDim Preserve TheData(lngBufferSize - 1)
CopyMemory TheData(0), arrByteArray(0), lngBufferSize

'return error code (if any)
DecompressByteArray = lngResult

End Function

Private Sub Class_Initialize()
    SessionStart
End Sub

Private Sub Class_Terminate()
    SessionEnd
End Sub

Public Function CompressFile11(FilePathIn() As Byte, FilePathOut As String, CompressionLevel As Integer, fullen As Long, ZlibTime As Long, FileSize As Long) As Long
Dim lngResult As Long, lngFileLen As Long

'計算總大小來加壓
lngFileLen = UBound(FilePathIn) + 1

'加壓過程
ZlibTime = timeGetTime()
lngResult = CompressByteArray11(FilePathIn(), CompressionLevel)
ZlibTime = timeGetTime() - ZlibTime

'傳回資料
FileSize = lngFileLen
fullen = UBound(FilePathIn) + 1
CompressFile11 = lngResult

End Function

Public Function DecompressFile11(FilePathIn, TheBytes() As Byte, FileSize As Long, dcodetime As Long) As Long
Dim lngResult As Long

'解壓內容
dcodetime = timeGetTime()
lngResult = DecompressByteArray11(TheBytes(), FileSize)
dcodetime = timeGetTime() - dcodetime

'傳回正常
DecompressFile11 = lngResult

End Function

Public Function CompressByteArray11(TheData() As Byte, CompressionLevel As Integer) As Long
'compress a byte array
Dim lngResult As Long
Dim lngBufferSize As Long
Dim arrByteArray() As Byte

lngDecompressedSize = UBound(TheData) + 1

'Allocate memory for byte array
lngBufferSize = UBound(TheData) + 1
lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
ReDim arrByteArray(lngBufferSize)

'Compress byte array (data)
lngResult = compress2(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)

'Truncate to compressed size
ReDim Preserve TheData(lngBufferSize - 1)
CopyMemory TheData(0), arrByteArray(0), lngBufferSize

'Set property
lngCompressedSize = UBound(TheData) + 1

'return error code (if any)
CompressByteArray11 = lngResult

End Function

Public Function DecompressByteArray11(TheData() As Byte, OriginalSize As Long) As Long
'decompress a byte array
Dim lngResult As Long
Dim lngBufferSize As Long
Dim arrByteArray() As Byte

lngDecompressedSize = OriginalSize
lngCompressedSize = UBound(TheData) + 1

'Allocate memory for byte array
lngBufferSize = OriginalSize
lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
ReDim arrByteArray(lngBufferSize)

'Decompress data
lngResult = uncompress(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1)

'Truncate buffer to compressed size
ReDim Preserve TheData(lngBufferSize - 1)
CopyMemory TheData(0), arrByteArray(0), lngBufferSize

'return error code (if any)
DecompressByteArray11 = lngResult

End Function

