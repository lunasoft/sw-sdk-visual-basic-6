Attribute VB_Name = "basUtf8FromString"
Attribute VB_Name = "basUtf8FromString"
' basUtf8FromString

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>

Option Explicit

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function
