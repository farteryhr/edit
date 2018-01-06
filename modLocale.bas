Attribute VB_Name = "modLocale"
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public Const CP_UTF8 As Long = 65001

Public Function StrConvToWide(sData() As Byte, ByVal CP As Long) As Byte()  ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = MultiByteToWideChar(CP, 0, StrPtr(sData), UBound(sData) + 1, 0, 0)
    ReDim aRetn(0 To 2 * nSize - 1) As Byte
    MultiByteToWideChar CP, 0, VarPtr(sData(0)), UBound(sData) + 1, VarPtr(aRetn(0)), nSize
    StrConvToWide = aRetn
End Function

Public Function StrConvFromWide(ByVal sData As String, ByVal CP As Long) As Byte() ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = WideCharToMultiByte(CP, 0, StrPtr(sData), Len(sData), 0, 0, 0, 0)
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP, 0, StrPtr(sData), Len(sData), VarPtr(aRetn(0)), nSize, 0, 0
    StrConvFromWide = aRetn
End Function

Public Function readbin(path As String) As Byte()
    Dim chars() As Byte
    ReDim chars(FileLen(path) - 1)
    Open path For Binary As #1
    Get #1, , chars
    Close #1
    readbin = chars
End Function
