Attribute VB_Name = "modClock"
Option Explicit
 
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private freq As Double

Private clockstack(100) As Double, namestack(100) As String, stackptr As Long

Sub QPCInit()
    Dim f As Currency
    QueryPerformanceFrequency f
    freq = CDbl(f) * 10000
    Debug.Print "qpf: " & freq
End Sub

Public Function QPC() As Double
    Dim t As Currency
    QueryPerformanceCounter t
    QPC = Round(CDbl(t) * 10000)
End Function

Public Sub QPCIn(name As String)
    If freq = 0 Then
        QPCInit
    End If
    clockstack(stackptr) = QPC()
    namestack(stackptr) = name
    stackptr = stackptr + 1
End Sub

Public Function QPCOut(name As String, Optional doprint As Boolean = False) As Double
    If stackptr = 0 Then
        Err.Raise 9
    End If
    stackptr = stackptr - 1
    If namestack(stackptr) <> name Then
        Err.Raise 5
    End If
    Dim t As Double
    t = (QPC() - clockstack(stackptr)) / freq
    QPCOut = t
    If True Then Debug.Print name & " " & Format(t * 1000, "0.00ms ")
End Function
