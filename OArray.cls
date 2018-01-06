VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr() As Object
Private arrwei() As Long, arrweiacc() As Long
Private siz As Long
Private cap As Long

Private Sub clrrange(l As Long, r As Long)
    Dim xh As Long
    For xh = l To r - 1
        If Not arr(xh) Is Nothing Then Set arr(xh) = Nothing
    Next xh
End Sub
Public Sub resize(n As Long)
    If n = cap Then
        Exit Sub
    ElseIf n = 0 Then
        If cap <> 0 Then clrrange 0, siz
        Erase arr, arrwei, arrweiacc
    ElseIf cap = 0 Then
        ReDim arr(n), arrwei(n), arrweiacc(n)
        clrrange 0, n
    Else
        If n < cap Then clrrange n, cap
        ReDim Preserve arr(n), arrwei(n), arrweiacc(n)
        If n > cap Then clrrange cap, n
    End If
    cap = n
End Sub

Public Property Get size() As Long
    size = siz
End Property

Public Property Get capa() As Long
    capa = cap
End Property

Public Sub ins(ByVal i As Long, o As Object, ByVal w As Long)
    If i < 0 Or i > siz Then Err.Raise 9
    If cap = siz Then
        If cap = 0 Then
            resize 8
        Else
            resize cap * 2
        End If
    End If
    Dim xh As Long
    For xh = siz - 1 To i Step -1
        Set arr(xh + 1) = arr(xh)
        arrwei(xh + 1) = arrwei(xh)
        arrweiacc(xh + 1) = arrweiacc(xh) + w
    Next xh
    siz = siz + 1
    Set arr(i) = o
    arrwei(i) = w
    arrweiacc(i) = weiacc(i) + w
End Sub

Public Sub insArr(ByVal i As Long, a As OArray)
    If i < 0 Or i > siz Then Err.Raise 9
    If a.size = 0 Then Exit Sub
    Dim sizins As Long, siznew As Long
    sizins = a.size
    siznew = siz + sizins
    If cap < siznew Then
        resize IIf(siznew >= 8, IIf(siznew > cap * 2, siznew, cap * 2), 8)
    End If
    Dim xh As Long
    For xh = siz - 1 To i Step -1
        Set arr(xh + sizins) = arr(xh)
        arrwei(xh + sizins) = arrwei(xh)
        arrweiacc(xh + sizins) = arrweiacc(xh) + a.weiacc(sizins)
    Next xh
    siz = siznew
    For xh = 0 To sizins - 1
        Set arr(i + xh) = a(xh)
        arrwei(i + xh) = a.wei(xh)
        arrweiacc(i + xh) = weiacc(i) + a.weiacc(xh)
    Next xh
End Sub

Public Sub add(o As Object, ByVal w As Long)
    ins siz, o, w
End Sub

Public Sub addArr(a As OArray)
    insArr siz, a
End Sub

Public Sub del(ByVal i As Long, Optional ByVal n As Long = 1)
    If i < 0 Or n < 0 Or i + n > siz Then Err.Raise 9
    If n = 0 Then Exit Sub
    Dim xh As Long, w As Long, siznew As Long, capnew As Long
    siznew = siz - n
    w = weiacc(i + n) - weiacc(i)
    For xh = 0 To n - 1
        Set arr(i + xh) = Nothing
    Next xh
    For xh = i To siznew - 1 Step 1
        Set arr(xh) = arr(xh + n)
        arrwei(xh) = arrwei(xh + n)
        arrweiacc(xh) = arrweiacc(xh + n) - w
    Next xh
    For xh = siznew To siz - 1 Step 1
        Set arr(xh) = Nothing
    Next xh
    If siznew <= cap \ 2 Then
        If siznew = 0 Then
            resize 0
        Else
            capnew = cap
            Do While capnew >= 8 And capnew >= siznew
                capnew = capnew \ 2
            Loop
            resize capnew * 2
        End If
    End If
    siz = siznew
End Sub

Public Property Get at(ByVal i As Long) As Object
Attribute at.VB_UserMemId = 0
    If i < 0 Or i >= siz Then Err.Raise 9
    Set at = arr(i)
End Property

Public Property Set at(ByVal i As Long, o As Object)
    If i < 0 Or i >= siz Then Err.Raise 9
    Set arr(i) = o
End Property

Public Property Get wei(ByVal i As Long) As Long
    If i < 0 Or i >= siz Then Err.Raise 9
    wei = arrwei(i)
End Property

Public Property Let wei(ByVal i As Long, w As Long)
    If i < 0 Or i >= siz Then Err.Raise 9
    If arrwei(i) = w Then Exit Property
    Dim xh As Long
    For xh = i To siz - 1
        arrweiacc(xh) = arrweiacc(xh) + arrwei(i) - w
    Next xh
    arrwei(i) = w
End Property

Public Property Get weiacc(ByVal i As Long) As Long
    If i < 0 Or i > siz Then Err.Raise 9
    If i = 0 Then weiacc = 0 Else weiacc = arrweiacc(i - 1)
End Property

Public Function atwei(ByVal w As Long) As Long
    Dim bsl As Long, bsr As Long, bsm As Long
    If w < 0 Then
        atwei = -1
        Exit Function
    End If
    bsl = 0
    bsr = size
    Do While bsl < bsr
        bsm = (bsl + bsr) \ 2
        If weiacc(bsm + 1) > w Then
            bsr = bsm
        Else
            bsl = bsm + 1
        End If
    Loop
    atwei = bsl
End Function

Public Sub weiset(ByVal i As Long, ByVal n As Long, ByRef w() As Long)
    If i < 0 Or i + n > siz Or n < 0 Then Err.Raise 9
    If n = 0 Then Exit Sub
    Dim cursum As Long, xh As Long
    cursum = weiacc(i)
    For xh = 0 To n - 1
        arrwei(i + xh) = w(xh)
        cursum = cursum + w(xh)
        arrweiacc(i + xh) = cursum
    Next xh
    For xh = i + n To siz - 1
        cursum = cursum + arrwei(xh)
        If arrweiacc(xh) = cursum Then Exit Sub
        arrweiacc(xh) = cursum
    Next xh
End Sub

Private Sub Class_Initialize()
    cap = 0
    siz = 0
End Sub

Private Sub Class_Terminate()
    resize 0
End Sub
