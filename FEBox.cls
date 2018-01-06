VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FEBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public lines As OArray ' of FELine

Public cursor As FEPos

Public scrY As Long, scrX As Long
Public scrmovY As Single, scrmovX As Single
Public scrmovlstY As Long, scrmovlstX As Long
Public moving As Boolean
Public rendering As Boolean
Public options As FEOptions

Public refWindow As PictureBox

Private width As Long
Private height As Long
Private spaceWidth As Long
Private lineHeight As Long
Private origSize As Single

Public Function XYToPos(x As Long, y As Long, Optional yonly As Boolean = False) As FEPos
    Dim yLine As Long, xTab As Long, xChar As Long
    Dim refLine As FELine, refTab As FETab
    Dim xoffset As Long
    yLine = lines.atwei(y)
    If yLine < 0 Then yLine = 0
    If yLine >= lines.size Then yLine = lines.size - 1
    If Not yonly Then
        Set refLine = lines.at(yLine)
        xTab = refLine.tabs.atwei(x)
        If xTab < 0 Then xTab = 0
        If xTab >= refLine.tabs.size Then xTab = refLine.tabs.size - 1
        Set refTab = refLine.tabs.at(xTab)
        xoffset = x - refLine.tabs.weiacc(xTab)
        If refLine.isEmptyLine Then refWindow.FontSize = origSize * options.emptyHeight
        Dim bsl As Long, bsr As Long, bsm As Long
        bsl = 0
        bsr = Len(refTab.text)
        Do While bsl < bsr
            bsm = (bsl + bsr) \ 2
            If refWindow.TextWidth(Left(refTab.text, bsm + 1)) >= xoffset Then
                bsr = bsm
            Else
                bsl = bsm + 1
            End If
        Loop
        xChar = bsl
        If bsl < Len(refTab.text) Then
            If refWindow.TextWidth(Left(refTab.text, bsl + 1)) - xoffset < xoffset - refWindow.TextWidth(Left(refTab.text, bsl)) Then
                xChar = bsl + 1
            End If
        Else
            If xTab < refLine.tabs.size - 1 Then
                If refTab.width - xoffset < xoffset - refTab.innerWidth Then
                    xChar = 0
                    xTab = xTab + 1
                End If
            End If
        End If
        If refLine.isEmptyLine Then refWindow.FontSize = origSize
    End If
    Dim pos As New FEPos
    With pos
        .yLine = yLine
        .xTab = xTab
        .xChar = xChar
    End With
    Set XYToPos = pos
End Function

Public Function XYscrmovToPos(ByVal x As Long, ByVal y As Long, Optional ByVal yonly As Boolean = False) As FEPos
    Set XYscrmovToPos = XYToPos(Round(scrmovX) + x, Round(scrmovY) + y, yonly)
End Function

Public Sub posToXY(ByRef x As Long, ByRef y As Long, ByRef pos As FEPos)
    Dim refLine As FELine, ox As Long
    y = lines.weiacc(pos.yLine)
    Set refLine = lines.at(pos.yLine)
    x = refLine.tabs.weiacc(pos.xTab)
    If refLine.isEmptyLine Then refWindow.FontSize = origSize * options.emptyHeight
    ox = refWindow.TextWidth( _
        Left(refLine.tabs.at(pos.xTab).text, pos.xChar) _
    )
    If refLine.isEmptyLine Then refWindow.FontSize = origSize
    x = x + ox
End Sub

Public Sub posToXYScrmov(ByRef x As Long, ByRef y As Long, ByRef pos As FEPos)
    posToXY x, y, pos
    x = x - Round(scrmovX)
    y = y - Round(scrmovY)
End Sub

Public Function isEmptyText(ByRef s As String) As Boolean
    Dim l As Long, xhc As Long, curChar As Integer
    l = Len(s)
    isEmptyText = True
    If l = 0 Then Exit Function
    For xhc = 0 To l - 1
        curChar = AscW(Mid(s, xhc + 1, 1))
        Select Case curChar
            Case 32, 59, 44, 42, 47, 92, 40, 41, 91, 93, 123, 125:
            Case Else:
                isEmptyText = False
                Exit Function
        End Select
    Next xhc
End Function

Public Sub updateHeight(ByVal l As Long, ByVal r As Long)
    Dim xhl As Long, xht As Long
    Dim refLine As FELine, refTab As FETab
    Dim hei As Long, heiMax As Long, isEmptyLine As Boolean
    Dim heiNew() As Long
    If l < 0 Or r < l Or r > lines.size Then Err.Raise 9
    If l = r Then Exit Sub
    ReDim heiNew(r - l - 1) As Long
    For xhl = l To r - 1
        Set refLine = lines.at(xhl)
        isEmptyLine = True
        For xht = 0 To refLine.tabs.size - 1
            Set refTab = refLine.tabs.at(xht)
            'hei = refWindow.TextHeight(refTab.text)
            hei = lineHeight
            If heiMax < hei Then heiMax = hei
            If isEmptyText(refTab.text) = False Then isEmptyLine = False
        Next xht
        refLine.isEmptyLine = isEmptyLine
        If isEmptyLine Then heiMax = CLng(heiMax * options.emptyHeight)
        refLine.height = heiMax
        heiNew(xhl - l) = heiMax
    Next xhl
    lines.weiset l, r - l, heiNew
    Erase heiNew
End Sub

Public Sub tabInnerWidth(ByRef oTab As FETab)
    oTab.innerWidth = refWindow.TextWidth(oTab.text)
End Sub

Public Sub updateWidth(ByVal l As Long, ByVal r As Long)
    Dim xhl As Long, xht As Long, xhlm As Long
    Dim refLine As FELine, refTab As FETab
    Dim tabWidthMax(100) As Long, tabColFirst(100) As Long
    Dim lastTab As Long, curTab As Long
    Dim widNew() As Long
    
    Dim cnt As Long
    If l < 0 Or r < l Or r > lines.size Then Err.Raise 9
    If l = r Then Exit Sub
    lastTab = 0
    For xhl = l To r
        If xhl = r Then 'special finalize
            Set refLine = Nothing
            curTab = 0
        Else
            Set refLine = lines.at(xhl)
            curTab = refLine.tabs.size
        End If
        curTab = curTab - 1 'the last tab doesn't participate
        For xht = 0 To curTab - 1 'won't enter on finalize
            Set refTab = refLine.tabs.at(xht)
            If xht >= lastTab Then
                tabColFirst(xht) = xhl
                tabWidthMax(xht) = options.tabMin * spaceWidth
            End If
            If tabWidthMax(xht) < refTab.innerWidth + spaceWidth * options.tabMargin Then
                tabWidthMax(xht) = refTab.innerWidth + spaceWidth * options.tabMargin
            End If
            cnt = cnt + 1
        Next xht
        If curTab < lastTab Then
            For xht = IIf(curTab < 0, 0, curTab) To lastTab - 1
                For xhlm = tabColFirst(xht) To xhl - 1 'last but line #r - virtual
                    lines.at(xhlm).tabs.at(xht).width = tabWidthMax(xht)
                Next xhlm
            Next xht
        End If
        If curTab >= 0 Then refLine.tabs(curTab).width = refLine.tabs(curTab).innerWidth + spaceWidth * options.tabMargin
        lastTab = curTab
    Next xhl
    For xhl = l To r - 1
        Set refLine = lines.at(xhl)
        ReDim widNew(refLine.tabs.size - 1)
        For xht = 0 To refLine.tabs.size - 1
            widNew(xht) = refLine.tabs.at(xht).width
        Next xht
        refLine.tabs.weiset 0, refLine.tabs.size, widNew
        Erase widNew
    Next xhl
End Sub

Public Sub updateInnerWidth(ByVal l As Long, ByVal r As Long)
    Dim xhl As Long, xht As Long
    Dim refLine As FELine, refTab As FETab
    If l < 0 Or r < l Or r > lines.size Then Err.Raise 9
    If l = r Then Exit Sub
    For xhl = l To r - 1
        Set refLine = lines.at(xhl)
        For xht = 0 To refLine.tabs.size - 1
            Set refTab = refLine.tabs.at(xht)
            tabInnerWidth refTab
        Next xht
    Next xhl
End Sub

Public Sub bind(pic As PictureBox)
    Set refWindow = pic
    refWindow.ScaleMode = 3
End Sub

Public Sub updateCache()
    width = CLng(refWindow.ScaleWidth)
    height = CLng(refWindow.ScaleHeight)
    spaceWidth = CLng(refWindow.TextWidth(" "))
    lineHeight = CLng(refWindow.TextHeight(" "))
    origSize = refWindow.FontSize
End Sub

Public Sub loadtext(s As String)
    Dim arrLines() As String
    Dim arrTabs() As String
    Dim xhl As Long, xht As Long, nTabs As Long
    Dim oLine As FELine, oTab As FETab
    lines.del 0, lines.size
    If s = "" Then
        ReDim arrLines(0)
        arrLines(0) = ""
    Else
        arrLines = Split(s, vbLf)
    End If
    ReDim arrLinesTabs(UBound(arrLines))
    For xhl = 0 To UBound(arrLines)
        Set oLine = New FELine
        If xhl = UBound(arrLines) Then
            oLine.lineEnd = ""
        ElseIf Right(arrLines(xhl), 1) = vbCr Then
            oLine.lineEnd = vbCrLf
            arrLines(xhl) = Left(arrLines(xhl), Len(arrLines(xhl)) - 1)
        Else
            oLine.lineEnd = vbLf
        End If
        For xht = 0 To Len(arrLines(xhl)) - 1
            If Mid$(arrLines(xhl), xht + 1, 1) <> " " Then Exit For
        Next xht
        nTabs = xht \ options.tabMin
        For xht = 0 To nTabs - 1
            Set oTab = New FETab
            oLine.tabs.add oTab, 0
            Set oTab = Nothing
        Next xht
        If nTabs <> 0 Then
            arrLines(xhl) = Right(arrLines(xhl), Len(arrLines(xhl)) - nTabs * options.tabMin)
        End If
        If arrLines(xhl) = "" Then
            ReDim arrTabs(0)
            arrTabs(0) = ""
        Else
            arrTabs = Split(arrLines(xhl), vbTab)
        End If
        For xht = 0 To UBound(arrTabs)
            Set oTab = New FETab
            oTab.text = arrTabs(xht)
            oLine.tabs.add oTab, 0
            Set oTab = Nothing
        Next xht
        lines.add oLine, 0
        Set oLine = Nothing
    Next xhl
    Set cursor = Nothing
    Set cursor = New FEPos
    scrX = 0
    scrY = 0
    updateCache
    updateHeight 0, lines.size
    updateInnerWidth 0, lines.size
    'updateWidth 0, lines.size
    moving = True
End Sub

Public Function gettext() As String
    Dim arrLines() As String
    Dim arrTabs() As String
    Dim xhl As Long, xht As Long
    Dim refLine As FELine, refTab As FETab
    Dim xhv As Variant
    ReDim arrLines(lines.size - 1)
    For xhl = 0 To lines.size - 1
        Set refLine = lines.at(xhl)
        ReDim arrTabs(refLine.tabs.size - 1)
        For xht = 0 To refLine.tabs.size - 1
            Set refTab = refLine.tabs.at(xht)
            arrTabs(xht) = refTab.text + IIf(xht = refLine.tabs.size - 1, refLine.lineEnd, vbTab)
        Next xht
        arrLines(xhl) = Join(arrTabs, "")
    Next xhl
    gettext = Join(arrLines, "")
End Function

Public Sub render()
    'If rendering Then Exit Sub Else rendering = True
    DoEvents
    Dim linel As Long, liner As Long
    Dim posLU As FEPos, posRB As FEPos
    QPCIn "all"
    QPCIn "cache"
        updateCache
    QPCOut "cache"
    QPCIn "camera"
        Set posLU = XYscrmovToPos(0, 0, True)
        Set posRB = XYscrmovToPos(0, height, True)
        If posRB.yLine < lines.size Then posRB.yLine = posRB.yLine + 1
        Dim xhl As Long, xht As Long
        Dim refLine As FELine, refTab As FETab
        linel = posLU.yLine
        liner = posRB.yLine
    QPCOut "camera"
    QPCIn "update"
        updateWidth linel, liner
    QPCOut "update"
    QPCIn "prepare"
        refWindow.Visible = False
        refWindow.Cls
    QPCOut "prepare"
    QPCIn "tabs"
        For xhl = linel To liner - 1
            Set refLine = lines.at(xhl)
            With refLine.tabs
                If refLine.isEmptyLine Then
                    refWindow.FontSize = origSize * options.emptyHeight
                Else
                    refWindow.FontSize = origSize
                End If
                
                For xht = 0 To .size - 1
                    If .weiacc(xht) < scrmovX + width And .weiacc(xht + 1) > scrmovX Then
                        If .at(xht).text = "" And xht <> .size - 1 Then
                            refWindow.Line (.weiacc(xht) - scrmovX, lines.weiacc(xhl) - scrmovY)- _
                                (.weiacc(xht + 1) - scrmovX - 1, lines.weiacc(xhl + 1) - scrmovY - 1), _
                                RGB(255 - IIf(((xht) Mod 2 = 1) Or xht = .size, 32, 16), 255, 255), BF
                        End If
                        If .at(xht).text <> "" Then
                            refWindow.CurrentX = .weiacc(xht) - scrmovX
                            refWindow.CurrentY = lines.weiacc(xhl) - scrmovY
                            refWindow.Print .at(xht).text
                        End If
                    End If
                    'Debug.Print "t";
                Next xht
            End With
        Next xhl
    QPCOut "tabs"
    QPCIn "final"
        refWindow.FontSize = origSize
        Dim curX As Long, curY As Long
        posToXYScrmov curX, curY, cursor
        refWindow.Line (curX, curY)-(curX, curY + lineHeight * (IIf(lines.at(cursor.yLine).isEmptyLine, options.emptyHeight, 1)))
        refWindow.Visible = True
        refWindow.CurrentX = curX
        refWindow.CurrentX = curY
    QPCOut "final"
    QPCOut "all", True
    DoEvents
    'rendering = False
End Sub

Public Sub insertChar(char As Integer, pos As FEPos)
    Dim refTab As FETab, refPreTab As FETab, refLine As FELine, refPreLine As FELine
    Dim oTab As FETab, oLine As FELine
    Dim sTemp As String
    With pos
        Set refLine = lines.at(.yLine)
        Set refTab = refLine.tabs.at(.xTab)
        If char = 8 Then
            If .xChar = 0 Then
                If .xTab = 0 Then
                    If .yLine = 0 Then
                        Beep
                    Else
                        Set refPreLine = lines.at(.yLine - 1)
                        .xTab = refPreLine.tabs.size - 1
                        Set refPreTab = refPreLine.tabs.at(.xTab)
                        .xChar = Len(refPreTab.text)
                        refPreTab.text = refPreTab.text + refTab.text
                        refPreLine.tabs.addArr refLine.tabs
                        Set refTab = Nothing
                        refPreLine.tabs.del .xTab + 1, 1
                        Set refLine = Nothing
                        lines.del .yLine, 1
                        .yLine = .yLine - 1
                        tabInnerWidth refPreTab
                        updateHeight .yLine, .yLine + 1
                    End If
                Else
                    Set refPreTab = refLine.tabs.at(.xTab - 1)
                    .xChar = Len(refPreTab.text)
                    refPreTab.text = refPreTab.text + refTab.text
                    Set refTab = Nothing
                    lines.at(.yLine).tabs.del .xTab, 1
                    .xTab = .xTab - 1
                    tabInnerWidth refPreTab
                End If
            Else
                refTab.text = Left(refTab.text, .xChar - 1) & Right(refTab.text, Len(refTab.text) - .xChar)
                .xChar = .xChar - 1
                updateHeight .yLine, .yLine + 1
                tabInnerWidth refTab
            End If
        ElseIf char = 9 Then
            Set oTab = New FETab
            oTab.text = Right(refTab.text, Len(refTab.text) - .xChar)
            refTab.text = Left(refTab.text, .xChar)
            refLine.tabs.ins .xTab + 1, oTab, 0
            .xChar = 0
            .xTab = .xTab + 1
            updateHeight .yLine, .yLine + 1
            tabInnerWidth oTab
            tabInnerWidth refTab
            Set oTab = Nothing
        ElseIf char = 13 Then
            Set oTab = New FETab
            Set oLine = New FELine
            oTab.text = Right(refTab.text, Len(refTab.text) - .xChar)
            refTab.text = Left(refTab.text, .xChar)
            oLine.lineEnd = refLine.lineEnd
            refLine.lineEnd = vbCrLf
            Dim xht As Long, autotab As Long
            For xht = 0 To .xTab - 1
                If refLine.tabs.at(xht).text = "" Then
                    oLine.tabs.add New FETab, 0 ' refLine.tabs.wei(xht)
                    autotab = xht + 1
                Else
                    Exit For
                End If
            Next xht
            oLine.tabs.add oTab, 0
            For xht = .xTab + 1 To refLine.tabs.size - 1
                oLine.tabs.add refLine.tabs.at(xht), 0
                tabInnerWidth refTab
            Next xht
            refLine.tabs.del .xTab + 1, refLine.tabs.size - (.xTab + 1)
            lines.ins .yLine + 1, oLine, 0
            tabInnerWidth oTab
            tabInnerWidth refTab
            updateHeight .yLine, .yLine + 2
            Set oTab = Nothing
            Set oLine = Nothing
            .xTab = autotab
            .xChar = 0
            .yLine = .yLine + 1
        Else
            refTab.text = Left(refTab.text, .xChar) & Chr(char) & Right(refTab.text, Len(refTab.text) - .xChar)
            .xChar = .xChar + 1
            updateHeight .yLine, .yLine + 1
            tabInnerWidth refTab
        End If
        .xPosMem = -1
    End With
    'updateWidth 0, lines.size
    viewFollow pos
End Sub

Public Sub cursorMove(ByRef pos As FEPos, dire As Integer)
    Dim x As Long, y As Long, newpos As FEPos
    Dim refLine As FELine
    If dire = 0 Then
        If pos.xChar = Len(lines.at(pos.yLine).tabs.at(pos.xTab).text) Then
            If pos.xTab = lines.at(pos.yLine).tabs.size - 1 Then
                If pos.yLine = lines.size - 1 Then
                    Beep
                Else
                    pos.xChar = 0
                    pos.xTab = 0
                    pos.yLine = pos.yLine + 1
                End If
            Else
                pos.xChar = 0
                pos.xTab = pos.xTab + 1
            End If
        Else
            pos.xChar = pos.xChar + 1
        End If
        pos.xPosMem = -1
    ElseIf dire = 1 Then
        If pos.xChar = 0 Then
            If pos.xTab = 0 Then
                If pos.yLine = 0 Then
                    Beep
                Else
                    pos.xChar = Len(lines.at(pos.yLine - 1).tabs.at(lines.at(pos.yLine - 1).tabs.size - 1).text)
                    pos.xTab = lines.at(pos.yLine - 1).tabs.size - 1
                    pos.yLine = pos.yLine - 1
                End If
            Else
                pos.xChar = Len(lines.at(pos.yLine).tabs.at(pos.xTab - 1).text)
                pos.xTab = pos.xTab - 1
            End If
        Else
            pos.xChar = pos.xChar - 1
        End If
        pos.xPosMem = -1
    ElseIf dire = 2 Or dire = 8 Then
        If pos.yLine = lines.size - 1 Then
            Beep
        Else
            posToXY x, y, pos
            If dire = 2 Then
                y = y + lines.at(pos.yLine).height
            Else
                y = y + height
            End If
            If pos.xPosMem <> -1 Then x = pos.xPosMem
            Set newpos = XYToPos(x, y)
            pos.copyfrom newpos
            If pos.xPosMem = -1 Then pos.xPosMem = x
        End If
    ElseIf dire = 3 Or dire = 9 Then
        If pos.yLine = 0 Then
            Beep
        Else
            posToXY x, y, pos
            If dire = 3 Then
                y = y - 1
            Else
                y = y - height + lines.at(pos.yLine).height - 1
            End If
            If pos.xPosMem <> -1 Then x = pos.xPosMem
            Set newpos = XYToPos(x, y)
            pos.copyfrom newpos
            If pos.xPosMem = -1 Then pos.xPosMem = x
        End If
    ElseIf dire = 4 Then
        Set refLine = lines.at(pos.yLine)
        pos.xTab = refLine.tabs.size - 1
        pos.xChar = Len(refLine.tabs.at(refLine.tabs.size - 1).text)
        pos.xPosMem = -1
    ElseIf dire = 5 Then
        Set refLine = lines.at(pos.yLine)
        For x = 0 To refLine.tabs.size - 1
            If refLine.tabs.at(x).text <> "" Then Exit For
        Next x
        If x < pos.xTab Or (x = pos.xTab And pos.xChar <> 0) Then
            pos.xTab = x
            pos.xChar = 0
        Else
            pos.xTab = 0
            pos.xChar = 0
        End If
        pos.xPosMem = -1
    ElseIf dire = 6 Then
        pos.yLine = lines.size - 1
        Set refLine = lines.at(pos.yLine)
        pos.xTab = refLine.tabs.size - 1
        pos.xChar = Len(refLine.tabs.at(pos.xTab).text)
        pos.xPosMem = -1
    ElseIf dire = 7 Then
        pos.xChar = 0
        pos.xTab = 0
        pos.yLine = 0
        pos.xPosMem = -1
    End If
    viewFollow pos
End Sub
Public Sub viewFollow(ByRef pos As FEPos)
    Dim x As Long, y As Long
    Dim refLine As FELine
    Set refLine = lines.at(pos.yLine)
    posToXY x, y, pos
    If x - scrX < 0 Then
        scrX = x
        moving = True
    ElseIf x >= scrX + width Then
        scrX = x - width + 1
        moving = True
    End If
    If y - scrY < 0 Then
        scrY = y
        moving = True
    ElseIf y + refLine.height > scrY + height Then
        scrY = y + refLine.height - height
        moving = True
    End If
    If lines.weiacc(lines.size) < height Then
        If scrY <> 0 Then
            scrY = 0
            moving = True
        End If
    ElseIf lines.weiacc(lines.size) < scrY + height Then
        scrY = lines.weiacc(lines.size) - height
        moving = True
    End If
'    If refLine.tabs.weiacc(refLine.tabs.size) < width Then
'        scrX = 0
'        moving = True
'    ElseIf refLine.tabs.weiacc(refLine.tabs.size) < scrX + width - 1 Then
'        scrX = refLine.tabs.weiacc(refLine.tabs.size) - width + 1
'        moving = True
'    End If
End Sub
Private Sub Class_Initialize()
    Set lines = New OArray
    Set cursor = New FEPos
    scrY = 0
    scrX = 0
    scrmovY = 0
    scrmovX = 0
    moving = False
    Set options = New FEOptions
End Sub
