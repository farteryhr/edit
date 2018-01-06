VERSION 5.00
Begin VB.Form frmEditor 
   AutoRedraw      =   -1  'True
   Caption         =   "fartEditor"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "Curlz MT"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer tmrMov 
      Interval        =   10
      Left            =   0
      Top             =   1440
   End
   Begin VB.PictureBox picEditor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   1395
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As New OArray
Dim FE As New FEBox

Private filepath As String
Private fileCode As Long
Private fileBom As Boolean

Private Sub Form_Initialize()
    FE.bind picEditor
    If Command <> "" Then
        If Command Like """*""" Then
            FE.loadtext readfile(Mid$(Command, 2, Len(Command) - 2))
        Else
            FE.loadtext readfile(Command)
        End If
    Else
        Dim s As String
        s = "drag-drop to load file\nright-click to change font\nno selection yet\n\n1\t2\t3\t4\t\none\ttwo\tthreeeeee\tfour\t\n\nwhile(--c){\n\tif(a==b){\n\t\twow\n\t}\n}\n"
        s = Replace(s, "\n", vbCrLf)
        s = Replace(s, "\t", vbTab)
        FE.loadtext s
    End If
    'FE.render
End Sub

Private Sub Form_Resize()
    picEditor.width = Me.ScaleWidth
    picEditor.height = Me.ScaleHeight
    FE.render
End Sub

Private Sub picEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "Keydown", KeyCode, Shift
    Select Case KeyCode
        Case vbKeyDown: FE.cursorMove FE.cursor, 2
        Case vbKeyUp:  FE.cursorMove FE.cursor, 3
        Case vbKeyRight:  FE.cursorMove FE.cursor, 0
        Case vbKeyLeft:  FE.cursorMove FE.cursor, 1
        Case vbKeyEnd: If Shift And 2 Then FE.cursorMove FE.cursor, 6 Else FE.cursorMove FE.cursor, 4
        Case vbKeyHome: If Shift And 2 Then FE.cursorMove FE.cursor, 7 Else FE.cursorMove FE.cursor, 5
        Case vbKeyPageDown: FE.cursorMove FE.cursor, 8
        Case vbKeyPageUp: FE.cursorMove FE.cursor, 9
        Case vbKeyTab: FE.insertChar 9, FE.cursor
        Case vbKeyReturn: FE.insertChar 13, FE.cursor
        Case vbKeyBack: FE.insertChar 8, FE.cursor
        Case Else:
            If KeyCode = vbKeyS And Shift And 2 Then
                savefile filepath, FE.gettext
            Else
                Exit Sub
            End If
    End Select
    If Not FE.moving Then
        FE.render
    End If
End Sub

Private Sub picEditor_KeyPress(KeyAscii As Integer)
    'Debug.Print "Keypress", KeyAscii
    If KeyAscii < 0 Or KeyAscii >= &H20 Then
        FE.insertChar KeyAscii, FE.cursor
        If Not FE.moving Then
            FE.render
        End If
    End If
End Sub

Private Sub picEditor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Set FE.cursor = FE.XYscrmovToPos(x, y)
        FE.render
    End If
End Sub

Private Function readfile(path As String) As String
    Dim chars() As Byte
    Dim retstr As String
    Dim lenfile As Long
    lenfile = FileLen(path)
    If lenfile <= 0 Then
        retstr = ""
        fileCode = 65001
        fileBom = False
        GoTo final
    End If
    ReDim chars(lenfile - 1)
    Open path For Binary As #1
    Get #1, , chars
    Close #1
    If UBound(chars) >= 3 - 1 Then
        If chars(0) = &HEF And chars(1) = &HBB And chars(2) = &HBF Then
            retstr = StrConvToWide(RightB(chars, UBound(chars) + 1 - 3), 65001)
            fileCode = 65001
            fileBom = True
            GoTo final
        End If
    End If
    If UBound(chars) >= 2 - 1 Then
        If chars(0) = &HFF And chars(1) = &HFE Then
            retstr = RightB(chars, UBound(chars) + 1 - 2)
            fileCode = -1
            fileBom = True
            GoTo final
        End If
    End If
    'verify utf-8
    Dim ptr As Long, ntry As Long, xhtry As Long
    ptr = 0
    Do While ptr < lenfile
        If (chars(ptr) And 128) = 0 Then
            ntry = 0
        ElseIf (chars(ptr) And 224) = 192 Then
            ntry = 1
        ElseIf (chars(ptr) And 240) = 224 Then
            ntry = 2
        ElseIf (chars(ptr) And 248) = 240 Then
            ntry = 3
        ElseIf (chars(ptr) And 252) = 248 Then
            ntry = 4
        ElseIf (chars(ptr) And 254) = 252 Then
            ntry = 5
        Else
            GoTo ansi
        End If
        ptr = ptr + 1
        If ntry = 0 Then
            'accelerate
        ElseIf ptr + ntry > lenfile Then
            GoTo ansi
        Else
            For xhtry = 0 To ntry - 1
                If (chars(ptr + xhtry) And 192) <> 128 Then
                    GoTo ansi
                End If
            Next xhtry
            ptr = ptr + ntry
        End If
    Loop
    retstr = StrConvToWide(chars, 65001)
    fileCode = 65001
    fileBom = False
    GoTo final
ansi:
    'Debug.Print "utf8fail:" & ptr
    retstr = StrConv(chars, vbUnicode)
    fileCode = 0
    fileBom = False
final:
    filepath = path
    readfile = retstr
End Function

Private Sub savefile(path As String, strfile As String)
    Dim chars() As Byte
    Open path For Output As #1
    Close #1
    Open path For Binary As #1
    
    If fileCode = -1 Then
        chars = strfile
        If fileBom Then
            Put #1, , CByte(&HFF)
            Put #1, , CByte(&HFE)
        End If
        Put #1, , chars
    ElseIf fileCode = 65001 Then
        chars = StrConvFromWide(strfile, fileCode)
        If fileBom Then
            Put #1, , CByte(&HEF)
            Put #1, , CByte(&HBB)
            Put #1, , CByte(&HBF)
        End If
        Put #1, , chars
    ElseIf fileCode = 0 Then
        chars = StrConv(strfile, vbFromUnicode)
        Put #1, , chars
    Else
        chars = StrConvFromWide(strfile, fileCode)
        Put #1, , chars
    End If
    Close #1
End Sub

Private Sub picEditor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        picEditor.FontBold = False
        picEditor.FontItalic = False
        picEditor.FontSize = False
        picEditor.FontStrikethru = False
        Dim strin As String
        strin = InputBox("font name:", "FEdit", picEditor.FontName)
        If strin <> "" Then picEditor.FontName = strin
        strin = InputBox("font size:", "FEdit", picEditor.FontSize)
        If strin <> "" Then picEditor.FontSize = Val(strin)
        FE.updateCache
        FE.updateHeight 0, FE.lines.size
        FE.updateInnerWidth 0, FE.lines.size
        FE.render
    End If
End Sub

Private Sub picEditor_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intFile As Integer
    With Data
        If .GetFormat(15) = False Then Exit Sub
        If .Files.Count <> 1 Then Exit Sub
        Me.Caption = "loading"
        FE.loadtext readfile(.Files.Item(1))
        Me.Caption = "fartEditor"
        FE.render
    End With
End Sub

Private Sub tmrMov_Timer()
    Dim newY As Single, newX As Single
    If FE.moving Then
        newY = (FE.scrY - FE.scrmovY) * 0.3 + FE.scrmovY
        newX = (FE.scrX - FE.scrmovX) * 0.3 + FE.scrmovX
        FE.scrmovY = newY
        FE.scrmovX = newX
        If Round(newX) <> FE.scrmovlstX Or Round(newY) <> FE.scrmovlstY Then
            FE.render
            FE.scrmovlstY = Round(newY)
            FE.scrmovlstX = Round(newX)
        ElseIf Round(newX) = FE.scrX And Round(newY) = FE.scrY Then
            FE.render
            FE.moving = False
        End If
    End If
End Sub
