VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FEPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public yLine As Long
Public xTab As Long
Public xChar As Long
Public xPosMem As Long

Private Sub Class_Initialize()
    yLine = 0
    xTab = 0
    xChar = 0
    xPosMem = -1
End Sub

Public Sub copyfrom(ByRef pos As FEPos)
    yLine = pos.yLine
    xTab = pos.xTab
    xChar = pos.xChar
End Sub
