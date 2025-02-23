VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Value As String
Private m_Index As Long

Private Sub Class_Initialize()
    m_Index = 1
End Sub

Friend Sub New_(aValue As String)
    m_Value = aValue
End Sub

''copy to MNew:
'Public Function StreamStr(ByVal aValue As String) As StreamStr
'    Set StreamStr = New StreamStr: StreamStr.New_ aValue
'End Function
Public Property Get Index() As Long
    Index = m_Index
End Property

Public Function NextChar() As String
    'If Me.EOF Then Exit Function
    NextChar = MidB(m_Value, m_Index, 2)
    m_Index = m_Index + 2
End Function

Public Function CurChar() As String
    'If Me.EOF Then Exit Function
    Dim i As Long: i = m_Index - 1
    CurChar = MidB(m_Value, i, 2)
End Function

Public Function Lookahead1() As String
    Dim i As Long: i = m_Index + 1
    Lookahead1 = Mid(m_Value, i, 1)
End Function

Public Sub Back1()
    m_Index = m_Index - 2
End Sub
Public Sub Comsume1()
    m_Index = m_Index + 2
End Sub

Public Function EOF() As Boolean
    EOF = Len(m_Value) < m_Index
End Function

Public Function ExtractSkipWS(ByVal nChars As Long) As String
    Dim c As String
    Do Until Me.EOF
        c = Me.NextChar
        Select Case c
        Case " ", vbTab, ";", vbCr, vbLf, vbCrLf
        Case Else: Back1: ExtractSkipWS = MidB(m_Value, m_Index, nChars * 2): Exit Function
        End Select
    Loop
End Function

Public Function ExtractFromTo(ByVal ifrom As Long, ByVal iTo As Long) As String
    Dim n As Long: n = iTo - ifrom + 1
    If n > 0 Then ExtractFromTo = MidB(m_Value, ifrom, n)
End Function
