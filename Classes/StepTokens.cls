VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'StepTokens could also be a StepToken
Private m_List As List 'Of StepToken

Private Sub Class_Initialize()
    Set m_List = MNew.List(vbObject)
End Sub

Public Sub Add(obj As StepToken)
    m_List.Add obj
End Sub

Public Property Get Item(ByVal Index As Long) As StepToken
    Set Item = m_List.Item(Index)
End Property

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeStepTokens Me
End Sub

Public Function ContainsObjectHash() As Boolean
    Dim i As Long, tok As StepToken
    For i = 0 To m_List.Count - 1
        Set tok = m_List.Item(i)
        If tok.TokenType = tkt_ObjectHash Then
            ContainsObjectHash = True
            Exit Function
        End If
    Next
End Function

