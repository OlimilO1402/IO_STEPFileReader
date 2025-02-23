VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Objects As List 'Of StepObject

Private Sub Class_Initialize()
    Set m_Objects = MNew.List(vbObject)
End Sub

Public Function Add(aStepObject As StepObject) As StepObject
    Set Add = m_Objects.Add(aStepObject)
    Add.Hash = m_Objects.Version
End Function

Public Property Get Count() As Long
    Count = m_Objects.Count
End Property

Public Property Get Item(ByVal Index As Long) As StepObject
    Set Item = m_Objects.Item(Index)
End Property

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeStepData Me
End Sub
