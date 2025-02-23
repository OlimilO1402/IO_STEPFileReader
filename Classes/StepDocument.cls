VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Header As StepHeader
Private m_Datas  As Collection 'Of StepData 'normally there is only one data section but new standards define there could also be more than one

Private Sub Class_Initialize()
    Set m_Header = New StepHeader
    Set m_Datas = New Collection
End Sub

Friend Sub New_(aHeader As StepHeader, aData As StepData)
    Set m_Header = aHeader: Me.Add aData
End Sub

''copy to MNew
'Public Function StepDocument(aHeader As StepHeader, aData As StepData) As StepDocument
'    Set StepDocument = New StepDocument: StepDocument.New_ aHeader, aData
'End Function

Public Property Get Header() As StepHeader
    Set Header = m_Header
End Property

Public Function Add(obj As StepData) As StepData
    Set Add = obj: m_Datas.Add Add
End Function

Public Property Get Count() As Long
    Count = m_Datas.Count
End Property

Public Property Get Item(ByVal Index As Long) As StepData
    Index = Index + 1
    Set Item = m_Datas.Item(Index)
End Property

Public Property Get Data() As StepData
    If m_Datas.Count > 0 Then Set Data = m_Datas.Item(1)
End Property

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeStepDocument Me
End Sub


