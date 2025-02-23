VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Hash As Long
Public Name As String
Private m_Tokens As StepTokens

Private Sub Class_Initialize()
    Set m_Tokens = New StepTokens
End Sub

Friend Sub New_(aHash As Long, aClassName As String)
    Me.Hash = aHash: Me.Name = aClassName
End Sub

''copy to MNew:
'Public Function StepObject(ByVal aHash As Long, ByVal aClassName As String) As StepObject
'    Set StepObject = New StepObject: StepObject.New_ aHash, aClassName
'End Function

Public Property Get Tokens() As StepTokens
    Set Tokens = m_Tokens
End Property

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeStepObject Me
End Sub
