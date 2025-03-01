VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nTupel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse stellt ein n-Tupel dar das in der icx-datei so dargestellt wird:
'entweder: (0.1,0.2), wenn es sich um numerische Werte handelt
'oder:     (#15,#16), wenn es sich um Objekte handelt
'ICXTupel sind keine eigenständigen icx-Objekte daher gibt es keine Liste für ICXTupel im ICXDocument

Private m_Typ      As VbVarType
Private m_nCount   As Long
Private m_Values() As Variant

Private Sub Class_Initialize()
    m_Typ = vbDouble
    m_nCount = 2
    ReDim m_Values(0 To m_nCount - 1)
End Sub

Friend Sub New_(Of_Typ As VbVarType, ByVal nCount As Long)
    m_Typ = Of_Typ
    If 0 < nCount Then
        m_nCount = nCount
        ReDim m_Values(0 To m_nCount - 1)
    End If
End Sub

''copy To MNew:
'Public Function nTupel(Of_Typ As VbVarType, ByVal nCount As Long) As nTupel
'    Set nTupel = New nTupel: nTupel.New_ Of_Typ, nCount
'End Function

Public Property Get Count() As Long
    Count = m_nCount
End Property

Public Sub Add(Value)
    ReDim Preserve m_Values(0 To m_nCount)
    If VarType(Value) = vbObject Then
        Set m_Values(m_nCount) = Value
    Else
        m_Values(m_nCount) = Value
    End If
    m_nCount = m_nCount + 1
End Sub

Public Sub AddToStepData(Data As StepData)
    Dim Obj As StepObject
    Dim i As Long
    For i = 0 To m_nCount - 1
        Set Obj = m_Values(i)
        'If TypeName(Obj) = "StepObject" Then
        Obj.AddToStepData Data
    End If
End Sub

Public Property Get Item(ByVal Index As Long)
    If m_nCount <= Index Then Exit Property
    If VarType(m_Values(Index)) = vbObject Then
        Set Item = m_Values(Index)
    Else
        Item = m_Values(Index)
    End If
End Property

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeNTupel Me
End Sub

Public Function ToStr() As String
    Dim s As String
    If AllValuesAreEmpty Then
        ToStr = "$"
        Exit Function
    End If
    Dim i As Long, si As String, Obj As Object
    If m_Typ = vbObject Then
        For i = 0 To m_nCount - 1
            Set Obj = m_Values(i)
            If Not Obj Is Nothing Then
                si = Obj.ToStr
                s = s & IIf(Len(s) And Len(si), ",", "") & si
            End If
            si = vbNullString
        Next
    Else
        For i = 0 To m_nCount - 1
            If Not IsEmpty(m_Values(i)) Then
                si = Trim(str(m_Values(i)))
                s = s & IIf(Len(s) And Len(si), ",", "") & si
            End If
            si = vbNullString
        Next
    End If
    If InStr(s, ",") Then s = "(" & s & ")"
    ToStr = s
End Function

Public Property Get AllValuesAreEmpty() As Boolean
    Dim i As Long
    For i = 0 To m_nCount - 1
        If Not IsEmpty(m_Values(i)) Then
            If m_Typ = vbObject Then
                If Not m_Values(i) Is Nothing Then Exit Property
            Else
                Exit Property
            End If
        End If
    Next
    AllValuesAreEmpty = True
End Property

Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Property Get Values(ByVal Index As Long)
    If m_nCount <= Index Then Exit Property
    If m_Typ = vbObject Then Set Values = m_Values(Index) Else Values = m_Values(Index)
End Property
Public Property Let Values(ByVal Index As Long, Value)
    If m_nCount <= Index Then Exit Property
    If m_Typ <> vbObject Then m_Values(Index) = Value
End Property
Public Property Set Values(ByVal Index As Long, Value)
    If m_nCount <= Index Then Exit Property
    If m_Typ = vbObject Then Set m_Values(Index) = Value
End Property

Public Property Get Value0(): If m_Typ = vbObject Then Set Value0 = Me.Values(0) Else Value0 = Me.Values(0): End Property
Public Property Let Value0(Value):     Me.Values(0) = Value: End Property
Public Property Set Value0(Value): Set Me.Values(0) = Value: End Property

Public Property Get Value1(): If m_Typ = vbObject Then Set Value1 = Me.Values(1) Else Value1 = Me.Values(1): End Property
Public Property Let Value1(Value):     Me.Values(1) = Value: End Property
Public Property Set Value1(Value): Set Me.Values(1) = Value: End Property

Public Property Get Value2(): If m_Typ = vbObject Then Set Value2 = Me.Values(2) Else Value2 = Me.Values(2): End Property
Public Property Let Value2(Value):     Me.Values(2) = Value: End Property
Public Property Set Value2(Value): Set Me.Values(2) = Value: End Property

