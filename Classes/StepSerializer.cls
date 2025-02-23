VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_s As String
Private m_IsEndAdded As Boolean

Private Sub Class_Initialize()
    m_s = "ISO-10303-21;" & vbCrLf
End Sub

Public Sub SerializeStepDocument(obj As StepDocument)
    m_s = "ISO-10303-21;" & vbCrLf
    With obj
        .Header.Serialize Me
        Dim i As Long, ObjData As StepData
        For i = 0 To .Count - 1
            Set ObjData = .Item(i)
            ObjData.Serialize Me
        Next
    End With
    m_s = m_s & "END-ISO-10303-21;"
    m_IsEndAdded = True
End Sub

Public Sub SerializeStepHeader(obj As StepHeader)
    'HEADER;
    'FILE_DESCRIPTION(('FileDescription'),'1;1');
    'FILE_NAME('filename.step','2003-12-27T11:57:53',('Lothar Klein'),('LKSoft'),'','IDA-STEP','');
    'FILE_SCHEMA(('AUTOMOTIVE_DESIGN { 1 0 10303 214 2 1 1}'));
    'ENDSEC;
    m_s = m_s & "HEADER;" & vbCrLf
    With obj
        m_s = m_s & "FILE_DESCRIPTION(('" & IIf(Len(.FileDescription), .FileDescription, "FileDescription") & "'),'" & IIf(Len(.FileImplLevel), .FileImplLevel, "1;1") & "');" & vbCrLf
        m_s = m_s & "FILE_NAME('" & IIf(Len(.FileName), .FileName, "filename.step") & "','" & .TimeStampStr & "',('" & .Author & "'),('" & .Organization & "'),'" & .PreProcVersion & "','" & .OriginSystem & "','" & .Authorization & "');" & vbCrLf
        m_s = m_s & "FILE_SCHEMA(('" & .FileSchema & "'));" & vbCrLf
    End With
    m_s = m_s & "ENDSEC;" & vbCrLf
End Sub

Public Sub SerializeStepData(obj As StepData)
    m_s = m_s & "DATA;" & vbCrLf
    With obj
        Dim i As Long, StepObj As StepObject
        For i = 0 To .Count - 1
            Set StepObj = .Item(i)
            StepObj.Serialize Me
        Next
    End With
    m_s = m_s & "ENDSEC;" & vbCrLf
End Sub

Public Sub SerializeStepObject(obj As StepObject)
    With obj
        m_s = m_s & "#" & .Hash & "=" & .Name ' & "("
        .Tokens.Serialize Me
    End With
    m_s = m_s & ";" & vbCrLf
End Sub

Public Sub SerializeStepTokens(obj As StepTokens)
    m_s = m_s & "("
    With obj
        Dim tok As StepToken
        Dim i As Long, u As Long: u = .Count - 1
        For i = 0 To u
            Set tok = .Item(i): tok.Serialize Me
            If i < u Then m_s = m_s & ","
        Next
    End With
    m_s = m_s & ")"
End Sub

Public Sub SerializeStepToken(obj As StepToken)
    Dim stks As StepTokens
    With obj
        Select Case .TokenType
        Case EStepTokenType.tkt_Boolean:            m_s = m_s & "." & IIf(.Value, "T", "F") & "."
        Case EStepTokenType.tkt_Identifier:         m_s = m_s & "#" & CStr(CLng(.Value))
        Case EStepTokenType.tkt_EmptyOrDefault:     m_s = m_s & "$"
        Case EStepTokenType.tkt_EnumIdentifier:     m_s = m_s & "." & CStr(.Value) & "."
        Case EStepTokenType.tkt_nTupelList:         Set stks = .Value
                                                    If Not stks Is Nothing Then If stks.Count = 0 Then m_s = m_s & "$" Else stks.Serialize Me
        Case EStepTokenType.tkt_NumericFlt:         m_s = m_s & Trim(Str(Val(.Value)))
        
        Case EStepTokenType.tkt_NumericInt:         m_s = m_s & CStr(CLng(.Value))
        Case EStepTokenType.tkt_ObjectHash:         m_s = m_s & "#" & CStr(CLng(.Value))
                                    
        Case EStepTokenType.tkt_String:             m_s = m_s & "'" & CStr(.Value) & "'"
        
        End Select
    End With
End Sub

Public Function ToStr() As String
    ToStr = m_s & IIf(m_IsEndAdded, "", vbCrLf & "END-ISO-10303-21;")
End Function
