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

Public Sub SerializeStepHeader(obj As StepHeader)
''FILE_DESCRIPTION (
''                     /* description */          ('A minimal AP214 example with a single part'),
''                     /* implementation_level */ '2;1'
''                 );
'Public FileDescription As String
'Public FileImplLevel   As String
'
''FILE_NAME        (
''                     /* name */                 'demo',
''                     /* time_stamp */           '2003-12-27T11:57:53',
''                     /* author */               ('Lothar Klein'),
''                     /* organization */         ('LKSoft'),
''                     /* preprocessor_version */ ' ',
''                     /* originating_system */   'IDA-STEP',
''                     /* authorization */        ' '
''                 );
'Public FileName        As String
'Public TimeStamp       As String
'Public Author          As String
'Public Organization    As String
'Public PreProcVersion  As String
'Public OriginSystem    As String
'Public Authorization   As String
'
''FILE_SCHEMA      (
''                                                ('AUTOMOTIVE_DESIGN { 1 0 10303 214 2 1 1}')
''                 );
'Public FileSchema      As String
    Dim s As String
    With obj
        s = .ISO1030321:      If Len(s) = 0 Then s = "ISO-10303-21":    m_s = m_s & s & ";" & vbCrLf
        m_s = m_s & "FILE_DESCRIPTION("
        s = .FileDescription: If Len(s) = 0 Then s = "FileDescription": m_s = m_s & "('" & s & "')"
        s = .FileImplLevel:   If Len(s) = 0 Then s = "2;1":             m_s = m_s & ",'" & s & "');" & vbCrLf
        m_s = m_s & "FILE_NAME("
        s = .FileName:        If Len(s) = 0 Then s = "filename.ifc":    m_s = m_s & "'" & s & "'" & vbCrLf
        Dim d As Date: d = Now: Dim ts As String: ts = Year(d) & "-" & Month(d) & "-" & Day(d) & "T" & Hour(d) & ":" & Minute(d) & ":" & Second(d)
        s = .TimeStamp:       If Len(s) = 0 Then s = ts:    m_s = m_s & ",'" & s & "'" & vbCrLf
        
    End With
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
        Case EStepTokenType.tkt_NumericFlt:         m_s = m_s & Trim(str(Val(.Value)))
        
        Case EStepTokenType.tkt_NumericInt:         m_s = m_s & CStr(CLng(.Value))
        Case EStepTokenType.tkt_ObjectHash:         m_s = m_s & "#" & CStr(CLng(.Value))
                                    
        Case EStepTokenType.tkt_String:             m_s = m_s & CStr(.Value): s = Replace(s, "'", "''")
                                                    m_s = m_s & "'" & s & "'"
        
        End Select
    End With
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

Public Sub SerializeStepObject(obj As StepObject)
    With obj
        m_s = m_s & "#" & .Hash & "=" & .ClassName
        .Arguments.Serialize Me
    End With
    m_s = m_s & ";" & vbCrLf
End Sub

Public Function ToStr() As String
    ToStr = m_s
End Function
