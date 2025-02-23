VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private m_Header    As StepHeader
Private m_Stack     As Stack
'Private m_Stream    As StreamStr
Private m_Tokenizer As StepTokenizer

Private Sub Class_Initialize()
    Set m_Stack = MNew.List(EDataType.vbObject)
End Sub

Friend Sub New_(aStream As StreamStr)
    'Set m_Stream = aStream
    Set m_Tokenizer = MNew.StepTokenizer(aStream)
End Sub

Public Function IsStepDocument() As Boolean
    Dim tok As StepToken: Set tok = m_Tokenizer.NextToken
    If tok.Value = "ISO-10303-21" Then IsStepDocument = True
End Function

''HEADER;
''FILE_DESCRIPTION(('FileDescription'),'1;1');
''FILE_NAME('filename.step','2003-12-27T11:57:53',('Lothar Klein'),('LKSoft'),'','IDA-STEP','');
''FILE_SCHEMA(('AUTOMOTIVE_DESIGN { 1 0 10303 214 2 1 1}'));
''ENDSEC;
'm_s = m_s & "HEADER;" & vbCrLf
'With obj
'    m_s = m_s & "FILE_DESCRIPTION(('" & IIf(Len(.FileDescription), .FileDescription, "FileDescription") & "'),'" & IIf(Len(.FileImplLevel), .FileImplLevel, "1;1") & "');" & vbCrLf
'    m_s = m_s & "FILE_NAME('" & IIf(Len(.FileName), .FileName, "filename.step") & "','" & .TimeStampStr & "',('" & .Author & "'),('" & .Organization & ")','" & .PreProcVersion & "','" & .OriginSystem & "','" & .Authorization & "');" & vbCrLf
'    m_s = m_s & "FILE_SCHEMA(('" & .FileSchema & "'));" & vbCrLf
'End With
'm_s = m_s & "ENDSEC;" & vbCrLf


Private Function ReadHeader() As StepHeader
    Dim bReadon As Boolean
    Dim tok As StepToken
    
    Set tok = m_Tokenizer.NextToken
    If Not tok Is Nothing Then If tok.TokenType = tkt_Identifier Then bReadon = UCase(tok.Value) = "HEADER"
    If Not bReadon Then Exit Function
    
        Set tok = m_Tokenizer.NextToken
        If Not tok Is Nothing Then If tok.TokenType = tkt_Identifier Then bReadon = UCase(tok.Value) = "FILE_DESCRIPTION"
        If Not bReadon Then Exit Function
        
        Set tok = m_Tokenizer.NextToken
        If Not tok Is Nothing Then If tok.TokenType = tkt_Identifier Then bReadon = UCase(tok.Value) = "FILE_DESCRIPTION"
        If Not bReadon Then Exit Function


'    'Set ReadHeader = New StepHeader
'    With ReadHeader
'        .FileDescription
'        .FileImplLevel
'        .FileName
'        .TimeStamp
'        .Author
'        .Organization
'        .PreProcVersion
'        .OriginSystem
'        .Authorization
'
''FILE_SCHEMA
'        .FileSchema
'
'        .FileDescription
'        .FileImplLevel
'
'        .FileName
'        .TimeStamp
'        .Author
'        .Organization
'        .PreProcVersion
'        .OriginSystem
'        .Authorization
'        .FileSchema
'    End With
End Function

Public Function NextStepObject() As StepObject
    Dim tok As StepToken
    Do Until m_Tokenizer.IsEOF
        Set tok = m_Tokenizer.NextToken
        
        
    Loop
End Function

