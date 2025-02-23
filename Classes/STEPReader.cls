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
Private m_Header    As StepHeader
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

Public Property Get StepHeader() As StepHeader
    If m_Header Is Nothing Then Set m_Header = ReadHeader
    Set StepHeader = m_Header
End Property


'Public ISO1030321      As String
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
''Public TimeStamp       As String
'Private m_TimeStamp    As Date
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




Private Function ReadHeader() As StepHeader
    Dim bReadon As Boolean
    Dim tok As StepToken
    Do Until m_Tokenizer.IsEOF
        Set tok = m_Tokenizer.NextToken
        If Not tok Is Nothing Then
            If tok.TokenType = tkt_Identifier Then bReadon = tok.Value = "ISO-10303-21"
            If Not bReadon Then Exit Function
        End If
        Set tok = m_Tokenizer.NextToken
        If Not tok Is Nothing Then
            If tok.TokenType = tkt_Identifier Then bReadon = tok.Value = "FILE_DESCRIPTION"
            If Not bReadon Then Exit Function
        End If
    Loop

    'Set ReadHeader = New StepHeader
    With ReadHeader
        .FileDescription
        .FileImplLevel
        .FileName
        .TimeStamp
        .Author
        .Organization
        .PreProcVersion
        .OriginSystem
        .Authorization
        
'FILE_SCHEMA
        .FileSchema
        
        .FileDescription
        .FileImplLevel
        
        .FileName
        .TimeStamp
        .Author
        .Organization
        .PreProcVersion
        .OriginSystem
        .Authorization
        .FileSchema
    End With
End Function

Public Function NextStepObject() As StepObject
    Do Until m_Tokenizer.NextToken Is Nothing
        
    Loop
End Function

