VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ISO1030321      As String
'FILE_DESCRIPTION (
'                     /* description */          ('A minimal AP214 example with a single part'),
'                     /* implementation_level */ '2;1'
'                 );
Public FileDescription As String
Public FileImplLevel   As String

'FILE_NAME        (
'                     /* name */                 'demo',
'                     /* time_stamp */           '2003-12-27T11:57:53',
'                     /* author */               ('Lothar Klein'),
'                     /* organization */         ('LKSoft'),
'                     /* preprocessor_version */ ' ',
'                     /* originating_system */   'IDA-STEP',
'                     /* authorization */        ' '
'                 );
Public FileName        As String
'Public TimeStamp       As String
Private m_TimeStamp    As Date
Public Author          As String
Public Organization    As String
Public PreProcVersion  As String
Public OriginSystem    As String
Public Authorization   As String

'FILE_SCHEMA      (
'                                                ('AUTOMOTIVE_DESIGN { 1 0 10303 214 2 1 1}')
'                 );
Public FileSchema      As String

Friend Sub New_(FDescr As String, FImpl As String, PFNam As String, Auth As String, Organis As String)
    ISO1030321 = "ISO-10303-21": FileDescription = FDescr: FileImplLevel = FImpl
    FileName = PFNam: m_TimeStamp = Now: Author = Auth: Organization = Organis
End Sub
''Copy To MNew.bas
'Public Function StepHeader(ByVal FDescr As String, ByVal FImpl As String, ByVal PFNam As String, ByVal Auth As String, ByVal Organis As String) As StepHeader
'    Set StepHeader = New StepHeader: StepHeader.New_ FDescr, FImpl, PFNam, Auth, Organis
'End Function

Public Sub Serialize(Serializer As StepSerializer)
    Serializer.SerializeStepHeader Me
End Sub

Public Property Get TimeStamp() As Date
    TimeStamp = m_TimeStamp
End Property
Public Property Let TimeStamp(ByVal Value As Date)
    m_TimeStamp = Value
End Property

Public Property Get TimeStamp_ToStr() As String
    Dim d As Date: d = m_TimeStamp
    TimeStamp_ToStr = Year(d) & "-" & Month(d) & "-" & Day(d) & "T" & Hour(d) & ":" & Minute(d) & ":" & Second(d)
End Property
