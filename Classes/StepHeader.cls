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
Public FileDescription As String
Public FileImplLevel   As String
Public FileName        As String
Private m_TimeStamp    As Date
Public Author          As String
Public Organization    As String
Public PreProcVersion  As String
Public OriginSystem    As String
Public Authorization   As String
Public FileSchema      As String

Private Sub Class_Initialize()
    FileDescription = "FileDescription": FileImplLevel = "1;1"
    m_TimeStamp = Now
End Sub

Friend Sub New_(FDescr As String, FImpl As String, PFNam As String, DateTimeStamp, Auth As String, Organis As String, FSchema As String)
    FileDescription = FDescr: FileImplLevel = FImpl: FileName = PFNam: Author = Auth: Organization = Organis: FileSchema = FSchema
    Dim vt As VbVarType: vt = VarType(DateTimeStamp)
    If vt = vbDate Then m_TimeStamp = DateTimeStamp Else If vt = vbString And Len(DateTimeStamp) Then TimeStampStr = DateTimeStamp
End Sub
''Copy To MNew.bas
'Public Function StepHeader(ByVal FDescr As String, ByVal FImpl As String, ByVal PFNam As String, DateTimeStamp, ByVal Auth As String, ByVal Organis As String, ByVal FSchema As String) As StepHeader
'    Set StepHeader = New StepHeader: StepHeader.New_ FDescr, FImpl, PFNam, DateTimeStamp, Auth, Organis, FSchema
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

Public Property Get TimeStampStr() As String
    Dim d As Date: d = m_TimeStamp
    TimeStampStr = Year(d) & "-" & Dec2(Month(d)) & "-" & Dec2(Day(d)) & "T" & Dec2(Hour(d)) & ":" & Dec2(Minute(d)) & ":" & Dec2(Second(d))
End Property

'2024-12-15:T15:16:17
Public Property Let TimeStampStr(ByVal Value As String)
Try: On Error Resume Next
    If Left(Value, 1) = "'" Then Value = Mid(Value, 2)
    If Right(Value, 1) = "'" Then Value = Left(Value, Len(Value) - 1)
    Dim y  As Integer:  y = CInt(Left(Value, 4))
    Dim m  As Integer:  m = CInt(Mid(Value, 6, 2))
    Dim d  As Integer:  d = CInt(Mid(Value, 9, 2))
    Dim h  As Integer:  h = CInt(Mid(Value, 13, 2))
    Dim mn As Integer: mn = CInt(Mid(Value, 16, 2))
    Dim s  As Integer:  s = CInt(Mid(Value, 19, 2))
    m_TimeStamp = DateSerial(y, m, d) + TimeSerial(h, mn, s)
End Property
