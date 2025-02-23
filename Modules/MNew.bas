Attribute VB_Name = "MNew"
Option Explicit

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function StepDocument(aHeader As StepHeader, aData As StepData) As StepDocument
    Set StepDocument = New StepDocument: StepDocument.New_ aHeader, aData
End Function

Public Function StepHeader(ByVal FDescr As String, ByVal FImpl As String, ByVal PFNam As String, DateTimeStamp, ByVal Auth As String, ByVal Organis As String) As StepHeader
    Set StepHeader = New StepHeader: StepHeader.New_ FDescr, FImpl, PFNam, DateTimeStamp, Auth, Organis
End Function

Public Function StepObject(ByVal aHash As Long, ByVal aClassName As String) As StepObject
    Set StepObject = New StepObject: StepObject.New_ aHash, aClassName
End Function

'Public Function StepReader(aPathFileName As PathFileName) As StepReader
'    Set StepReader = New StepReader: StepReader.New_ aPathFileName
'End Function
Public Function StepReader(aStr As StreamStr) As StepReader
    Set StepReader = New StepReader: StepReader.New_ aStr
End Function

Public Function StepToken(aTokenType As EStepTokenType, aValue) As StepToken
    Set StepToken = New StepToken: StepToken.New_ aTokenType, aValue
End Function

Public Function StreamStr(ByVal aValue As String) As StreamStr
    Set StreamStr = New StreamStr: StreamStr.New_ aValue
End Function

Public Function StepTokenizer(aStream As StreamStr) As StepTokenizer
    Set StepTokenizer = New StepTokenizer: StepTokenizer.New_ aStream
End Function

'Class-hierarchiy
'
'Class StepDocument
'    Header As StepHeader
'        Class StepHeader
'
'    Datas  As Collection Of StepData
'
'    Class StepData
'        Objects As List Of StepObject
'            Class StepObject
'                Hash As Long
'                Name As String
'                Tokens As StepTokens
'                    Class StepTokens
'                        Tokens As List Of StepToken
'                            Class StepToken
'                                TokenType As EStepTokenType
'                                Value     As Variant
'    Class StepReader
'        Tokenizer As StepTokenizer
'            Class StepTokenizer
'
'    Class StepSerializer
'        Str As StreamStr
'            Class StreamStr
