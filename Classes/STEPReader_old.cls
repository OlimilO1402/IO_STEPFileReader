VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STEPReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN   As PathFileName
Private m_Stack As Stack

Public Description         As String
Public ImplementLevel      As String
Public Name                As String
Public TimeStamp           As String
Public Author              As String
Public Organisation        As String
Public PreprocessorVersion As String
Public OriginatingSystem   As String
Public Authorisation       As String
Public FileScheme          As String

Private Sub Class_Initialize()
    Set m_Stack = MNew.List(EDataType.vbString)
End Sub

Friend Sub New_(aPathFileName As PathFileName)
    Set m_PFN = aPathFileName
    Read
End Sub
'Copy to MNew.bas:
'Public Function STEPReader(aPathFileName As PathFileName) As STEPReader
'    Set STEPReader = New STEPReader: STEPReader.New_ aPathFileName
'End Function

Public Function Read() As Boolean
Try: On Error GoTo Catch
    Dim line As String, state As Long
    m_PFN.OpenFile EFileMode.FileMode_Input
    Do Until m_PFN.IsEOF
        line = ReadLine
        Select Case UCase(line)
        Case "ISO-10303-21;"
            If state = 0 Then
                state = 1
            Else
                MsgBox "Second header?"
            End If
        Case "HEADER;"
            If state = 1 Then
                ReadHeader
                state = 2
            Else
                MsgBox "Error during Reading"
            End If
        Case "ENDSEC;"
            If state = 0 Then
                MsgBox "header-section and data-section not yet read."
            Else
                state = 1
            End If
        Case "DATA;"
            ReadData
        Case "END-ISO-10303-21;"
            state = -1
        End Select
    Loop
    Read = True
    Exit Function
Catch:
End Function

Private Function ReadLine() As String
    Dim line As String: line = Trim(m_PFN.ReadLine)
    Do Until Right(line, 1) = ";"
        line = line & Trim(m_PFN.ReadLine)
    Loop
    Debug.Print line
    ReadLine = DeleteComments(line)
End Function

Private Function DeleteComments(line As String) As String
    If Len(line) = 0 Then Exit Function
    Dim pos0 As Long: pos0 = InStr(1, line, "/*")
    If pos0 <= 0 Then DeleteComments = line: Exit Function
    Dim pos1 As Long: pos1 = InStr(pos0, line, "*/")
    If pos1 <= 0 Then DeleteComments = line: Exit Function
    DeleteComments = Left(line, pos0 - 1) & Mid(line, pos1 + 2)
    DeleteComments = DeleteComments(DeleteComments)
End Function

Private Function ReadHeader() As Boolean
Try: On Error GoTo Catch
    Dim line As String, uline As String, s As String
    Do Until m_PFN.IsEOF
        line = ReadLine
        uline = UCase(line)
        If MString.StartsWith(uline, "FILE_DESCRIPTION") Then
            line = Trim(Mid(line, 16 + 2, Len(line) - (16 + 1 + 2)))
            If Not ReadFileDescription(line) Then
                MsgBox "Error during reading file_descriptiopn!"
            End If
        ElseIf MString.StartsWith(uline, "FILE_NAME") Then
            line = Mid(line, 9 + 1, Len(line) - (9 + 1 + 2))
            If Not ReadFileName(line) Then
                MsgBox "Error during reading file_name!"
            End If
        ElseIf MString.StartsWith(uline, "FILE_SCHEMA") Then
            line = Mid(line, 11 + 1, Len(line) - (11 + 1 + 2))
            If Not ReadFileScheme(line) Then
                MsgBox "Error during reading file_name!"
            End If
        End If
    Loop
    ReadHeader = True
    Exit Function
Catch:
End Function

Private Function ReadFileDescription(line As String) As Boolean
Try: On Error GoTo Catch
    Dim s As String, sa() As String: sa = Split(line, ",")
    Dim i As Long, u As Long: u = UBound(sa)
    If i <= u Then
        s = Trim(sa(i))
        If Left(s, 2) = "('" Then s = Mid(s, 3)
        If Right(s, 2) = "')" Then s = Mid(s, 1, Len(s) - 2)
        Me.Description = s
    End If
    i = i + 1
    If i <= u Then
        s = Trim(sa(i))
        If Left(s, 1) = "'" Then s = Mid(s, 2)
        If Right(s, 1) = "'" Then s = Mid(s, 1, Len(s) - 1)
        Me.ImplementLevel = s
    End If
    ReadFileDescription = True
    Exit Function
Catch:
End Function

Private Function ReadFileName(line As String) As Boolean
Try: On Error GoTo Catch
    Dim s As String, sa() As String: sa = Split(line, ",")
    Dim i As Long, u As Long: u = UBound(sa)
    If i <= u Then
        s = Trim(sa(i))
        If Left(s, 1) = "'" Then s = Mid(s, 2)
        If Right(s, 1) = "'" Then s = Mid(s, 1, Len(s) - 3)
        Me.Name = s
    End If
    i = i + 1
    If i <= u Then
        s = Trim(sa(i))
        If Left(s, 1) = "'" Then s = Mid(s, 2)
        If Right(s, 1) = "'" Then s = Mid(s, 1, Len(s) - 3)
        Me.TimeStamp = s
    End If
    i = i + 1
    If i <= u Then
        s = Trim(sa(i))
        If Left(s, 1) = "'" Then s = Mid(s, 2)
        If Right(s, 1) = "'" Then s = Mid(s, 1, Len(s) - 3)
        Me.Author = s
    End If
    ReadFileName = True
    Exit Function
Catch:
End Function
Private Function ReadString(line As String) As String
    
End Function
Private Function ReadFileScheme() As Boolean
    
End Function

Private Function ReadData() As Boolean
    '
End Function

Public Function Header_ToStr() As String
    Dim s As String: s = "HEADER;" & vbCrLf
    s = s & "FILE_DESCRIPTION(" & "('" & Me.Description & "'),'" & Me.ImplementLevel & "');" & vbCrLf
    s = s & "FILE_NAME(" & "'" & Me.Name & "','" & Me.TimeStamp & "','" & Me.Author & "','" & Me.Organisation & "','" & Me.PreprocessorVersion & "','" & Me.OriginatingSystem & "','" & Me.Authorisation & "');" & vbCrLf
    s = s & "FILE_SCHEMA(" & "('" & Me.FileScheme & "'));" & vbCrLf
    s = s & "ENDSEC;" & vbCrLf
    Header_ToStr = s
End Function

Public Function Data_ToStr() As String
    Dim s As String: s = "DATA;" & vbCrLf
    s = s & "ENDSEC;" & vbCrLf
    Data_ToStr = s
End Function

Public Function ToStr() As String
    Dim s As String: s = "ISO-10300-21;" & vbCrLf
    s = s & Header_ToStr
    s = s & Data_ToStr
    s = s & "END-ISO-10300-21;"
    ToStr = s
End Function
