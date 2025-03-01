VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14370
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14370
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   11280
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   4080
      Width           =   12975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   720
      Width           =   12975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As PathFileName
Private m_stp As StepReader

Private Sub Command1_Click()
'    Set m_PFN = MNew.PathFileName(App.Path & "\test.stp")
'    If Not m_PFN.Exists Then
'        MsgBox "File not found:" & vbCrLf & m_PFN.Value
'        Exit Sub
'    End If
'    Set m_stp = MNew.StepReader(m_PFN)
'    Text1.Text = m_stp.ToStr
    
    
    Dim Obj As StepObject
    Dim ser As StepSerializer
    Dim str As StreamStr
    Dim Reader As StepReader
    
    Set Obj = MNew.StepObject(1, "COLOR")
    Obj.Tokens.Add MNew.StepToken(tkt_NumericInt, 1)
    Obj.Tokens.Add MNew.StepToken(tkt_NumericInt, 0)
    Obj.Tokens.Add MNew.StepToken(tkt_NumericInt, 0)
    Obj.Tokens.Add MNew.StepToken(tkt_NumericInt, 0)
    
    Set ser = New StepSerializer
    Obj.Serialize ser
    Text1.Text = ser.ToStr
    
End Sub

Private Sub Command2_Click()
    Dim str As StreamStr: Set str = MNew.StreamStr(Text1.Text)
    Dim Reader As StepReader
    
    Set Reader = MNew.StepReader(str)
    
    Set Obj = Reader.NextStepObject
    Obj.Serialize ser
    Text2.Text = ser.ToStr
    
End Sub

Private Sub Command3_Click()
    
    Dim sth As StepHeader:    Set sth = MNew.StepHeader("Dateibeschreibung", "FImpl 2.1", "C:\test.ifc", "", "OlimilO1402", "MBO-Ing.com", "Icx03")
    
    Dim std As StepData:      Set std = New StepData
    Dim sdoc As StepDocument: Set sdoc = MNew.StepDocument(sth, std)
    
    Dim Color As StepObject:  Set Color = MNew.StepObject(0, "ICXCOLOR")
    With Color
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericInt, 3)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericInt, 255)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericInt, 0)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericInt, 0)
    End With
    
    Dim Layer As StepObject: Set Layer = MNew.StepObject(0, "ICXLAYER")
    With Layer
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericInt, 1)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_String, "Standardlayer1")
    End With
    
    Dim LinTyp As StepObject: Set LinTyp = MNew.StepObject(0, "ICXLINETYPE")
    With LinTyp
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_EnumIdentifier, "SOLID")
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_Boolean, True)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_EmptyOrDefault, Nothing)
    End With
    
    Dim Point0 As StepObject: Set Point0 = MNew.StepObject(0, "ICXPOINT")
    With Point0
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 0#)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 0#)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 0#)
    End With
    
    Dim Point1 As StepObject: Set Point1 = MNew.StepObject(0, "ICXPOINT")
    With Point1
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 1#)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 0#)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_NumericFlt, 0#)
    End With
    
    'Dim Points As StepTokens: Set Points = New StepTokens
    Dim Points As nTupel: Set Points = MNew.nTupel(vbObject, 2)
    Points.Add MNew.StepToken(EStepTokenType.tkt_StepObject, Point0)
    Points.Add MNew.StepToken(EStepTokenType.tkt_StepObject, Point1)
    
    Dim Line As StepObject: Set Line = MNew.StepObject(0, "ICXLINE")
    With Line
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_String, "BIMID1")
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_StepObject, Color)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_StepObject, Layer)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_StepObject, LinTyp)
        .Tokens.Add MNew.StepToken(EStepTokenType.tkt_nTupelList, Points)
    End With
    
    Line.AddToStepData std
    
    Dim ser As StepSerializer: Set ser = New StepSerializer
    sdoc.Serialize ser
    Text2.Text = ser.ToStr
End Sub
