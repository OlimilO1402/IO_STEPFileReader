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
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   2400
      Width           =   10455
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
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   720
      Width           =   10455
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
    Dim obj As StepObject
    Dim ser As StepSerializer
    Dim str As StreamStr
    Dim Reader As StepReader
    
    Set obj = MNew.StepObject(1, "COLOR")
    obj.Arguments.Add MNew.StepToken(tkt_NumericInt, 1)
    obj.Arguments.Add MNew.StepToken(tkt_NumericInt, 0)
    obj.Arguments.Add MNew.StepToken(tkt_NumericInt, 0)
    obj.Arguments.Add MNew.StepToken(tkt_NumericInt, 0)
    
    Set ser = New StepSerializer
    obj.Serialize ser
    Text1.Text = ser.ToStr
    
End Sub

Private Sub Command2_Click()
    Set str = MNew.StreamStr(Text1.Text)
    Set Reader = MNew.StepReader(str)
    
    Set obj = Reader.NextStepObject
    obj.Serialize ser
    Text2.Text = ser.ToStr
    
End Sub

