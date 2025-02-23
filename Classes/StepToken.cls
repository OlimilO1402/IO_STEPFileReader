VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EStepTokenType
    tkt_ObjectHash         ' start of line "#" NumericInt
    tkt_Identifier         ' Identifier contains letters (A-Z)*,(a-z)*,(0-9)*,(_),(-) must start with alphabetic letter
    tkt_NumericInt         ' ("+"|"-")?("0"|"1"|"2"|"3"|"4"|"5"|"6"|"7"|"8"|"9")+
    tkt_NumericFlt         ' ("+"|"-")?("0"|"1"|"2"|"3"|"4"|"5"|"6"|"7"|"8"|"9")*"."("0"|"1"|"2"|"3"|"4"|"5"|"6"|"7"|"8"|"9")+("e"("+"|"-")?("0"|"1"|"2"|"3"|"4"|"5"|"6"|"7"|"8"|"9")*
    tkt_String             ' everything between the first "'" and the second "'" if string contains character "'" then only escaped with double character "'"
    tkt_EmptyOrDefault     ' "$"
    tkt_Boolean            ' "."("F"|"T")"."
    tkt_EnumIdentifier     ' "."EnumConstIdentifier"."
    tkt_nTupelList         ' "("(StpToken?(","StpToken)*")"
    tkt_BracketOpen        ' "("
    tkt_BracketClose       ' ")"
    tkt_Comment            ' /* any text as comment */
End Enum

Public TokenType As EStepTokenType
Public Value     As Variant       '

Friend Sub New_(aTokenType As EStepTokenType, aValue)
    TokenType = aTokenType
    If VarType(aValue) = vbObject Then Set Value = aValue Else Value = aValue
End Sub
''copy to MNew:
'Public Function StepToken(aTokenType As EStepTokenType, aValue) As StepToken
'    Set StepToken = New StepToken: StepToken.New_ aTokenType, aValue
'End Function

Public Function Serialize(Serializer As StepSerializer) As String
    Serializer.SerializeStepToken Me
End Function
