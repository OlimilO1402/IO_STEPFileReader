VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StepTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum StepStates
    sts_Error = -1
    sts_Initial = 0
    sts_HeaderInitial
    sts_DataInitial              'reset the state to DataInitial after every object-line resp after ";"
    sts_Hash                     '#    at beginning of line
    
    sts_NumericPointOrEnumOrBool '. could be followed by a NumericFloatDigit, or by a EnumConstIdentifier or by a Boolean
    sts_NumericSign0             '(+|-)? '? = 0 oder 1
    sts_NumericDigits            '(0-9)* '* = 0 oder viele
    sts_NumericExpoChar          '(e|E)
    sts_NumericExpoSign          'eE(+|-)
    sts_NumericExpoDigits        'eE(+|-)(0-9)
    sts_NumericFloatPoint        '(.)?
    sts_NumericFloatDigits       '.(0-9)
    sts_NumericFloatExpoChar     'eE
    sts_NumericFloatExpoSign     'eE+-
    sts_NumericFloatExpoDigit    '0-9
    
    sts_HashEqualObject          '#123= at beginning of line
    
    sts_EnumConstIdentOrBool     '.T, .t, .F, .f
    sts_IdentifierEnumConst      '.EnumConstIndent.
    sts_IdentifierClass          '
    sts_Boolean                  '.F., .f., .t., .T.
    sts_String0                  ' "'"
    sts_String1                  ' "'"
    sts_HashDigits               '#123 at beginning of line
    sts_HashObject               '#    inside object parameters
    sts_HashObjectDigits         '#456 inside object parameters
    sts_OpenBracket              ' "("
    sts_ListSeparator            ' ","
    sts_CloseBracket             ' ")"
    sts_Semicolon                ' ";"
    
    sts_CommentStart0            ' "/"
    sts_CommentStart1            ' "/*"
    sts_CommentAnyString         ' "/* blablabla"
    sts_CommentEnd0              ' "/* blablabla*"
    sts_CommentEnd1              ' "/* blablabla */"
End Enum
Private m_Str As StreamStr

Friend Sub New_(aStream As StreamStr)
    Set m_Str = aStream
End Sub

'q =  -1:  ErrorState
'q =   0:  InitialState
'q =   1:  "#" Hash
'q =
'q = 100:  NumericInt
'q = 150:  NumericFlt
'q = 200:  String
'q = 300:  Inside Arg-List
'q = 400:  #HashObject
Public Function NextToken() As StepToken
    Dim q0 As StepStates, q1 As StepStates
    Dim c0 As String, c1 As String
    Dim i0 As Long: i0 = m_Str.Index
    'Dim i1 As Long
    Do Until m_Str.EOF
        c0 = m_Str.NextChar
        Select Case c0
        Case "#" ' AscW(c0) = 35
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_Hash
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" ' 48 - 57
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_NumericDigits
            Case sts_IdentifierEnumConst:                         q1 = StepStates.sts_IdentifierEnumConst
            Case sts_IdentifierClass:                             q1 = StepStates.sts_IdentifierClass
            Case sts_NumericSign0:                                q1 = StepStates.sts_NumericDigits
            Case sts_NumericDigits:                               q1 = StepStates.sts_NumericDigits
            Case sts_NumericExpoChar:                             q1 = StepStates.sts_NumericDigits
            Case sts_NumericExpoSign:                             q1 = StepStates.sts_NumericDigits
            Case sts_NumericPointOrEnumOrBool:                    q1 = StepStates.sts_NumericFloatDigits
            Case sts_NumericFloatPoint:                           q1 = StepStates.sts_NumericFloatDigits
            Case sts_NumericFloatDigits:                          q1 = StepStates.sts_NumericFloatDigits
            Case sts_NumericFloatExpoChar:                        q1 = StepStates.sts_NumericFloatExpoDigit
            Case sts_Hash:                                        q1 = StepStates.sts_HashDigits
            Case sts_HashDigits:                                  q1 = StepStates.sts_HashDigits
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "=" ' 61
            Select Case q0
            Case sts_HashDigits:                                  Set NextToken = MNew.StepToken(tkt_ObjectHash, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "A" To "D", "G" To "S", "U" To "Z", _
             "a" To "d", "g" To "s", "u" To "z"     ' A 65-68,71-83,85-90, a 97-100,103-115,117-122    ohne E,e, F,f, T,t entweder hier explizit ausklammern oder erst nach den anderen
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_IdentifierClass
            Case sts_HashEqualObject
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "E", "e"  '69, 101
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_IdentifierClass
            Case sts_NumericSign0:                                q1 = StepStates.sts_NumericExpoChar
            Case sts_NumericFloatPoint:                           q1 = StepStates.sts_NumericFloatExpoChar
            Case sts_NumericFloatDigits:                          q1 = StepStates.sts_NumericFloatExpoChar
            Case sts_HashEqualObject:                             q1 = StepStates.sts_ClassIdentifier
            Case sts_ClassIdentifier:                             q1 = StepStates.sts_ClassIdentifier
            Case sts_NumericPointOrEnumOrBool:                    q1 = StepStates.sts_EnumConst
            Case sts_EnumConst:                                   q1 = StepStates.sts_EnumConst
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "F", "f"  '70, 102
            Select Case q0
            Case sts_NumericPointOrEnumOrBool:                    q1 = StepStates.sts_EnumConstOrBoolean
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "T", "t"  '84, 116
            Select Case q0
            Case sts_NumericPointOrEnumOrBool:                    q1 = StepStates.sts_EnumConstOrBoolean
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "_"       '95
            Select Case q0
            Case sts_ClassIdentifier:                             q1 = StepStates.sts_ClassIdentifier
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "-"       '45
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_NumericSign0
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "."       '46  'NumericFlt oder Enum oder Boolean
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_NumericPointOrEnumOrBool
            Case sts_NumericSign0, sts_NumericDigits:             q1 = StepStates.sts_NumericFloatPoint
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "+"       '43
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_NumericSign0
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
        Case "'"       '39     'evtl String
            Select Case q0
            Case sts_Initial, sts_HeaderInitial, sts_DataInitial: q1 = StepStates.sts_NumericExpoChar
            Case sts_String0:                                     q1 = StepStates.sts_String1
            Case sts_String1:                                     q1 = StepStates.sts_String0
            End Select
        Case ","       '44
            ' OK dies heißt jetzt für viele Tokens wir sind zuende mit wasimmer wir bis hierher gelesen haben
            ' wir erzeugen also das Token, und verlassen die Funktion, wir müssen nicht q0,q1 und nicht i0,i1 setzen
            Select Case q0
            Case StepStates.sts_NumericDigits:                    Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericInt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_NumericExpoDigits:                Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericFlt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_NumericFloatDigits:               Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericFlt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_NumericFloatExpoDigit:            Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericFlt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_NumericFloatPoint:                Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericFlt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_NumericFloatPoint:                Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericFlt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case sts_String1:                                     Set NextToken = MNew.StepToken(EStepTokenType.tkt_NumericInt, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            'Case Else:                                            q1 = StepStates.sts_Error
            End Select
            'i0 = m_Str.Index
        Case "("
            Select Case q0
            Case StepStates.sts_IdentifierClass:                  Set NextToken = MNew.StepToken(EStepTokenType.tkt_Identifier, m_Str.ExtractFromTo(i0, m_Str.Index - 1)): Exit Function
            Case StepStates.sts_String0:                          q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString      'bei Comment alles durchlassen
            Case Else:                                            q1 = StepStates.sts_Error
            End Select
            ': Set NextToken = MNew.StepToken(EStepTokenType.tkt_OpenBracket, c0)
        Case ")"
            'this is the end of the
            ': Set NextToken = MNew.StepToken(EStepTokenType.tkt_CloseBracket, c0)
        Case ";"
            Select Case q0
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString      'bei Comment alles durchlassen
            Case Else:                                            q1 = sts_DataInitial:       i0 = 1
            'this is the end of the class definition so reset
            'should we spit out a token for that?
            ': Set NextToken = MNew.StepToken(EStepTokenType.tkt_Semicolon, c0)
            End Select
        Case " ", vbCr, vbLf, vbCrLf
            Select Case q0
            Case sts_String0:                                     q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString      'bei Comment alles durchlassen
            Case Else:                                            q1 = sts_DataInitial:       i0 = 1
            End Select
        Case "/"
            Select Case q0
            Case StepStates.sts_Initial:                          q1 = StepStates.sts_CommentStart0
            Case StepStates.sts_CommentEnd0:                      q1 = StepStates.sts_CommentEnd1
            Case StepStates.sts_String0:                          q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString      'bei Comment alles durchlassen
            Case Else:                                            q1 = sts_DataInitial:       i0 = 1
            End Select
        Case "*"
            Select Case q0
            Case StepStates.sts_CommentStart0:                    q1 = StepStates.sts_CommentStart1
            Case StepStates.sts_String0:                          q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString      'bei Comment alles durchlassen
            Case Else:                                            q1 = sts_DataInitial:       i0 = 1
            End Select
        Case Else 'any other character could be contained in a string or a comment
            Select Case q0
            Case StepStates.sts_String0:                          q1 = StepStates.sts_String0               'bei String0 alles durchlassen
            Case StepStates.sts_CommentAnyString:                 q1 = StepStates.sts_CommentAnyString
            Case Else:                                            q1 = StepStates.sts_Error:  i0 = 1
            End Select
        End Select
        q0 = q1
        c1 = c0
    Loop
End Function

Public Function IsEOF() As Boolean
    IsEOF = m_Str.EOF
End Function
