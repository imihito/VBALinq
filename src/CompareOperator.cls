VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompareOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'演算子による比較用クラス
    '大分汚いのできれいにしたいけれど…

Implements IFunc
Option Explicit


'比較演算子指定用
    '名前はPowerShellの比較演算子をイメージ
Public Enum CompareOperators
    opEq        '=
    opNe        '<>
    opLt        '<
    opLe        '<=
    opGt        '>
    opGe        '>=
    opLike      ' Like
    opNotLike   '( Like ) = False
    opiMatch    'RegExp.Test( ) = True  / IgnoreCase = True
    opiNotMatch 'RegExp.Test( ) = False / IgnoreCase = True
    opcMatch    'RegExp.Test( ) = True  / IgnoreCase = False
    opcNotMatch 'RegExp.Test( ) = False / IgnoreCase = False
    opIs        ' Is = True
    opIsNot     ' Is = False
End Enum

Private operator_ As CompareOperators
Private expression_ As Variant

Private clsRegExp As VBScript_RegExp_55.RegExp

Public Property Get Operator() As CompareOperators:     Let Operator = operator_:       End Property
Public Property Get Expression() As Variant:            Let Expression = expression_:   End Property

Friend Function Init( _
        ByVal iOperator As CompareOperators, _
        ByVal iExpresion As Variant _
    ) As CompareOperator
'Attribute Init.VB_UserMemId = 0
    
    If Me Is CompareOperator Then
        With New CompareOperator
            Set Init = .Init(iOperator, iExpresion)
        End With 'New CompareOperator
        Exit Function
    End If
    
    Select Case iOperator
        Case opEq, opNe
            expression_ = iExpresion
            
        Case opLt, opLe
            expression_ = iExpresion
            
        Case opGt, opGe
            expression_ = iExpresion
            
        Case opLike, opNotLike
            expression_ = iExpresion
            
        Case opiMatch, opiNotMatch
            Set clsRegExp = NewRegExp(iExpresion, iIgnoreCase:=True)
            expression_ = iExpresion
            
        Case opcMatch, opcNotMatch
            Set clsRegExp = NewRegExp(iExpresion, iIgnoreCase:=False)
            expression_ = iExpresion
            
        Case opIs, opIsNot
            If Not VBA.IsObject(iExpresion) Then ThrowLINQ ArgumentException
            Set expression_ = iExpresion
            
        Case Else
            ThrowLINQ ArgumentException
            
    End Select
    
    
    Let operator_ = iOperator
    Set Init = Me
    
End Function

Private Sub Class_Terminate()
    Set clsRegExp = Nothing
End Sub

Private Function IFunc_Exec(iElement As Variant) As Variant
    Dim rslt As Boolean
    
    Select Case operator_
        Case opEq, opNe
            rslt = ((iElement = expression_) = (operator_ = opEq))
            
        Case opLt, opLe
            rslt = (iElement < expression_)
            If (rslt = False) And (operator_ = opLe) Then _
                rslt = (iElement = expression_)
            
        Case opGt, opGe
            rslt = (iElement > expression_)
            If (rslt = False) And (operator_ = opGe) Then _
                rslt = (iElement = expression_)
            
        Case opLike, opNotLike
            rslt = ((iElement Like expression_) = (operator_ = opLike))
            
        Case opiMatch, opiNotMatch
            rslt = ((clsRegExp.Test(iElement)) = (operator_ = opiMatch))
            
        Case opcMatch, opcNotMatch
            rslt = ((clsRegExp.Test(iElement)) = (operator_ = opcNotMatch))
            
        Case opIs, opIsNot
            rslt = ((iElement Is expression_) = (operator_ = opIsNot))
            
    End Select
    
    Let IFunc_Exec = rslt
    
End Function
