VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparePredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'演算子による比較用クラス
    '大分汚いのできれいにしたいけれど…

Implements IPredicate
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

Private useFunc_    As IFunc
Private operator_   As CompareOperators
Private expression_ As Variant

Private regex_ As Object 'As VBScript_RegExp_55.RegExp

Public Property Get UseFunc() As IFunc:                 Set UseFunc = useFunc_:         End Property
Public Property Get Operator() As CompareOperators:     Let Operator = operator_:       End Property
Public Property Get Expression() As Variant:            Let Expression = expression_:   End Property

Friend Function Init( _
        ByVal iUseFunc As IFunc, _
        ByVal iOperator As CompareOperators, _
        ByVal iExpresion As Variant _
    ) As ComparePredicate
'Attribute Init.VB_UserMemId = 0
    
    If Me Is ComparePredicate Then
        With New ComparePredicate
            Set Init = .Init(iUseFunc, iOperator, iExpresion)
        End With 'New ComparePredicate
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
            Set regex_ = NewRegExp(iExpresion, iIgnoreCase:=True)
            expression_ = iExpresion
            
        Case opcMatch, opcNotMatch
            Set regex_ = NewRegExp(iExpresion, iIgnoreCase:=False)
            expression_ = iExpresion
            
        Case opIs, opIsNot
            If Not VBA.IsObject(iExpresion) Then ThrowLINQ ArgumentException
            Set expression_ = iExpresion
            
        Case Else
            ThrowLINQ ArgumentException
            
    End Select
    
    Set useFunc_ = iUseFunc
    Let operator_ = iOperator
    Set Init = Me
    
End Function

Public Function AsIPredicate() As IPredicate
    Set AsIPredicate = Me
End Function

Private Sub Class_Terminate()
    Set regex_ = Nothing
    Set useFunc_ = Nothing
End Sub

Private Function IPredicate_Test(iElement As Variant) As Boolean
    Dim buf As Variant
    If Me.UseFunc Is Nothing Then
        AssignVal buf, iElement
    Else
        AssignVal buf, Me.UseFunc.Exec(iElement)
    End If
    
    Dim rslt As Boolean
    
    Select Case operator_
        Case opEq, opNe
            rslt = ((buf = expression_) = (operator_ = opEq))
            
        Case opLt, opLe
            rslt = (buf < expression_)
            If (rslt = False) And (operator_ = opLe) Then _
                rslt = (buf = expression_)
            
        Case opGt, opGe
            rslt = (buf > expression_)
            If (rslt = False) And (operator_ = opGe) Then _
                rslt = (buf = expression_)
            
        Case opLike, opNotLike
            rslt = ((buf Like expression_) = (operator_ = opLike))
            
        Case opiMatch, opiNotMatch
            rslt = ((regex_.Test(buf)) = (operator_ = opiMatch))
            
        Case opcMatch, opcNotMatch
            rslt = ((regex_.Test(buf)) = (operator_ = opcNotMatch))
            
        Case opIs, opIsNot
            rslt = ((buf Is expression_) = (operator_ = opIsNot))
            
    End Select
    
    Let IPredicate_Test = rslt

End Function

Private Property Get IPredicate_UseFunc() As IFunc
    Set IPredicate_UseFunc = Me.UseFunc
End Property
