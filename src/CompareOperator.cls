VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompareOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'スコープ適当。読み取り専用プロパティにする予定
Public Operator As CompareOperators
Public Expression As Variant

Private clsRegExp As VBScript_RegExp_55.RegExp

    
Public Function Init( _
        ByVal iOperator As CompareOperators, _
        ByVal iExpresion As Variant _
    ) As CompareOperator
Attribute Init.VB_UserMemId = 0
'Attribute Init.VB_UserMemId = 0


    If Me Is CompareOperator Then
        With New CompareOperator
            Set Init = .Init(iOperator, iExpresion)
        End With
        Exit Function
    End If
    
    Select Case iOperator
        Case opEq, opNe
            Expression = iExpresion
            
        Case opLt, opLe
            Expression = iExpresion
            
        Case opGt, opGe
            Expression = iExpresion
            
        Case opLike, opNotLike
            Expression = iExpresion
            
        Case opiMatch, opiNotMatch
            Set clsRegExp = NewRegExp(iExpresion, iIgnoreCase:=True)
            Expression = iExpresion
            
        Case opcMatch, opcNotMatch
            Set clsRegExp = NewRegExp(iExpresion, iIgnoreCase:=False)
            Expression = iExpresion
            
        Case opIs, opIsNot
            If Not VBA.IsObject(iExpresion) Then ThrowLINQ ArgumentException
            Set Expression = iExpresion
            
        Case Else
            ThrowLINQ ArgumentException
            
    End Select
    
    
    Let Me.Operator = iOperator
    Set Init = Me
    
End Function

Private Sub Class_Terminate()
    Set clsRegExp = Nothing
End Sub

Private Function IFunc_Exec(iElement As Variant) As Variant
    Dim rslt As Boolean
    
    Select Case Me.Operator
        Case opEq, opNe
            rslt = ((iElement = Expression) = (Me.Operator = opEq))
            
        Case opLt, opLe
            rslt = (iElement < Expression)
            If (rslt = False) And (Me.Operator = opLe) Then _
                rslt = (iElement = Expression)
            
        Case opGt, opGe
            rslt = (iElement > Expression)
            If (rslt = False) And (Me.Operator = opGe) Then _
                rslt = (iElement = Expression)
            
        Case opLike, opNotLike
            rslt = ((iElement Like Expression) = (Me.Operator = opLike))
            
        Case opiMatch, opiNotMatch
            rslt = ((clsRegExp.Test(iElement)) = (Me.Operator = opiMatch))
            
        Case opcMatch, opcNotMatch
            rslt = ((clsRegExp.Test(iElement)) = (Me.Operator = opcNotMatch))
            
        Case opIs, opIsNot
            rslt = ((iElement Is Expression) = (Me.Operator = opIsNot))
            
    End Select
    
    Let IFunc_Exec = rslt
    
End Function
