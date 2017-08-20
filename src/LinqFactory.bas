Attribute VB_Name = "LinqFactory"
Option Explicit

Public Function FCallBy( _
             ByVal iMemeberName As String, _
             ByVal iCallType As VBA.VbCallType, _
        ParamArray iArgs() As Variant _
    ) As CallByFunc
    
    'JPN:ParamArrayÇÕéQè∆ìnÇµÇ≈Ç´Ç»Ç¢ÇΩÇﬂ
    'ParamArray can't use byref argument.
    Dim copyArgs() As Variant
    copyArgs = iArgs
    
    With New CallByFunc
        Set FCallBy = .Init(iMemeberName, iCallType, copyArgs)
    End With 'New CallByFunc
    
End Function

Public Function FCompOp( _
        ByVal iOperator As CompareOperators, _
        ByVal iExpresion As Variant _
    ) As CompareOperator
    
    With New CompareOperator
        Set FCompOp = .Init(iOperator, iExpresion)
    End With 'New CompareOperator
End Function
