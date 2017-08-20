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


Function InsertToHead( _
             ByVal iBaseArray As Variant, _
        ParamArray iInsertElements() As Variant _
    ) As Variant
    
    Const ARRAY_BASE& = 0
    
    If LBound(iBaseArray) <> ARRAY_BASE Then _
        ThrowLINQ IndexOutOfRangeException
    
    Dim insertCnt As Long
    insertCnt = UBound(iInsertElements) + 1
    
    ReDim Preserve iBaseArray(ARRAY_BASE To UBound(iBaseArray) + insertCnt)
    
    Dim i As Long
    For i = UBound(iBaseArray) - insertCnt To ARRAY_BASE Step -1
        AssignVal iBaseArray(i + insertCnt), iBaseArray(i)
    Next i
    
    For i = ARRAY_BASE To insertCnt - 1
        AssignVal iBaseArray(i), iInsertElements(i)
    Next i
    
    Let InsertToHead = iBaseArray
    
End Function

Sub opifhpiuafheuip()
    Dim tmp
    tmp = Array()
    Dim tmp2
    tmp2 = InsertToHead(tmp, "a", 1)
    Set tmp2 = FCallBy("Name", VbGet)
    Stop
End Sub
