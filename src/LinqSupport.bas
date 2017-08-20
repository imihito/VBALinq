Attribute VB_Name = "LinqSupport"
'オレオレLINQ用補助処理

Option Explicit
'Option Private Module

Enum LinqErrNumber
    ArgumentException = 5
    IndexOutOfRangeException = 9
    InvalidOperationException = 17
    NotImplementedException = 32768
End Enum

'代入用
'outVal = setVal
Sub AssignVal(ByRef outVal As Variant, ByRef setVal As Variant)
    If VBA.IsObject(setVal) Then
        Set outVal = setVal
    Else
        Let outVal = setVal
    End If
End Sub

'VBA.Interaction.CallByName拡張
    '引数をParamArrayではなく、Variant型配列で受け取って実行する
    '処理べた書きのため、引数の数に制限有り
    
    '基本はCallByFunc用だけど他にも応用できそうなので一旦外に
Function CallByNameEx(ByVal iObject As Object, ByVal iProcName As String, ByVal iCallType As VBA.VbCallType, ByRef iArgs() As Variant) As Variant
    If VBA.CBool(Not Not iArgs) Then
        Select Case UBound(iArgs)
            Case -1:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType)
            Case 0:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0))
            Case 1:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1))
            Case 2:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2))
            Case 3:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3))
            Case 4:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4))
            Case 5:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5))
            Case 6:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6))
            Case 7:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7))
            Case 8:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8))
            Case 9:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9))
            Case 10:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10))
            Case 11:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10), iArgs(11))
            Case 12:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10), iArgs(11), iArgs(12))
            
            Case Else:  ThrowLINQ ArgumentException
        End Select
    Else
        AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType)
    End If
End Function

'CompareOperator用
    'CompareOperatorの構成が決まっていないので一旦外に
Function NewRegExp( _
                 ByVal iPattern As String, _
        Optional ByVal iGlobal As Boolean, _
        Optional ByVal iIgnoreCase As Boolean, _
        Optional ByVal iMultiLine As Boolean _
    ) As VBScript_RegExp_55.RegExp
    
    Dim tmpRegEx As VBScript_RegExp_55.RegExp
    Set tmpRegEx = VBA.CreateObject("VBScript.RegExp")
    
    With tmpRegEx
        .Pattern = iPattern
        .Global = iGlobal
        .IgnoreCase = iIgnoreCase
        .MultiLine = iMultiLine
    End With    'tmpRegEx
    
    Set NewRegExp = tmpRegEx
    
End Function


Sub ThrowLINQ(ByVal iErrNo As LinqErrNumber)
    Select Case iErrNo
        Case NotImplementedException
            Err.Raise iErrNo, , "機能がまだ実装されていません。" & vbNewLine & "NotImplemented"
        Case Else
            Err.Raise iErrNo
    End Select
End Sub

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


