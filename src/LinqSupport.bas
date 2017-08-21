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

'ソート用ユーザー定義型
Private Type SortElement
    Object As Object
    Value As Variant
End Type

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

'ComparePredicate用
    'ComparePredicateの構成が決まっていないので一旦外に



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


'[オブジェクト用マージソート - Qiita](http://qiita.com/nukie_53/items/88ff2227c20cb2f04344 "オブジェクト用マージソート - Qiita")
'の改造
'引数
    'Objects    ：ソートしたいオブジェクトのVBA.Collection
    'MemberName ：プロパティ（VbGet）やメソッド（VbMethod）の名前。
    'CallType   ：省略可。MemberNameの種類。VbGetもしくはVbMethod。省略時VbGet。
    'Ascending  ：省略可。Trueなら昇順、Falseなら降順。省略時True。

'戻り値
    'ソートされたVBA.Collection

Public Function MergeSort(ByVal Objects As VBA.Collection, _
                           ByVal KeyCallBack As IFunc, _
                           Optional ByVal Ascending As Boolean = True _
                        ) As VBA.Collection
    
    Dim basArray() As SortElement
    ReDim basArray(1 To Objects.Count)

    Dim i&, obj As Object
    For Each obj In Objects
        i = i + 1
        Set basArray(i).Object = obj
        Let basArray(i).Value = KeyCallBack.Exec(obj)
    Next obj

    'コピーを作成。ちゃんと考えれば領域確保だけでも良いかも。
    Dim OutArray() As SortElement
    OutArray = basArray

    'ソート
    Call RecurseMergeSort(basArray, OutArray, 1, Objects.Count, Ascending)

    Erase basArray

    '出力用に入れ直し
    Dim oCol As VBA.Collection
    Set oCol = New VBA.Collection
    For i = 1 To Objects.Count
        oCol.Add OutArray(i).Object
    Next i

    Set MergeSort = oCol

End Function


Private Sub RecurseMergeSort( _
        ByRef InptArray() As SortElement, _
        ByRef OutArray() As SortElement, _
        ByVal Start As Long, _
        ByVal Length As Long, _
        ByVal Ascending As Boolean)

    Dim halfLen As Long
    halfLen = VBA.CLng(Length / 2)

    '前半のソート
    If halfLen >= 2 Then
        Call RecurseMergeSort(OutArray, InptArray, Start, halfLen, Ascending)
    End If

    '後半のソート
    If Length - halfLen >= 2 Then
        Call RecurseMergeSort(OutArray, InptArray, Start + halfLen, Length - halfLen, Ascending)
    End If

    '前半部分の添え字と最大値
    Dim lwIndex As Long:    lwIndex = Start
    Dim lwMax As Long:      lwMax = Start + halfLen - 1

    '後半部分の添え字と最大値
    Dim upIndex As Long:    upIndex = Start + halfLen
    Dim upMax As Long:      upMax = Start + Length - 1

    '全体の添え字と最大値
    Dim oIndex As Long:     oIndex = Start
    Dim oMax As Long:       oMax = Start + Length - 1

    Dim leftIndex As Long   '余り用

    Dim flg As Boolean

    For oIndex = Start To oMax Step 1
        '値が同じなら順番維持
        flg = (InptArray(lwIndex).Value = InptArray(upIndex).Value)

        '値が同じじゃない場合、再判定
        If Not flg Then flg = (Ascending = (InptArray(lwIndex).Value < InptArray(upIndex).Value))

        If flg Then
            OutArray(oIndex) = InptArray(lwIndex)
            If lwIndex = lwMax Then
                leftIndex = upIndex
                Exit For
            Else
                lwIndex = lwIndex + 1
            End If
        Else
            OutArray(oIndex) = InptArray(upIndex)
            If upIndex = upMax Then
                leftIndex = lwIndex
                Exit For
            Else
                upIndex = upIndex + 1
            End If
        End If
    Next oIndex

    'Next oIndexを飛ばした分インクリメントする
    For oIndex = oIndex + 1 To oMax Step 1
        OutArray(oIndex) = InptArray(leftIndex)
        leftIndex = leftIndex + 1
    Next oIndex

End Sub


