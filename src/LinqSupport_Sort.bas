Attribute VB_Name = "LinqSupport_Sort"
'[オブジェクト用マージソート - Qiita](http://qiita.com/nukie_53/items/88ff2227c20cb2f04344 "オブジェクト用マージソート - Qiita")
'基本は上記コードの改造
    '一部意図しない動作があったため要修正

Option Explicit

'ソート用ユーザー定義型
Private Type SortElement
    Object As Object
    Value As Variant
End Type

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

