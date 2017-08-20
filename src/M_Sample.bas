Attribute VB_Name = "M_Sample"
Option Explicit

Sub Sample()
    
    '既定のインスタンスのFromメソッドにForEachできるものを指定してインスタンス
    Dim myLinq As Enumerable
    Set myLinq = Enumerable.From(ThisWorkbook.Worksheets)
    
    'IFuncはデリゲートの代わり
    Dim selectFunc As IFunc
    Set selectFunc = FCallBy("UsedRange", VbGet)
    
    'Worksheetだけにして、UsedRangeを取得
    Dim selectLinq As Enumerable
    Set selectLinq = myLinq.OfType("Worksheet").Select1(selectFunc)
    
    
    Dim nameFunc As CallByFunc
    Set nameFunc = FCallBy("Name", VbGet)
    
    'NameプロパティがLike演算子で"Sheet[0-9]"とマッチするもの
        '「CallByFunc.SetChild」でCallByFuncの結果をもとに次のIFunc呼び出し
    Dim whereLinq As Enumerable
    Set whereLinq = myLinq.Where(nameFunc.SetChild( _
                                FCompOp(opLike, "Sheet[0-9]") _
                            ) _
                        )
    
    '全ブックのワークシート一覧
    Dim selectManyLinq As Enumerable
    Set selectManyLinq = Enumerable.From(Workbooks) _
                            .SelectMany(FCallBy("Worksheets", VbGet))
    
    Dim tWs As Excel.Worksheet
    For Each tWs In selectManyLinq
        Debug.Print tWs.Name
    Next tWs
    
    Stop
    
End Sub



Sub uihwaefiuhpf()
    Debug.Print Enumerable(ThisWorkbook.Sheets).Count(FCallBy("Name", VbGet).SetChild(FCompOp(opLike, "Sheet*")))
End Sub
