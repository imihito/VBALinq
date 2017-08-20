VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private sequence_ As VBA.Collection

Public Function From(ByVal iEnumerable As Variant) As Enumerable
Attribute From.VB_UserMemId = 0
'Attribute From.VB_UserMemId = 0
    If Me Is Enumerable Then
        With New Enumerable
            Set From = .From(iEnumerable)
        End With 'New Enumerable
        Exit Function
    End If
    
    If TypeOf iEnumerable Is VBA.Collection Then
        Set sequence_ = iEnumerable
    Else
        Set sequence_ = New VBA.Collection
        Dim iter As Variant
        For Each iter In iEnumerable
            sequence_.Add iter
        Next iter
    End If
    
    Set From = Me
End Function

'ForEach対応用。ユーザーは使用しない
Public Function GetEnumerator() As stdole.IEnumVARIANT
Attribute GetEnumerator.VB_UserMemId = -4
'Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = sequence_.[_NewEnum]
End Function


''他の自作クラスに依存しない処理群

'他の要素群を末尾に追加します。重複判定は行いません。
Public Function Concat(ByVal iEnumerable As Variant) As Enumerable
    Dim newCol As VBA.Collection
    Set newCol = Me.ToCollection
    
    Dim iter As Variant
    For Each iter In iEnumerable
        newCol.Add iter
    Next iter
    
    Set Concat = Enumerable.From(newCol)
End Function

'指定したインデックスの要素を取得。1始まり。
Public Function ElementAt(ByVal iIndex As Long) As Variant
    AssignVal ElementAt, sequence_.Item(iIndex)
End Function

'TypeName関数の結果が指定された名前に一致する要素のみにフィルターします。
    'Like演算子で判定を行うため、ワイルドカードなどが使用できます。
    'TypeOf 演算子のような厳密な判定は行いません。
Public Function OfType(ByVal iTypeName As String) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        If VBA.TypeName(iter) Like iTypeName Then _
            newCol.Add iter
        
    Next iter
    
    Set OfType = Enumerable.From(newCol)
End Function

'要素の順番を反転させます。
Public Function Reverse() As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant, cnt As Long
    For Each iter In sequence_
        cnt = cnt + 1
        If cnt = 1 Then
            newCol.Add iter
        Else
            newCol.Add iter, Before:=1
        End If
    Next iter
    
    Set Reverse = Enumerable.From(newCol)
    
End Function

'指定された数の要素を除いた要素を返します。
Public Function Skip(ByVal iCount As Long) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant, cnt As Long
    For Each iter In sequence_
        cnt = cnt + 1
        If cnt > iCount Then newCol.Add iter
            
    Next iter
    
    Set Skip = Enumerable.From(newCol)
    
End Function

'先頭から指定された数の要素を返します。
Public Function Take(ByVal iCount As Long) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant, cnt As Long
    For Each iter In sequence_
        cnt = cnt + 1
        If cnt > iCount Then Exit For
        
        newCol.Add iter
        
    Next iter
    
    Set Take = Enumerable.From(newCol)
    
End Function

Public Function ToArray( _
        Optional ByVal iType As VBA.VbVarType = vbVariant, _
        Optional ByVal iLBound As Long = 0 _
    ) As Variant
    
    ThrowLINQ NotImplementedException
    
    Select Case iType
        Case VbVarType.vbVariant
    End Select
    
End Function

'VBA.Collectionに変換します。
    '遅延評価をしておらず、そのままForEachできるため、
    'Enumerableのまま連れ回しても特に問題は無いはずです。
Public Function ToCollection() As VBA.Collection
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        newCol.Add iter
    Next iter
    
    Set ToCollection = newCol
    
End Function


''IFuncに依存する処理群

Public Function OrderBy(ByVal func As IFunc) As Enumerable
    Set OrderBy = Enumerable.From(MergeSort(sequence_, func, True))
End Function

Public Function OrderByDescending(ByVal func As IFunc) As Enumerable
    Set OrderByDescending = Enumerable.From(MergeSort(sequence_, func, False))
End Function

'Selectが予約語なのでSelect1
'Select is reserved as VBA keyword.
Public Function Select1(ByVal func As IFunc) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        newCol.Add func.Exec(iter)
    Next iter
    
    Set Select1 = Enumerable.From(newCol)
End Function

Public Function SelectMany(ByVal func As IFunc) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant, childIter As Variant
    For Each iter In sequence_
        For Each childIter In func.Exec(iter)
            newCol.Add childIter
        Next childIter
    Next iter
    
    Set SelectMany = Enumerable.From(newCol)
End Function

Public Sub ForEach(ByVal func As IFunc)
    Dim iter As Variant
    For Each iter In sequence_
        Call func.Exec(iter)
    Next iter
End Sub


''IPredicateに依存する処理群
 
Public Function All(ByVal predicate As IPredicate) As Boolean
    ThrowLINQ NotImplementedException
End Function

'Anyは予約語
Public Function Any1(Optional ByVal predicate As IPredicate) As Boolean
    If predicate Is Nothing Then
        Let Any1 = (sequence_.Count <> 0)
        Exit Function
    End If
    
    ThrowLINQ NotImplementedException
    
End Function

Public Function Count(Optional ByVal predicate As IPredicate) As Long
    If predicate Is Nothing Then
        Let Count = sequence_.Count
        Exit Function
    End If
    
    Dim iter As Variant, cnt As Long
    For Each iter In sequence_
        If predicate.Test(iter) Then _
            cnt = cnt + 1
        
    Next iter
    
    Let Count = cnt
    
End Function

Public Function First(Optional ByVal predicate As IPredicate) As Variant
    If predicate Is Nothing Then
        AssignVal First, sequence_.Item(1)
        Exit Function
    End If
    
    ThrowLINQ NotImplementedException
    
End Function

Public Function Last(Optional ByVal predicate As IPredicate) As Variant
    If predicate Is Nothing Then
        AssignVal Last, sequence_.Item(sequence_.Count)
        Exit Function
    End If
    
    ThrowLINQ NotImplementedException
    
End Function

Public Function Single1(Optional ByVal predicate As IPredicate) As Variant
    If predicate Is Nothing Then
        If sequence_.Count = 1 Then
            AssignVal Single1, sequence_.Item(1)
            Exit Function
        Else
            ThrowLINQ IndexOutOfRangeException
        End If
    End If
    
    ThrowLINQ NotImplementedException
    
End Function

Public Function SkipWhile(ByVal predicate As IPredicate) As Enumerable
    ThrowLINQ NotImplementedException
End Function

Public Function TakeWhile(ByVal predicate As IPredicate) As Enumerable
    ThrowLINQ NotImplementedException
End Function

Public Function Where(ByVal predicate As IPredicate) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        If predicate.Test(iter) Then _
            newCol.Add iter
        
    Next iter
    
    Set Where = Enumerable.From(newCol)
End Function






