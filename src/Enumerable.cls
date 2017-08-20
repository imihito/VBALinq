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

'Select‚ª—\–ñŒê‚È‚Ì‚ÅSelect1
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

Public Function Where(ByVal predicate As IFunc) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        If predicate.Exec(iter) Then _
            newCol.Add iter
        
    Next iter
    
    Set Where = Enumerable.From(newCol)
End Function

Public Function OfType(ByVal iTypeName As String) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        If VBA.TypeName(iter) Like iTypeName Then _
            newCol.Add iter
        
    Next iter
    
    Set OfType = Enumerable.From(newCol)
End Function

Public Function OrderBy(ByVal func As IFunc) As Enumerable
    Set OrderBy = Enumerable.From(MergeSort(sequence_, func, True))
End Function

Public Function OrderByDescending(ByVal func As IFunc) As Enumerable
    Set OrderByDescending = Enumerable.From(MergeSort(sequence_, func, False))
End Function

Public Function ToCollection() As VBA.Collection
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In sequence_
        newCol.Add iter
    Next iter
    
    Set ToCollection = newCol
    
End Function

Public Function Count(Optional ByVal predicate As IFunc) As Long
    If predicate Is Nothing Then
        Let Count = sequence_.Count
        Exit Function
    End If
    
    Dim iter As Variant, cnt As Long
    For Each iter In sequence_
        If predicate.Exec(iter) Then _
            cnt = cnt + 1
        
    Next iter
    
    Let Count = cnt
    
End Function

Public Function GetEnumerator() As stdole.IEnumVARIANT
Attribute GetEnumerator.VB_UserMemId = -4
'Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = sequence_.[_NewEnum]
End Function
