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

Private clsCollection As VBA.Collection

Public Function From(ByVal iEnumerable As Variant) As Enumerable
Attribute From.VB_UserMemId = 0
'Attribute From.VB_UserMemId = 0
    If TypeOf iEnumerable Is VBA.Collection Then
        Set clsCollection = iEnumerable
    Else
        Set clsCollection = New VBA.Collection
        Dim iter As Variant
        For Each iter In iEnumerable
            clsCollection.Add iter
        Next iter
    End If
    
    Set From = Me
End Function

Public Function Where(ByVal iPredicate As IFunc) As Enumerable
    Dim newCol As VBA.Collection: Set newCol = New VBA.Collection
    
    Dim iter As Variant
    For Each iter In clsCollection
        If iPredicate.Exec(iter) Then
            newCol.Add iter
        End If
    Next iter
    
    Set Where = LinqFrom(newCol)
End Function


Public Function GetEnumerator() As stdole.IEnumVARIANT
Attribute GetEnumerator.VB_UserMemId = -4
'Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = clsCollection.[_NewEnum]
End Function
