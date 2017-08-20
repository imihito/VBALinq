Attribute VB_Name = "Module1"
Option Explicit

Private Sub TemporaryProc()
    
    Dim nameFunc As CallByFunc
    Set nameFunc = FCallBy("Name", VbGet)
    
    With Enumerable.From(Excel.Workbooks)
        '.All
        
        Debug.Assert .Any1 = True
        'debug.Assert .Any1=True
        
'        Debug.Assert .Count
'        Debug.Assert .Count
'
'        Debug.Assert .ElementAt
'
'        Debug.Assert .First
'        Debug.Assert .First
'
'        Debug.Assert .Last
'        Debug.Assert .Last
'
'        Debug.Assert .Single1
'        Debug.Assert .Single1
    End With
    
End Sub
