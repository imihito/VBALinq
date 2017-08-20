Attribute VB_Name = "SpeedTest"
Option Explicit

Sub SpeedTest()
    Dim buf As VBA.Collection
    
    Const LOOP_COUNT& = 1000
    Dim stTime!, i&
    
    stTime = VBA.Timer
    For i = 1 To LOOP_COUNT
        Set buf = NormalVBA
    Next i
    Debug.Print "Normal", VBA.Timer - stTime
    
    stTime = VBA.Timer
    For i = 1 To LOOP_COUNT
        Set buf = UseLinq
    Next i
    Debug.Print "Linq", VBA.Timer - stTime
    
    stTime = VBA.Timer
    For i = 1 To LOOP_COUNT
        Set buf = DelayExec
    Next i
    Debug.Print "Delay", VBA.Timer - stTime
    
End Sub

Private Function NormalVBA() As VBA.Collection
    Dim oCol As VBA.Collection
    Set oCol = New VBA.Collection
    
    Dim tWb As Excel.Workbook
    Dim tSh As Object
    For Each tWb In Excel.Workbooks
        For Each tSh In tWb.Sheets
            If tSh.Name Like "Sheet[0-9]" Then
                oCol.Add tSh
            End If
        Next tSh
    Next tWb
    Set NormalVBA = oCol
End Function

Private Function UseLinq() As VBA.Collection
    Set UseLinq = _
            Enumerable.From(Workbooks) _
            .SelectMany(CallByFunc.Init("Sheets", VbGet)) _
            .Where(CallByFunc("Name", VbGet).SetChild(CompareOperator(opLike, "Sheet[0-9]"))) _
            .ToCollection()

End Function

Private Function DelayExec() As VBA.Collection
    Dim oCol As VBA.Collection
    Set oCol = New VBA.Collection
    
    Dim sheetsFunc As IFunc
    Set sheetsFunc = CallByFunc("Sheets", VbGet)
    
    Dim predict As IFunc
    Set predict = CallByFunc("Name", VbGet).SetChild(CompareOperator(opLike, "Sheet[0-9]"))
    
    Dim iter1 As Variant, iter2 As Variant
    For Each iter1 In Enumerable.From(Workbooks)
        For Each iter2 In sheetsFunc.Exec(iter1)
            If predict.Exec(iter2) Then
                oCol.Add iter2
            End If
        Next iter2
    Next iter1
    
    
    Set DelayExec = oCol
End Function

