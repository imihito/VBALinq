VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PassThru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Implements IFunc
Option Explicit

Private Function IFunc_Exec(iElement As Variant) As Variant
    AssignVal IFunc_Exec, iElement
End Function
