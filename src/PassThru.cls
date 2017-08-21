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
Implements IPredicate
Option Explicit

Private Function IFunc_Exec(iElement As Variant) As Variant
    AssignVal IFunc_Exec, iElement
End Function

Private Function IPredicate_Test(iElement As Variant) As Boolean
    Let IPredicate_Test = VBA.CBool(iElement)
End Function

Private Property Get IPredicate_UseFunc() As IFunc
    Set IPredicate_UseFunc = Nothing
End Property

Public Function AsIFunc() As IFunc: Set AsIFunc = Me: End Function
Public Function AsIPredicate() As IPredicate: Set AsIPredicate = Me: End Function
