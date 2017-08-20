VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'汎用判定用インターフェイス
'IFuncに依存
Option Explicit

'IFunc(UseFunc)を実行した結果を元に判定を行う。
Public Function Test(iElement As Variant) As Boolean
Attribute Test.VB_UserMemId = 0
'Attribute Test.VB_UserMemId = 0
End Function

Public Property Get UseFunc() As IFunc
End Property
