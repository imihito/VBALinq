VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'JPN:汎用コールバック用インターフェイス
Option Explicit

'JPN:何かしらの処理を実行する
Public Function Exec(iElement As Variant) As Variant
Attribute Exec.VB_UserMemId = 0
End Function
