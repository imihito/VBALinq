VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallByFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'CallByNameを行うための情報を格納するクラス

Implements IFunc
Option Explicit

'スコープ適当。Argumentsをどんな形で見せるか…
Private clsMemberName As String
Private clsCallType As VBA.VbCallType
Private clsArguments() As Variant
Private clsChildCallBack As IFunc

Property Get MemberName() As String:        Let MemberName = clsMemberName:         End Property
Property Get CallType() As VBA.VbCallType:  Let CallType = clsCallType:             End Property
Function Arguments() As Variant:       Let Arguments = clsArguments:           End Function
Property Get ChildCallBack() As IFunc:  Set ChildCallBack = clsChildCallBack:   End Property
Property Set ChildCallBack(ByVal iCB As IFunc): Set clsChildCallBack = iCB:     End Property

'コンストラクタ・既定のプロシージャ
    '既定のインスタンスを使って
        'CallByFunc( ～
    'と使用することを想定


Public Function Init( _
             ByVal iMemeberName As String, _
             ByVal iCallType As VBA.VbCallType, _
        ParamArray iArgs() As Variant _
    ) As CallByFunc
Attribute Init.VB_UserMemId = 0
'

'Attribute Init.VB_UserMemId = 0
    If Me Is CallByFunc Then
        Dim callArgs() As Variant
        ReDim callArgs(LBound(iArgs) To UBound(iArgs) + 2)
        callArgs(0) = iMemeberName: callArgs(1) = iCallType
        Dim i As Long
        For i = 2 To UBound(callArgs)
            callArgs(i) = iArgs(i - 2)
        Next i
        
        Dim tCallByFunc As CallByFunc
        Set tCallByFunc = New CallByFunc
        Set Init = CallByNameEx(tCallByFunc, "Init", VbMethod, callArgs)
        Exit Function
    Else
    
        Static alreadyInit As Boolean
        If alreadyInit Then Err.Raise 17
        
        Let clsMemberName = iMemeberName
        Let clsCallType = iCallType
        Let clsArguments = iArgs
        
        Let alreadyInit = True
        
        Set Init = Me
    End If
End Function

Function SetChild(ByVal callBack As IFunc) As CallByFunc
    Set clsChildCallBack = callBack
    Set SetChild = Me
End Function

Private Function IFunc_Exec(iElement As Variant) As Variant
    Dim tmp As Variant
    AssignVal tmp, CallByNameEx(iElement, clsMemberName, clsCallType, clsArguments)
    
    If clsChildCallBack Is Nothing Then
        AssignVal IFunc_Exec, tmp
    Else
        AssignVal IFunc_Exec, clsChildCallBack.Exec(tmp)
    End If
End Function



