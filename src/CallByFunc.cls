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
Private memberName_ As String
Private callType_ As VBA.VbCallType
Private arguments_() As Variant
Private childFunc_ As IFunc

Property Get MemberName() As String:        Let MemberName = memberName_:   End Property
Property Get CallType() As VBA.VbCallType:  Let CallType = callType_:       End Property
Property Get Arguments() As Variant:        Let Arguments = arguments_:     End Property
Property Get ChildFunc() As IFunc:  Set ChildFunc = childFunc_:             End Property
Property Set ChildFunc(ByVal func As IFunc): Set childFunc_ = func:         End Property

'コンストラクタ・既定のプロシージャ

Friend Function Init( _
        ByVal iMemeberName As String, _
        ByVal iCallType As VBA.VbCallType, _
        ByRef iArgs() As Variant _
    ) As CallByFunc
'Attribute Init.VB_UserMemId = 0
    If Me Is CallByFunc Then
        With New CallByFunc
            Set Init = .Init(iMemeberName, iCallType, iArgs)
        End With 'New CallByFunc
    Else
        Static alreadyInit As Boolean
        If alreadyInit Then _
            ThrowLINQ InvalidOperationException
        
        Let memberName_ = iMemeberName
        Let callType_ = iCallType
        Let arguments_ = iArgs
        
        Let alreadyInit = True
        
        Set Init = Me
    End If
End Function

Public Function SetChild(ByVal func As IFunc) As CallByFunc
    Set childFunc_ = func
    Set SetChild = Me
End Function

Public Function RemoveChild() As CallByFunc
    Set childFunc_ = Nothing
    Set RemoveChild = Me
End Function

Public Function AsIFunc() As IFunc: Set AsIFunc = Me: End Function

Private Function IFunc_Exec(iElement As Variant) As Variant
    Dim tmp As Variant
    AssignVal tmp, CallByNameEx(iElement, memberName_, callType_, arguments_)
    
    If childFunc_ Is Nothing Then
        AssignVal IFunc_Exec, tmp
    Else
        AssignVal IFunc_Exec, childFunc_.Exec(tmp)
    End If
End Function
