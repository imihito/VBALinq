VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppRunFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IFunc
Option Explicit

Private clsMemberName As String
Private clsCallType As VBA.VbCallType
Private clsArguments() As Variant
Private clsChildCallBack As IFunc

Property Get MemberName() As String:        Let MemberName = clsMemberName:         End Property
Property Get CallType() As VBA.VbCallType:  Let CallType = clsCallType:             End Property
Property Get Arguments() As Variant():      Let Arguments = clsArguments:           End Property

Private Function IFunc_Exec(iElement As Variant) As Variant
    clsArguments(0) = "Proc"
    AssignVal clsArguments(1), iElement
    AssignVal IFunc_Exec, CallByNameEx(Application, "Run", VbMethod, clsArguments)
End Function

'aaa
Friend Function Init(ByVal iProcName As String, ParamArray iArgs() As Variant) As AppRunFunc
    ThrowLINQ NotImplementedException
End Function
