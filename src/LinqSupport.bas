Attribute VB_Name = "LinqSupport"
'�I���I��LINQ�p�⏕����

Option Explicit
'Option Private Module

Enum LinqErrNumber
    ArgumentException = 5
    IndexOutOfRangeException = 9
    InvalidOperationException = 17
    NotImplementedException = 32768
End Enum

'�\�[�g�p���[�U�[��`�^
Private Type SortElement
    Object As Object
    Value As Variant
End Type

'����p
'outVal = setVal
Sub AssignVal(ByRef outVal As Variant, ByRef setVal As Variant)
    If VBA.IsObject(setVal) Then
        Set outVal = setVal
    Else
        Let outVal = setVal
    End If
End Sub

'VBA.Interaction.CallByName�g��
    '������ParamArray�ł͂Ȃ��AVariant�^�z��Ŏ󂯎���Ď��s����
    '�����ׂ������̂��߁A�����̐��ɐ����L��
    
    '��{��CallByFunc�p�����Ǒ��ɂ����p�ł������Ȃ̂ň�U�O��
Function CallByNameEx(ByVal iObject As Object, ByVal iProcName As String, ByVal iCallType As VBA.VbCallType, ByRef iArgs() As Variant) As Variant
    If VBA.CBool(Not Not iArgs) Then
        Select Case UBound(iArgs)
            Case -1:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType)
            Case 0:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0))
            Case 1:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1))
            Case 2:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2))
            Case 3:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3))
            Case 4:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4))
            Case 5:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5))
            Case 6:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6))
            Case 7:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7))
            Case 8:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8))
            Case 9:     AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9))
            Case 10:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10))
            Case 11:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10), iArgs(11))
            Case 12:    AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType, iArgs(0), iArgs(1), iArgs(2), iArgs(3), iArgs(4), iArgs(5), iArgs(6), iArgs(7), iArgs(8), iArgs(9), iArgs(10), iArgs(11), iArgs(12))
            
            Case Else:  ThrowLINQ ArgumentException
        End Select
    Else
        AssignVal CallByNameEx, VBA.CallByName(iObject, iProcName, iCallType)
    End If
End Function

'ComparePredicate�p
    'ComparePredicate�̍\�������܂��Ă��Ȃ��̂ň�U�O��



Sub ThrowLINQ(ByVal iErrNo As LinqErrNumber)
    Select Case iErrNo
        Case NotImplementedException
            Err.Raise iErrNo, , "�@�\���܂���������Ă��܂���B" & vbNewLine & "NotImplemented"
        Case Else
            Err.Raise iErrNo
    End Select
End Sub

Function InsertToHead( _
             ByVal iBaseArray As Variant, _
        ParamArray iInsertElements() As Variant _
    ) As Variant
    
    Const ARRAY_BASE& = 0
    
    If LBound(iBaseArray) <> ARRAY_BASE Then _
        ThrowLINQ IndexOutOfRangeException
    
    Dim insertCnt As Long
    insertCnt = UBound(iInsertElements) + 1
    
    ReDim Preserve iBaseArray(ARRAY_BASE To UBound(iBaseArray) + insertCnt)
    
    Dim i As Long
    For i = UBound(iBaseArray) - insertCnt To ARRAY_BASE Step -1
        AssignVal iBaseArray(i + insertCnt), iBaseArray(i)
    Next i
    
    For i = ARRAY_BASE To insertCnt - 1
        AssignVal iBaseArray(i), iInsertElements(i)
    Next i
    
    Let InsertToHead = iBaseArray
    
End Function


'[�I�u�W�F�N�g�p�}�[�W�\�[�g - Qiita](http://qiita.com/nukie_53/items/88ff2227c20cb2f04344 "�I�u�W�F�N�g�p�}�[�W�\�[�g - Qiita")
'�̉���
'����
    'Objects    �F�\�[�g�������I�u�W�F�N�g��VBA.Collection
    'MemberName �F�v���p�e�B�iVbGet�j�⃁�\�b�h�iVbMethod�j�̖��O�B
    'CallType   �F�ȗ��BMemberName�̎�ށBVbGet��������VbMethod�B�ȗ���VbGet�B
    'Ascending  �F�ȗ��BTrue�Ȃ珸���AFalse�Ȃ�~���B�ȗ���True�B

'�߂�l
    '�\�[�g���ꂽVBA.Collection

Public Function MergeSort(ByVal Objects As VBA.Collection, _
                           ByVal KeyCallBack As IFunc, _
                           Optional ByVal Ascending As Boolean = True _
                        ) As VBA.Collection
    
    Dim basArray() As SortElement
    ReDim basArray(1 To Objects.Count)

    Dim i&, obj As Object
    For Each obj In Objects
        i = i + 1
        Set basArray(i).Object = obj
        Let basArray(i).Value = KeyCallBack.Exec(obj)
    Next obj

    '�R�s�[���쐬�B�����ƍl����Η̈�m�ۂ����ł��ǂ������B
    Dim OutArray() As SortElement
    OutArray = basArray

    '�\�[�g
    Call RecurseMergeSort(basArray, OutArray, 1, Objects.Count, Ascending)

    Erase basArray

    '�o�͗p�ɓ��꒼��
    Dim oCol As VBA.Collection
    Set oCol = New VBA.Collection
    For i = 1 To Objects.Count
        oCol.Add OutArray(i).Object
    Next i

    Set MergeSort = oCol

End Function


Private Sub RecurseMergeSort( _
        ByRef InptArray() As SortElement, _
        ByRef OutArray() As SortElement, _
        ByVal Start As Long, _
        ByVal Length As Long, _
        ByVal Ascending As Boolean)

    Dim halfLen As Long
    halfLen = VBA.CLng(Length / 2)

    '�O���̃\�[�g
    If halfLen >= 2 Then
        Call RecurseMergeSort(OutArray, InptArray, Start, halfLen, Ascending)
    End If

    '�㔼�̃\�[�g
    If Length - halfLen >= 2 Then
        Call RecurseMergeSort(OutArray, InptArray, Start + halfLen, Length - halfLen, Ascending)
    End If

    '�O�������̓Y�����ƍő�l
    Dim lwIndex As Long:    lwIndex = Start
    Dim lwMax As Long:      lwMax = Start + halfLen - 1

    '�㔼�����̓Y�����ƍő�l
    Dim upIndex As Long:    upIndex = Start + halfLen
    Dim upMax As Long:      upMax = Start + Length - 1

    '�S�̂̓Y�����ƍő�l
    Dim oIndex As Long:     oIndex = Start
    Dim oMax As Long:       oMax = Start + Length - 1

    Dim leftIndex As Long   '�]��p

    Dim flg As Boolean

    For oIndex = Start To oMax Step 1
        '�l�������Ȃ珇�Ԉێ�
        flg = (InptArray(lwIndex).Value = InptArray(upIndex).Value)

        '�l����������Ȃ��ꍇ�A�Ĕ���
        If Not flg Then flg = (Ascending = (InptArray(lwIndex).Value < InptArray(upIndex).Value))

        If flg Then
            OutArray(oIndex) = InptArray(lwIndex)
            If lwIndex = lwMax Then
                leftIndex = upIndex
                Exit For
            Else
                lwIndex = lwIndex + 1
            End If
        Else
            OutArray(oIndex) = InptArray(upIndex)
            If upIndex = upMax Then
                leftIndex = lwIndex
                Exit For
            Else
                upIndex = upIndex + 1
            End If
        End If
    Next oIndex

    'Next oIndex���΂������C���N�������g����
    For oIndex = oIndex + 1 To oMax Step 1
        OutArray(oIndex) = InptArray(leftIndex)
        leftIndex = leftIndex + 1
    Next oIndex

End Sub


