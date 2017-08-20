Attribute VB_Name = "M_Sample"
Option Explicit

Sub Sample()
    
    '����̃C���X�^���X��From���\�b�h��ForEach�ł�����̂��w�肵�ăC���X�^���X
    Dim myLinq As Enumerable
    Set myLinq = Enumerable.From(ThisWorkbook.Worksheets)
    
    'IFunc�̓f���Q�[�g�̑���
    Dim selectFunc As IFunc
    Set selectFunc = FCallBy("UsedRange", VbGet)
    
    'Worksheet�����ɂ��āAUsedRange���擾
    Dim selectLinq As Enumerable
    Set selectLinq = myLinq.OfType("Worksheet").Select1(selectFunc)
    
    
    Dim nameFunc As CallByFunc
    Set nameFunc = FCallBy("Name", VbGet)
    
    'Name�v���p�e�B��Like���Z�q��"Sheet[0-9]"�ƃ}�b�`�������
        '�uCallByFunc.SetChild�v��CallByFunc�̌��ʂ����ƂɎ���IFunc�Ăяo��
    Dim whereLinq As Enumerable
    Set whereLinq = myLinq.Where(PCompare(nameFunc, opLike, "Sheet[0-9]"))
    
    '�S�u�b�N�̃��[�N�V�[�g�ꗗ
    Dim selectManyLinq As Enumerable
    Set selectManyLinq = Enumerable.From(Workbooks) _
                            .SelectMany(FCallBy("Worksheets", VbGet))
    
    Debug.Print "----Before Sort----"
    Dim tWs As Excel.Worksheet
    For Each tWs In selectManyLinq
        Debug.Print tWs.Name
    Next tWs
    
    Debug.Print "----After Sort----"
    For Each tWs In selectManyLinq.OrderBy(nameFunc)
        Debug.Print tWs.Name
    Next tWs
    
    Stop
    
End Sub

