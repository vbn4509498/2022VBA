Attribute VB_Name = "Module1"
Sub �f�n�S���ħ��Ƨ�()
Attribute �f�n�S���ħ��Ƨ�.VB_Description = "�������D�n�Ω�d�߯S���ħ��f�n�w�s�q,�åѤj��p�Ƨ�\n"
Attribute �f�n�S���ħ��Ƨ�.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�S���ħ��Ƨ� ����
' �������D�n�Ω�d�߯S���ħ��f�n�w�s�q
'
' �ֳt��: Ctrl+q
'
    'Create by Yuhsin Hung 2020/3/7
    Range("B1").Select '�ʧ@1-���B1�x�s��
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear  '�ʧ@2-��ƱƧǳ]�w,�ھڤf�n�ƶqB��컼��Ƨ�
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort '����d��v�����Ƨ�
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'End of create
    
    'Modify by
     Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    'end of modify
    
    
    
End Sub
Sub �f�n�S���ħ��w�s�Ѥp�ܤj�Ƨ�()
Attribute �f�n�S���ħ��w�s�Ѥp�ܤj�Ƨ�.VB_Description = "�������D�n�O�ھڤf�n�w�s�q,�i��Ѥp��j���Ƨ�,�F�ѷ�e�����ħ��w�s�q�̤�,��K�i������޲z\n"
Attribute �f�n�S���ħ��w�s�Ѥp�ܤj�Ƨ�.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' �f�n�S���ħ��w�s�Ѥp�ܤj�Ƨ� ����
' �������D�n�O�ھڤf�n�w�s�q,�i��Ѥp��j���Ƨ�,�F�ѷ�e�����ħ��w�s�q�̤�,��K�i������޲z
'
' �ֳt��: Ctrl+n
'
    'Create by Yuhsin Hung 2020/3/7
    Range("B1").Select '�ʧ@1-���B1�x�s��
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear '�ʧ@2-��ƱƧǳ]�w,�ھڤf�n�ƶqB��컼�W�Ƨ�
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort '����d��v�����Ƨ�
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'End of create
    'Modify by
     Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    'end of modify
    
End Sub

Sub CalSumAvg()
Attribute CalSumAvg.VB_Description = "�p��f�n�`�q�M����"
Attribute CalSumAvg.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' CalSumAvg ����
' �p��f�n�`�q�M����
'
' �ֳt��: Ctrl+p
'Create by yh hung
    Range("E1").Select '���wE1�x�s��
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])" '�p��[�`
    Range("G1").Select '���wG1�x�s��
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])" '�p�⥭��
    Range("G1").Select
    'end of create
End Sub
