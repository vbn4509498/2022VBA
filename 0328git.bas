Attribute VB_Name = "Module1"
Sub 口罩特約藥局排序()
Attribute 口罩特約藥局排序.VB_Description = "本巨集主要用於查詢特約藥局口罩庫存量,並由大到小排序\n"
Attribute 口罩特約藥局排序.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩特約藥局排序 巨集
' 本巨集主要用於查詢特約藥局口罩庫存量
'
' 快速鍵: Ctrl+q
'
    'Create by Yuhsin Hung 2020/3/7
    Range("C1").Select '動作1-選擇B1儲存格
    ActiveWorkbook.Worksheets("工作表2").Sort.SortFields.Clear  '動作2-資料排序設定,根據口罩數量B欄位遞減排序
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort '對全範圍逐行執行排序
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
Sub 口罩特約藥局庫存由小至大排序()
Attribute 口罩特約藥局庫存由小至大排序.VB_Description = "本巨集主要是根據口罩庫存量,進行由小到大的排序,了解當前哪間藥局庫存量最少,方便進行供應管理\n"
Attribute 口罩特約藥局庫存由小至大排序.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' 口罩特約藥局庫存由小至大排序 巨集
' 本巨集主要是根據口罩庫存量,進行由小到大的排序,了解當前哪間藥局庫存量最少,方便進行供應管理
'
' 快速鍵: Ctrl+n
'
    'Create by Yuhsin Hung 2020/3/7
    Range("B1").Select '動作1-選擇B1儲存格
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear '動作2-資料排序設定,根據口罩數量B欄位遞增排序
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort '對全範圍逐行執行排序
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
Attribute CalSumAvg.VB_Description = "計算口罩總量和平均"
Attribute CalSumAvg.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' CalSumAvg 巨集
' 計算口罩總量和平均
'
' 快速鍵: Ctrl+p
'Create by yh hung
    Range("E1").Select '指定E1儲存格
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])" '計算加總
    Range("G1").Select '指定G1儲存格
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])" '計算平均
    Range("G1").Select
    'end of create
End Sub
