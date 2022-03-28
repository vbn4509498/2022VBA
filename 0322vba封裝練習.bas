Attribute VB_Name = "Module1"
Sub 口罩特約藥局由大至小排序()
Attribute 口罩特約藥局由大至小排序.VB_Description = "本巨集主要用於查詢特約藥局口罩庫存量,並由大到小排序\n"
Attribute 口罩特約藥局由大至小排序.VB_ProcData.VB_Invoke_Func = "q\n14"
' 嗨
' 口罩特約藥局排序 巨集
' 本巨集主要用於查詢特約藥局口罩庫存量
'
' 快速鍵: Ctrl+q ([問題回復]根據同學現場問題-快捷鍵註解excel不會幫你自動更新喔)
'
    'Create by naiium 2022/3/27
    Range("B1").Select '動作1-選擇B1儲存格
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear  '動作2-資料排序設定,根據口罩數量B欄位遞減排序
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
    
    'Modified by naiiun_2022/03/28
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])" '計算平均
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])" '計算加總
    'End of modifiy
    
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
    'Create by naiiun 2022/3/27
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
    
    'Modified by naiiun_2022/03/28
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])" '計算平均
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])" '計算加總
    'End of modifiy
    
End Sub
