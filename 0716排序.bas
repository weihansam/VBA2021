Attribute VB_Name = "Module1"
Option Explicit

Sub filter()
Attribute filter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' filter 巨集
'

'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("E1").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
End Sub
