Attribute VB_Name = "GetFixBom"

Option Explicit

Public Sub AddProject()

    Dim tableTitle, tableHeader as Range

    Set tableTitle = FirstNullY(Range("A2"))
    Dim tableHeaderStr() As Variant
    Dim tableColumn As Integer
    tableHeaderStr = Array("结构件类型", "截面类型", "截面规格", "截面材质", "长度(mm)", "成品壁厚公差(mm)", "单套数量", "备注", "名称", "操作")
    tableColumn = UBound(tableHeaderStr) - LBound(tableHeaderStr) + 1
    Set tableHeader = tableTitle.Offset(1, 0).Resize(1, tableColumn)

    If (Range("B1").Value <> "") Then
        tableTitle.Value = Range("B1").Value
    Else 
        MsgBox "排列名称不可为空"
        Exit Sub
    End if

    tableTitle.Resize(1, tableColumn).Merge
    tableHeader.Value = tableHeaderStr

    AddPost(tableTitle.Offset(2, 0))


End Sub










