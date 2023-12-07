Attribute VB_Name = "GetFixBom"

Public  Sub AddProject()

    Set titleRange = getFirstNullY(Range("A2"))
    Dim tableHeaderStr() As Variant
    Dim tableColNum As Integer
    tableHeaderStr = Array("结构件类型", "截面类型", "截面规格", "截面材质", "长度(mm)", "成品壁厚公差(mm)", "单套数量", "名称", "操作")
    tableColNum = UBound(tableHeaderStr) - LBound(tableHeaderStr) + 1
    Set tableHeader = titleRange.Offset(1, 0).Resize(1, tableColNum)
    ' B2单元格为当前排列方式的名字
    If (Range("B1").Value <> "") Then
        titleRange.Value = Range("B1").Value
    Else 
        MsgBox "Name cannot be empty!"
        Exit Sub
    End if
    titleRange.Resize(1, tableColNum).Merge
    tableHeader.Value = tableHeaderStr

End Sub

