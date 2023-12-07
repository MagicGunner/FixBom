Attribute VB_Name = "BomToolFunctions"

' 0"结构件类型", 1"截面类型", 2"截面规格", 3"截面材质", 4"长度(mm)", 5"成品壁厚公差(mm)", 6"单套数量", 7"名称", 8"操作"
' Function addPost(Range startRange)
'     Dim resourceSheet As Worksheet
'     Set resourceSheet = Worksheets("Resource") ' 后续考虑异常处理
    

'     Set targetCell = startRange.Offset(0, 1)
    
'     Set sectionValidationRangeStart = resourceSheet.Range("D2")
'     Set sectionValidationRange = sectionValidationRangeStart.Resize(getFirstNotNullY(sectionValidationRangeStart), 1)
'     ' 添加数据验证
'     With targetCell.Validation
'         .Delete
'         .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'         xlBetween, Formula1:="=" & sectionValidationRange.Address
'         .IgnoreBlank = True
'         .InCellDropdown = True
'         .ShowInput = True
'         .ShowError = True
'     End With

' End Function