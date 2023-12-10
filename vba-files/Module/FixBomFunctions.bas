Attribute VB_Name = "FixBomFunctions"

Sub AddItem(itemName As String, itemStartRange As Range)
    Dim resourceSheet As Worksheet
    Set resourceSheet = ThisWorkbook.Worksheets("Resource")
    itemStartRange.Value = itemName

    Dim rangeDY, rangeDX As Long
    Dim validationRangeStart, validationRangeEnd, validationRange As Range

    ' 增加截面类型数据验证
    Dim sectionTypeCell As Range
    Set sectionTypeCell = itemStartRange.Offset(0, 1)

    Select Case itemName
        Case "立柱"
            Set validationRangeStart = resourceSheet.Range("D2")
        Case "斜梁"
            Set validationRangeStart = resourceSheet.Range("E2")
        Case "斜撑"
            Set validationRangeStart = resourceSheet.Range("F2")
        Case "檩条"
            Set validationRangeStart = resourceSheet.Range("G2")
        Case "拉杆"
            Set validationRangeStart = resourceSheet.Range("I2")
        Case "撑杆"
            Set validationRangeStart = resourceSheet.Range("J2")
        Case "连接件"
            Set validationRangeStart = resourceSheet.Range("K2")
        Case "异型件"
            Set validationRangeStart = resourceSheet.Range("L2")
        Case "其他"
            Set validationRangeStart = resourceSheet.Range("M2")
    End Select

    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.row - validationRangeStart.row

    If (rangeDY <> 0) Then
        Set validationRange = validationRangeStart.Resize(rangeDY, 1)

        With sectionTypeCell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='" & resourceSheet.Name & "'!" & validationRange.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    End If
    

    ' 增加截面材质数据验证
    Dim sectionMaterialCell As Range
    Set sectionMaterialCell = itemStartRange.Offset(0, 3)
    Set validationRangeStart = resourceSheet.Range("H2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.row - validationRangeStart.row

    Set validationRange = validationRangeStart.Resize(rangeDY, 1)

    With sectionMaterialCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & resourceSheet.Name & "'!" & validationRange.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' 增加成品壁厚公差数据验证
    Dim toleranceCell As Range
    Set toleranceCell = itemStartRange.Offset(0, 5)
    Set validationRangeStart = resourceSheet.Range("C2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.row - validationRangeStart.row

    Set validationRange = validationRangeStart.Resize(rangeDY, 1)

    With toleranceCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & resourceSheet.Name & "'!" & validationRange.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' 备注数据验证
    Dim remarkCell As Range
    Set remarkCell = itemStartRange.Offset(0, 7)
    Set validationRangeStart = resourceSheet.Range("B2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.row - validationRangeStart.row

    Set validationRange = validationRangeStart.Resize(rangeDY, 1)

    With remarkCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & resourceSheet.Name & "'!" & validationRange.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' 操作按钮
    Dim buttonRange As Range
    Set buttonRange = itemStartRange.Offset(0, 9)
    AddButtonShapes buttonRange


End Sub


Sub AddButtonShapes(buttonRange As Range)

    Dim addButton As Shape
    Set addButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left + buttonRange.Width * 0.1, Top:=buttonRange.Top + buttonRange.Height * 0.1, Width:=buttonRange.Width * 0.8, Height:=buttonRange.Height * 0.8)
    addButton.TextFrame.Characters.Text = "新增"
    addButton.OnAction = "OnAddButton"

    Set buttonRange = buttonRange.Offset(0, 1)

    Dim deleteButton As Shape
    Set deleteButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left + buttonRange.Width * 0.1, Top:=buttonRange.Top + buttonRange.Height * 0.1, Width:=buttonRange.Width * 0.8, Height:=buttonRange.Height * 0.8)
    deleteButton.TextFrame.Characters.Text = "删除"
    deleteButton.OnAction = "DeleteCurrentRow"
End Sub
