Attribute VB_Name = "FixBomFunctions"

Sub AddItem(itemName As String, itemStartRange As Range)
    Dim resourceSheet As Worksheet
    Set resourceSheet = ThisWorkbook.Worksheets("Resource")
    itemStartRange.Value = itemName

    Dim rangeDY, rangeDX As Long
    Dim validationRangeStart, validationRangeEnd, validationRange As Range

    ' ���ӽ�������������֤
    Dim sectionTypeCell As Range
    Set sectionTypeCell = itemStartRange.Offset(0, 1)

    Select Case itemName
        Case "����"
            Set validationRangeStart = resourceSheet.Range("D2")
        Case "б��"
            Set validationRangeStart = resourceSheet.Range("E2")
        Case "б��"
            Set validationRangeStart = resourceSheet.Range("F2")
        Case "����"
            Set validationRangeStart = resourceSheet.Range("G2")
        Case "����"
            Set validationRangeStart = resourceSheet.Range("I2")
        Case "�Ÿ�"
            Set validationRangeStart = resourceSheet.Range("J2")
        Case "���Ӽ�"
            Set validationRangeStart = resourceSheet.Range("K2")
        Case "���ͼ�"
            Set validationRangeStart = resourceSheet.Range("L2")
        Case "����"
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
    

    ' ���ӽ������������֤
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

    ' ���ӳ�Ʒ�ں񹫲�������֤
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

    ' ��ע������֤
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

    ' ������ť
    Dim buttonRange As Range
    Set buttonRange = itemStartRange.Offset(0, 9)
    AddButtonShapes buttonRange


End Sub


Sub AddButtonShapes(buttonRange As Range)

    Dim addButton As Shape
    Set addButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left + buttonRange.Width * 0.1, Top:=buttonRange.Top + buttonRange.Height * 0.1, Width:=buttonRange.Width * 0.8, Height:=buttonRange.Height * 0.8)
    addButton.TextFrame.Characters.Text = "����"
    addButton.OnAction = "OnAddButton"

    Set buttonRange = buttonRange.Offset(0, 1)

    Dim deleteButton As Shape
    Set deleteButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left + buttonRange.Width * 0.1, Top:=buttonRange.Top + buttonRange.Height * 0.1, Width:=buttonRange.Width * 0.8, Height:=buttonRange.Height * 0.8)
    deleteButton.TextFrame.Characters.Text = "ɾ��"
    deleteButton.OnAction = "DeleteCurrentRow"
End Sub
