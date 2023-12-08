Attribute VB_Name = "FixBomFunctions"

Sub AddPost(postStartRange as Range)
    Dim resourceSheet As Worksheet
    Set resourceSheet = ThisWorkbook.Worksheets("Resource")
    postStartRange.Value = "����"

    Dim rangeDY, rangeDX As Long
    Dim validationRangeStart, validationRangeEnd, validationRange as Range

    ' ����������������������֤
    Dim sectionTypeCell As Range
    Set sectionTypeCell = postStartRange.Offset(0, 1)
    Set validationRangeStart = resourceSheet.Range("D2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.Row - validationRangeStart.Row

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

    ' ���������������������֤
    Dim sectionMaterialCell As Range
    Set sectionMaterialCell = postStartRange.Offset(0, 3)
    Set validationRangeStart = resourceSheet.Range("H2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.Row - validationRangeStart.Row

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
    Set toleranceCell = postStartRange.Offset(0, 5)
    Set validationRangeStart = resourceSheet.Range("C2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.Row - validationRangeStart.Row

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
    Set remarkCell = postStartRange.Offset(0, 7)
    Set validationRangeStart = resourceSheet.Range("B2")
    Set validationRangeEnd = FirstNullY(validationRangeStart)
    rangeDY = validationRangeEnd.Row - validationRangeStart.Row

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
    Dim addButton As Shape
    Dim deleteButton As Shape
    Dim buttonRange As Range
    Set buttonRange = postStartRange.Offset(0, 9)
    Set addButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left, Top:=buttonRange.Top, Width:=buttonRange.Width / 2, Height:=buttonRange.Height)
    Set buttonRange = buttonRange.Offset(0, 1)
    Set deleteButton = ActiveSheet.Shapes.AddFormControl(xlButtonControl, Left:=buttonRange.Left, Top:=buttonRange.Top, Width:=buttonRange.Width / 2, Height:=buttonRange.Height)


End Sub