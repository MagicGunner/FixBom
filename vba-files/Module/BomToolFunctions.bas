Attribute VB_Name = "BomToolFunctions"

' 0"�ṹ������", 1"��������", 2"������", 3"�������", 4"����(mm)", 5"��Ʒ�ں񹫲�(mm)", 6"��������", 7"����", 8"����"
' Function addPost(Range startRange)
'     Dim resourceSheet As Worksheet
'     Set resourceSheet = Worksheets("Resource") ' ���������쳣����
    

'     Set targetCell = startRange.Offset(0, 1)
    
'     Set sectionValidationRangeStart = resourceSheet.Range("D2")
'     Set sectionValidationRange = sectionValidationRangeStart.Resize(getFirstNotNullY(sectionValidationRangeStart), 1)
'     ' ���������֤
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