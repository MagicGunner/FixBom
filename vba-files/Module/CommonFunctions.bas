Attribute VB_Name = "CommonFunctions"

' ��ȡ��ǰ��Ԫ�����µ�һ��Ϊ�յĵ�Ԫ�񣬷���Range
Function getFirstNullY(currentRange as Range) As Range
    Dim offSetY As Integer
    offSetY = 0
    Do While currentRange.Offset(offSetY, 0).Value <> ""
        offSetY = offSetY + 1
    Loop
    Set getFirstNullY = currentRange.Offset(offSetY, 0)
End Function

