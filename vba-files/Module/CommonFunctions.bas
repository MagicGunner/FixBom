Attribute VB_Name = "CommonFunctions"

' 获取当前单元格往下第一个为空的单元格，返回Range
Function getFirstNullY(currentRange as Range) As Range
    Dim offSetY As Integer
    offSetY = 0
    Do While currentRange.Offset(offSetY, 0).Value <> ""
        offSetY = offSetY + 1
    Loop
    Set getFirstNullY = currentRange.Offset(offSetY, 0)
End Function

