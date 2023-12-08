Attribute VB_Name = "ExcelFunctions"

Function FirstNullY(currentRange as Variant) As Variant

    Dim offSetY As Integer
    offSetY = 0
    Do While currentRange.Offset(offSetY, 0).Value <> ""
        offSetY = offSetY + 1
    Loop
    Set FirstNullY = currentRange.Offset(offSetY, 0)
End Function