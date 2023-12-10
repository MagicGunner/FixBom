Attribute VB_Name = "ExcelFunctions"

' 获取当前单元格往下第一个为空的单元格
Function FirstNullY(currentRange As Variant) As Variant

    Dim offSetY As Integer
    offSetY = 0
    Do While currentRange.Offset(offSetY, 0).Value <> ""
        offSetY = offSetY + 1
    Loop
    Set FirstNullY = currentRange.Offset(offSetY, 0)
End Function

' 获取当前单元格往下第一个为空的单元格
Function FirstNullX(currentRange As Variant) As Variant

    Dim offSetX As Integer
    offSetX = 0
    Do While currentRange.Offset(offSetX, 0).Value <> ""
        offSetX = offSetX + 1
    Loop
    Set FirstNullY = currentRange.Offset(offSetX, 0)
End Function

' 设置单元格区域格式，内部边框为细边框，外围边框为粗边框
Sub SetRangeStyle(targetRange As Range)
    targetRange.Borders.LineStyle = xlNone

    ' 设置内部框为加细边框
    With targetRange.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
    End With
    With targetRange.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThick
    End With
    With targetRange.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThick
    End With
    With targetRange.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
    End With
    With targetRange.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
    End With
    With targetRange.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
    End With

End Sub

' 复制当前行到下面一行
Sub OnAddButton()
    ' 获取按钮对象
    Set btn = ActiveSheet.Shapes(Application.Caller)
    Dim currentRow As Range
    Set currentRow = btn.TopLeftCell.EntireRow
    ' 复制当前行
    currentRow.Copy
    ' 在下一行插入新行
    currentRow.Offset(1, 0).Insert Shift:=xlDown
     ' 清除剪贴板内容（可选）
    Application.CutCopyMode = False
     ' 取消当前行的选中状态
    currentRow.Offset(1, 0).Select


End Sub


' 删除当前行
Sub DeleteCurrentRow()

    ' 弹出对话框，获取用户的点击结果
    Dim result As VbMsgBoxResult
    result = MsgBox("当前操作不可逆，是否确定删除", vbQuestion + vbYesNo, "Confirmation")

    ' 根据用户的选择执行相应操作
    If result = vbYes Then
        ' 用户点击了确定按钮
        ' 获取按钮对象
        Set btn = ActiveSheet.Shapes(Application.Caller)

        Dim currentRow As Range
        Set currentRow = btn.TopLeftCell.EntireRow
        currentRow.Delete
    End If


    
End Sub

FuncTion GetDicValueBy

