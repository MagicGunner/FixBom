Attribute VB_Name = "ExcelFunctions"

' ��ȡ��ǰ��Ԫ�����µ�һ��Ϊ�յĵ�Ԫ��
Function FirstNullY(currentRange As Variant) As Variant

    Dim offSetY As Integer
    offSetY = 0
    Do While currentRange.Offset(offSetY, 0).Value <> ""
        offSetY = offSetY + 1
    Loop
    Set FirstNullY = currentRange.Offset(offSetY, 0)
End Function

' ��ȡ��ǰ��Ԫ�����µ�һ��Ϊ�յĵ�Ԫ��
Function FirstNullX(currentRange As Variant) As Variant

    Dim offSetX As Integer
    offSetX = 0
    Do While currentRange.Offset(offSetX, 0).Value <> ""
        offSetX = offSetX + 1
    Loop
    Set FirstNullY = currentRange.Offset(offSetX, 0)
End Function

' ���õ�Ԫ�������ʽ���ڲ��߿�Ϊϸ�߿���Χ�߿�Ϊ�ֱ߿�
Sub SetRangeStyle(targetRange As Range)
    targetRange.Borders.LineStyle = xlNone

    ' �����ڲ���Ϊ��ϸ�߿�
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

' ���Ƶ�ǰ�е�����һ��
Sub OnAddButton()
    ' ��ȡ��ť����
    Set btn = ActiveSheet.Shapes(Application.Caller)
    Dim currentRow As Range
    Set currentRow = btn.TopLeftCell.EntireRow
    ' ���Ƶ�ǰ��
    currentRow.Copy
    ' ����һ�в�������
    currentRow.Offset(1, 0).Insert Shift:=xlDown
     ' ������������ݣ���ѡ��
    Application.CutCopyMode = False
     ' ȡ����ǰ�е�ѡ��״̬
    currentRow.Offset(1, 0).Select


End Sub


' ɾ����ǰ��
Sub DeleteCurrentRow()

    ' �����Ի��򣬻�ȡ�û��ĵ�����
    Dim result As VbMsgBoxResult
    result = MsgBox("��ǰ���������棬�Ƿ�ȷ��ɾ��", vbQuestion + vbYesNo, "Confirmation")

    ' �����û���ѡ��ִ����Ӧ����
    If result = vbYes Then
        ' �û������ȷ����ť
        ' ��ȡ��ť����
        Set btn = ActiveSheet.Shapes(Application.Caller)

        Dim currentRow As Range
        Set currentRow = btn.TopLeftCell.EntireRow
        currentRow.Delete
    End If


    
End Sub

FuncTion GetDicValueBy

