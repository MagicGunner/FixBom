Attribute VB_Name = "GetFixBom"

Public  Sub AddProject()

    Set titleRange = getFirstNullY(Range("A2"))
    Dim tableHeaderStr() As Variant
    Dim tableColNum As Integer
    tableHeaderStr = Array("�ṹ������", "��������", "������", "�������", "����(mm)", "��Ʒ�ں񹫲�(mm)", "��������", "����", "����")
    tableColNum = UBound(tableHeaderStr) - LBound(tableHeaderStr) + 1
    Set tableHeader = titleRange.Offset(1, 0).Resize(1, tableColNum)
    ' B2��Ԫ��Ϊ��ǰ���з�ʽ������
    If (Range("B1").Value <> "") Then
        titleRange.Value = Range("B1").Value
    Else 
        MsgBox "Name cannot be empty!"
        Exit Sub
    End if
    titleRange.Resize(1, tableColNum).Merge
    tableHeader.Value = tableHeaderStr

End Sub

