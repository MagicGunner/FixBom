Attribute VB_Name = "GetFixBom"

Option Explicit

Public Sub AddProject()

    Dim tableTitle, tableHeader as Range

    Set tableTitle = FirstNullY(Range("A2"))
    Dim tableHeaderStr() As Variant
    Dim tableColumn As Integer
    tableHeaderStr = Array("�ṹ������", "��������", "������", "�������", "����(mm)", "��Ʒ�ں񹫲�(mm)", "��������", "��ע", "����", "����")
    tableColumn = UBound(tableHeaderStr) - LBound(tableHeaderStr) + 1
    Set tableHeader = tableTitle.Offset(1, 0).Resize(1, tableColumn)

    If (Range("B1").Value <> "") Then
        tableTitle.Value = Range("B1").Value
    Else 
        MsgBox "�������Ʋ���Ϊ��"
        Exit Sub
    End if

    tableTitle.Resize(1, tableColumn).Merge
    tableHeader.Value = tableHeaderStr

    AddPost(tableTitle.Offset(2, 0))


End Sub










