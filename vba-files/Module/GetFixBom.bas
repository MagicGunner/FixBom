Attribute VB_Name = "GetFixBom"

Option Explicit

Public Sub AddProject()

    Dim tableTitle, tableHeader As Range

    Set tableTitle = FirstNullY(Range("A2"))
    Dim tableHeaderStr() As Variant
    Dim tableColumn As Integer
    tableHeaderStr = Array("�ṹ������", "��������", "������", "�������", "����(mm)", "��Ʒ�ں񹫲�(mm)", "��������", "��ע", "����", "����1", "����2")
    tableColumn = UBound(tableHeaderStr) - LBound(tableHeaderStr) + 1
    Set tableHeader = tableTitle.Offset(1, 0).Resize(1, tableColumn)

    If (Range("B1").Value <> "") Then
        tableTitle.Value = Range("B1").Value
    Else
        MsgBox "�������Ʋ���Ϊ��"
        Exit Sub
    End If

    tableTitle.Resize(1, tableColumn).Merge
    tableHeader.Value = tableHeaderStr

    Dim currentRange As Range
    Set currentRange = tableTitle.Offset(2, 0)

    ' AddPost currentRange
    ' Set currentRange = currentRange.offset(1, 0)
    ' AddBeam currentRange
    ' Set currentRange = currentRange.offset(1, 0)
    ' AddBrace currentRange
    ' Set currentRange = currentRange.offset(1, 0)
    ' AddPurlin currentRange
    ' Set currentRange = currentRange.offset(1, 0)
    ' AddPullRod currentRange
    ' Set currentRange = currentRange.offset(1, 0)
    ' AddSupportRod currentRange
    ' Set currentRange = currentRange.offset(1, 0)

    Dim resourceSheet As Worksheet
    Set resourceSheet = ThisWorkbook.Worksheets("Resource")

    Call AddItem(itemName:="����", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="б��", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="б��", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="����", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="����", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="�Ÿ�", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="���Ӽ�", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="���ͼ�", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)
    Call AddItem(itemName:="����", itemStartRange:=currentRange)
    Set currentRange = currentRange.Offset(1, 0)

    Dim targetRange As Range
    Dim border As border

    Set targetRange = tableTitle.Resize(currentRange.row - tableTitle.row, tableColumn)

    SetRangeStyle targetRange

End Sub


Public Sub GeneralBom()

    ' �����������ֵ����
    Dim TotalDic As Object
    Set TotalDic = CreateObject("Scripting.Dictionary")

    Dim startRange As Range
    Set startRange = Range("A2")
    Dim projectCounter As Integer
    projectCounter = 0
    
    ' ��һ�α�����ȡ��Ŀ����
    Do While startRange.Value <> ""
        If startRange.MergeCells Then
            projectCounter = projectCounter + 1
        End If
        Set startRange = startRange.Offset(1, 0)
    Loop

    If projectCounter = 0 Then
        Exit Sub
    End If
    Dim projectDicArr() As Object
    Redim projectDicArr(1 To projectCounter)

    Dim i As Integer
    For i = 1 To projectCounter
        Set projectDicArr(i) = CreateObject("Scripting.Dictionary")
    Next i

    Set startRange = Range("A2").Offset(2, 0)
    '�ڶ��α�����ʼ���ֵ�, Ϊ���ж����������ˣ���ǰ����һ����Ŀ
    Dim currentProjectNum As Integer
    currentProjectNum = 1
    Do While startRange.Value <> ""
        if startRange.MergeCells = false Then
            Dim tempBomItem As BomClasses
            Set tempBomItem = New BomClasses
            Dim itemType As String
            Dim sectionType As String
            Dim section As String
            Dim material As String
            Dim tolerance As String
            Dim remark As String
            Dim displayName As String
            Dim length As Double
            Dim counter As Integer
            itemType = startRange.Value
            sectionType = startRange.Offset(0, 1).Value
            section = startRange.Offset(0, 2).Value
            material = startRange.Offset(0, 3).Value
            length = startRange.Offset(0, 4).Value
            tolerance = startRange.Offset(0, 5).Value
            counter = startRange.Offset(0, 6).Value
            remark = startRange.Offset(0, 7).Value
            If startRange.Offset(0, 8).Value = "" Then
                displayName = itemType
            Else 
                displayName = startRange.Offset(0, 8).Value
            End If
            Call tempBomItem.Initialize(itemType_:=itemType, sectionType_:=sectionType, section_:=section, material_:=material, tolerance_:=tolerance, remark_:=remark, displayName_:=displayName, length_:=length, counter_:=counter)
            projectDicArr(currentProjectNum).Add tempBomItem.Tag, tempBomItem
            Set startRange = startRange.Offset(1, 0)
        Else 
            currentProjectNum = currentProjectNum + 1
            Set startRange = startRange.Offset(2, 0)
        End if
    Loop

    Dim itemCountArr()
    Redim itemCountArr(1 To projectCounter)
    Dim sum As Long
    sum = 0
    For i = 1 To projectCounter
        itemCountArr(i) = projectDicArr(i).Count
        sum = sum + itemCountArr(i)
    Next i

    Do While sum > 0

        Dim tempItems
        
        Dim firstItem As New BomClasses
        Dim firstIndex As Integer
        Dim firstDic As Object
        Set firstDic = CreateObject("Scripting.Dictionary")
        firstIndex = 1
        ' �ҵ���һ���ֵ��е��û��ȡ���
        Do While itemCountArr(firstIndex) = 0
            firstIndex = firstIndex + 1
        Loop
        Set firstDic = projectDicArr(firstIndex)
        tempItems = firstDic.Keys ' �±��0��ʼ
        Set firstItem = firstDic(tempItems(firstDic.Count - itemCountArr(firstIndex)))
        For i = firstIndex + 1 To projectCounter
            Dim tempItem As New BomClasses
            Dim tempDic As Object
            Set tempDic = CreateObject("Scripting.Dictionary")
            Set tempDic = projectDicArr(i)
            tempItems = tempDic.Keys
            Set tempItem = tempDic(tempItems(tempDic.Count - itemCountArr(i)))
            If firstItem.priority > tempItem.priority Then
                Set firstItem = tempItem
                firstIndex = i
            End If
        Next i

        sum = sum - 1
        itemCountArr(firstIndex) = itemCountArr(firstIndex) - 1
        If TotalDic.Exists(firstItem.Tag) = false Then
            TotalDic.Add firstItem.Tag, true
        End If
    Loop

    Dim uniKey
    uniKey = TotalDic.Keys
    Dim keyLength
    keyLength = UBound(uniKey) - LBound(uniKey) + 1
    
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    Dim newRange As Range
    Set newRange = newWorkbook.Worksheets(1).Range("A1")

    newRange.Resize(keyLength, 1).Value = Application.WorksheetFunction.Transpose(uniKey)


End Sub










