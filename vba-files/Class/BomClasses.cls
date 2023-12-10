VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BomClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public itemType, sectionType, section, material, tolerance, remark, displayName As String
Public length As Double
Public counter As Integer
Public priority As Integer ' ���ȼ�Խ������Խ��ǰ

Public Tag As String


Public Sub Initialize(itemType_ As String, sectionType_ As String, section_ As String, material_ As String, tolerance_ As String, remark_ As String, displayName_ As String, length_ As Double, counter_ As Integer)
    itemType = itemType_
    sectionType = sectionType_
    section = section_
    material = material_
    tolerance = tolerance_
    remark = remark_
    displayName = displayName_
    length = length_
    counter = counter_
    SetTag

    Select Case itemType
      Case "����"
        priority = 1
      Case "б��"
        priority = 2
      Case "б��"
        priority = 3
      Case "����"
        priority = 4
      Case "����"
        priority = 5
      Case "�Ÿ�"
        priority = 6
      Case "���Ӽ�"
        priority = 7
      Case "���ͼ�"
        priority = 8
      Case "����"
        priority = 9
    End Select

End Sub

Private Sub SetTag()
    Tag = itemType & "_" & sectionType & "_" & section & "_" & material & "_" & tolerance & "_" & remark & "_" & CStr(length)
End Sub
    