Attribute VB_Name = "Module2"
Option Explicit

'���X�g�I�u�W�F�N�g�̍쐬(���m�ɂ͂��łɂ��郊�X�g�I�u�W�F�N�g�Ƀ`�F�b�N�{�b�N�X��ǉ�����)
'add���\�b�h���g�p���邱�Ƃł�蓮�I�Ƀ��X�g���쐬�ł���

Private Const Numbering_Column As Integer = 1
Private Const Name_Column As Integer = 2
Private Const Flg_Column As Integer = 3


Sub createList()
Attribute createList.VB_ProcData.VB_Invoke_Func = " \n14"

  Dim list As ListObject
  Dim lRows As ListRows
  Dim lRow As listRow
'  Dim coll As New Collection
  
  MyDeleteeCheckBox
  
  'coll.Add (1)
  'coll.Add (2)
  'coll.Add (3)
  
  '���X�g��checkBox��ǉ�
  With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
    For Each lRow In .ListRows
      Debug.Print lRow.Range(Numbering_Column).address
      createCheckBox (lRow.Range(Numbering_Column).address)
      '������
      lRow.Range(Numbering_Column).Value = "False"
    Next lRow
  End With
End Sub

'�S�`�F�b�N�{�b�N�X���폜
Private Sub MyDeleteeCheckBox()
    Dim tCtrl As Variant
    '�S�ẴR���g���[���̎擾
    For Each tCtrl In ActiveSheet.Shapes
        '�R���g���[���̖��O���`�F�b�N
        If Left(tCtrl.Name, 8) = "CheckBox" Then
            '�`�F�b�N�{�b�N�X�Ȃ�΍폜
            ActiveSheet.Shapes(tCtrl.Name).Delete
        End If
    Next
End Sub


'�`�F�b�N�{�b�N�X�̍쐬
'�����@address�@�^��String
Private Sub createCheckBox(address As String)
  With ThisWorkbook.Worksheets("Sheet4").OLEObjects.Add(ClassType:="Forms.CheckBox.1", Link:=False, DisplayAsIcon:=False)
         .Object.Caption = "�`�F�b�N"
         .Object.Font.Size = 9
         .Object.BackColor = &H80FF80 '&HE0E0E0
          .Top = ActiveSheet.Range(address).Top
          .Left = ActiveSheet.Range(address).Left
          .Width = ActiveSheet.Range(address).Width
          .Height = ActiveSheet.Range(address).Height
          .LinkedCell = address
  End With
End Sub


'hidden�̎���True
Private Sub getHidden()

Dim lRow As listRow
With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
    For Each lRow In .ListRows
      Debug.Print lRow.Range(Numbering_Column).address
      Debug.Print lRow.Range(Numbering_Column).EntireRow.Hidden
    Next lRow
  End With
End Sub

'�e�[�u���͈͓��ɔ�\�����������ꍇ����

Private Sub defaultedList()
Dim lRow As listRow
Dim lColumn As ListColumn
With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
    For Each lRow In .ListRows
      Debug.Print lRow.Range(Numbering_Column).address
      If lRow.Range.EntireRow.Hidden Then
        lRow.Range.EntireRow.Hidden = False
       End If
    Next lRow
    
    For Each lColumn In .ListColumns
      'Debug.Print lColumn.Range(Numbering_Column).address
      If lColumn.Range.EntireColumn.Hidden Then
        lColumn.Range.EntireColumn.Hidden = False
       End If
    Next lColumn
    
  End With
  
  
 End Sub


Private Sub defaultListFileter()
Dim lColumn As ListColumn
Dim iterator As Integer
 With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
  For iterator = 1 To .ListColumns.Count
  .Range.AutoFilter Field:=iterator
  Next iterator
  End With

End Sub



Sub bolean()
Dim bool As Boolean
bool = (1 = 1)
End Sub
