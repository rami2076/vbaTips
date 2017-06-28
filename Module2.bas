Attribute VB_Name = "Module2"
Option Explicit

'リストオブジェクトの作成(正確にはすでにあるリストオブジェクトにチェックボックスを追加する)
'addメソッドを使用することでより動的にリストを作成できる

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
  
  'リストにcheckBoxを追加
  With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
    For Each lRow In .ListRows
      Debug.Print lRow.Range(Numbering_Column).address
      createCheckBox (lRow.Range(Numbering_Column).address)
      '初期化
      lRow.Range(Numbering_Column).Value = "False"
    Next lRow
  End With
End Sub

'全チェックボックスを削除
Private Sub MyDeleteeCheckBox()
    Dim tCtrl As Variant
    '全てのコントロールの取得
    For Each tCtrl In ActiveSheet.Shapes
        'コントロールの名前をチェック
        If Left(tCtrl.Name, 8) = "CheckBox" Then
            'チェックボックスならば削除
            ActiveSheet.Shapes(tCtrl.Name).Delete
        End If
    Next
End Sub


'チェックボックスの作成
'引数　address　型はString
Private Sub createCheckBox(address As String)
  With ThisWorkbook.Worksheets("Sheet4").OLEObjects.Add(ClassType:="Forms.CheckBox.1", Link:=False, DisplayAsIcon:=False)
         .Object.Caption = "チェック"
         .Object.Font.Size = 9
         .Object.BackColor = &H80FF80 '&HE0E0E0
          .Top = ActiveSheet.Range(address).Top
          .Left = ActiveSheet.Range(address).Left
          .Width = ActiveSheet.Range(address).Width
          .Height = ActiveSheet.Range(address).Height
          .LinkedCell = address
  End With
End Sub


'hiddenの時にTrue
Private Sub getHidden()

Dim lRow As listRow
With ThisWorkbook.Worksheets("Sheet4").ListObjects("MyTable")
    For Each lRow In .ListRows
      Debug.Print lRow.Range(Numbering_Column).address
      Debug.Print lRow.Range(Numbering_Column).EntireRow.Hidden
    Next lRow
  End With
End Sub

'テーブル範囲内に非表示があった場合解除

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
