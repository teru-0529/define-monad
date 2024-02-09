VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "derive_elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'カラム番号
Private Enum col_no
  派生元項目名 = 1
  項目名J = 2
  項目名EC = 3
  項目名ES = 4
  説明 = 5
End Enum


'シート選択時
Private Sub Worksheet_Activate()
  Call work.element_reflesh '項目リストの最新化
  
  '// フォーマット
  Call areaFormat
End Sub


'項目名の変更追随
Public Sub replace_element_name(ByVal old_val As String, ByVal new_val As String)
  UsedRange.Resize(, 1).Replace what:=old_val, replacement:=new_val
End Sub

'// データLoad(From yaml)
Public Sub load(ByVal data As String)
  Dim records() As String, items() As String
  Dim i As Long, rowLength As Long

  '// 初期化
  UsedRange.Offset(1).Clear

  '// 行ごとのデータに分割
  records = Split(data, vbLf)
  rowLength = UBound(records)
  ReDim wkRange1(rowLength - 1, 2)
  ReDim wkRange2(rowLength - 1, 0)

  '// 列ごとの配列に変換
  For i = 0 To rowLength - 1
    items = Split(records(i), vbTab)

    wkRange1(i, 0) = items(0)  '派生元
    wkRange1(i, 1) = items(1)  '名称_和
    wkRange1(i, 2) = items(2)  '名称_英

    wkRange2(i, 0) = items(3)  '説明
  Next i

  '// 値セット
  Cells(2, col_no.派生元項目名).Resize(rowLength, 3) = wkRange1
  Cells(2, col_no.説明).Resize(rowLength) = wkRange2

  '// フォーマット
  Call areaFormat
End Sub

'フォーマット初期化
Private Sub areaFormat()
  Application.EnableEvents = False

  '// 枠線
  Call Formatter.borderAndFontsize(Range("DERI_FULL"))

  '// リスト選択
  Call Formatter.listSelect(Range("DERI_A"), "項目リスト")

  '// 関数
  Call Formatter.setFunction(Range("DERI_D"), "=camel2snake(RC[-1])")

  '// 条件付き書式
  Call Formatter.conditionFormat(Union(Range("DERI_B"), Range("DERI_E"), Range("DERI_C")), "=AND(RC1<>"""",RC="""")")

  '// IMEモード
  Call Formatter.imeMode(Union(Range("DERI_B"), Range("DERI_E")), Formatter.IME_MODE_ON)
  Call Formatter.imeMode(Range("DERI_C"), Formatter.IME_MODE_OFF)

  Application.EnableEvents = True
End Sub
