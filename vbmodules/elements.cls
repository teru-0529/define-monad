VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const COL_NAMEJP = 1
Const COL_DOMAIN = 4
Const COL_EXAMPLE = 10

Dim val As Variant '変更前の仮値

'シート選択時
Private Sub Worksheet_Activate()
  Application.ScreenUpdating = False
  '// フォーマット
  Call areaFormat
  Application.ScreenUpdating = True
End Sub

'セルの選択イベント
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error GoTo ENDUP
  
  '項目名を選択した場合
  If Target.Column = COL_NAMEJP Then
    val = ""
    If Target.value = "" Then Exit Sub
    If Target.Cells.Count <> 1 Then Exit Sub
  
    val = Target.value
  End If

ENDUP:
End Sub

'セルの変更イベント
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo ENDUP
  
  If Target.Cells.Count <> 1 Then Exit Sub
  
  '項目名の変更の場合
  If Target.Column = COL_NAMEJP Then
    If val = "" Then Exit Sub
    If Target.value = "" Then Exit Sub
  
    If Target.value <> val Then
      Call derive_elements.replace_element_name(val, Target.value)
      Call segments.replace_element_name(val, Target.value)
    End If
    Exit Sub
  
ENDUP:
  
  'ドメインの変更の場合
  ElseIf Target.Column = COL_DOMAIN Then
    '値リセット
    Target.Offset(, 1).Resize(, 5).ClearContents
    Call domainFormat(Target)
  End If

End Sub

'// データLoad(From yaml)
Public Sub load(ByVal data As String)
  Dim records() As String, items() As String
  Dim i As Long, rowLength As Long

  '// 初期化
  UsedRange.Offset(1).Clear
  Columns(COL_EXAMPLE).NumberFormatLocal = "@"

  '// 入力値が存在しない場合、フォーマットのみ実施
  If data = "" Then
    '// フォーマット
    Call areaFormat
    Exit Sub
  End If
    
  '// 行ごとのデータに分割
  records = Split(data, vbLf)
  rowLength = UBound(records)
  ReDim wkRange1(rowLength - 1, 1)
  ReDim wkRange2(rowLength - 1, 7)

  '// 列ごとの配列に変換
  For i = 0 To rowLength - 1
    items = Split(records(i), vbTab)

    wkRange1(i, 0) = items(0)  '名称_和
    wkRange1(i, 1) = items(1)  '名称_英

    wkRange2(i, 0) = items(2)  'ドメイン
    wkRange2(i, 1) = items(3)  '正規表現
    wkRange2(i, 2) = items(4)  '最小桁数
    wkRange2(i, 3) = items(5)  '最大桁数
    wkRange2(i, 4) = items(6)  '最小値
    wkRange2(i, 5) = items(7)  '最大値
    wkRange2(i, 6) = items(8)  '例
    wkRange2(i, 7) = items(9)  '説明
  Next i

  '// 値セット
  Cells(2, COL_NAMEJP).Resize(rowLength, 2) = wkRange1
  Cells(2, COL_DOMAIN).Resize(rowLength, 8) = wkRange2

  '// フォーマット
  Call areaFormat
  Call domainFormat(Cells(2, COL_DOMAIN).Resize(rowLength))
End Sub

'フォーマット
Private Sub areaFormat()
  Application.EnableEvents = False

  '// 枠線
  Call Formatter.borderAndFontsize(Range("ELEM_FULL"))
  
  '// リスト選択
  Call Formatter.listSelect(Range("ELEM_D"), "ドメインタイプKEY")
  
  '// 関数
  Call Formatter.setFunction(Range("ELEM_C"), "=camel2snake(RC[-1])")
  Call Formatter.setFunction(Range("ELEM_L"), "=countif(派生元項目KEY,RC[-11])")
  Call Formatter.setFunction(Range("ELEM_M"), "=countif(区分値項目KEY,RC[-12])")

  '// 表示形式
  Call Formatter.formatLocal(Union(Range("ELEM_L"), Range("ELEM_M")), "0_;;;")
    
  '// 条件付き書式
  Call Formatter.condFunction(Union(Range("ELEM_B"), Range("ELEM_D"), Range("ELEM_J"), Range("ELEM_K")), "=AND(RC1<>"""",RC="""")")
  
  '// IMEモード
  Call Formatter.imeMode(Union(Range("ELEM_A"), Range("ELEM_K")), Formatter.IME_MODE_ON)
  Call Formatter.imeMode(Range("ELEM_B"), Formatter.IME_MODE_OFF)

  Application.EnableEvents = True
End Sub

'// ドメインフォーマット
Private Sub domainFormat(ByRef Target As Range)
  Application.ScreenUpdating = False
  
  Dim elm As Range
  Dim areas As DecolationRanges
  
  '// フォーマット初期化
  Call Formatter.reset(Target.Offset(, 1).Resize(, 5))

  '// 行ごとにフォーマット対象を分類
  For Each elm In Target
    Call domain_types.devideRange(elm, areas)
  Next elm

  '// フォーマット
  Call Formatter.stringSpec(Util.unionRange(areas.stringRangeR, areas.stringRange))
  Call Formatter.lengthSpec(Util.unionRange(areas.lengthRangeR, areas.lengthRange))
  Call Formatter.numericSpec(Util.unionRange(areas.numericRangeR, areas.numericRange))
  Call Formatter.unabledSpec(areas.unabledRange)
  Call Formatter.condReqired(Util.unionRange(areas.stringRangeR, Util.unionRange(areas.lengthRangeR, areas.numericRangeR)))

  Application.ScreenUpdating = True
End Sub
