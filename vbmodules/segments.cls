VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "segments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'カラム番号
Private Enum col_no
  項目名J = 1
  区分値 = 2
  区分名 = 3
  説明 = 4
End Enum

Const MAX_COL = 4


'シート選択時
Private Sub Worksheet_Activate()
  Call work.segment_reflesh '区分値リストの最新化

  Dim start_row As Integer: start_row = Cells(rows.Count, "A").End(xlUp).Row
  Call init_format(start_row + 1, 0)
End Sub


'項目名の変更追随
Public Sub replace_element_name(ByVal old_val As String, ByVal new_val As String)
  UsedRange.Resize(, 1).Replace what:=old_val, replacement:=new_val
End Sub


'フォーマット初期化
Public Sub init_format(ByVal start_row As Long, ByVal rows As Long)
  Application.EnableEvents = False
  
  Dim f_rows As Integer: f_rows = rows + ADDED_ROWS 'フォーマット化行数(ADDED_ROWS分多めにフォーマット)
  Dim cond_format_1 As FormatCondition, cond_format_2 As FormatCondition

  '初期行以降をクリア
  UsedRange.Offset(start_row - 1).Clear
  
  '枠線
  With Cells(start_row, 1).Resize(f_rows, MAX_COL)
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(68, 114, 196) 'DEEP_BLUE
    .Borders(xlEdgeTop).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Borders(xlInsideHorizontal).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Font.Size = 9
  End With

  '入力規則
  With Cells(start_row, col_no.項目名J).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=区分値リスト"
  End With

  '条件付き書式
  Set cond_format_1 = Cells(start_row, col_no.項目名J).Resize(f_rows).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC<>"""",ISNA(MATCH(RC,区分値リスト,0)))")
  cond_format_1.Interior.Color = work.Range("入力必須").Interior.Color
  Set cond_format_2 = Union(Cells(start_row, col_no.区分値).Resize(f_rows), _
   Cells(start_row, col_no.区分名).Resize(f_rows), Cells(start_row, col_no.説明).Resize(f_rows)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC1<>"""",RC="""")")
  cond_format_2.Interior.Color = work.Range("入力必須").Interior.Color
  
  'IMEモード
  With Union(Cells(start_row, col_no.区分名).Resize(f_rows), Cells(start_row, col_no.説明).Resize(f_rows)).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOn
  End With
  With Cells(start_row, col_no.区分値).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOff
  End With
  
  Application.EnableEvents = True
End Sub


'■■■■■↓IO　■■■■■

'save_data読込
'// FIXME:★★★★Refactering to Backend
Public Sub load_(ByRef data() As String)
  Const PARAMS = 4
  Dim row_length As Long: row_length = (UBound(data) + 1) / PARAMS
  Dim i As Long
  ReDim wk_range_1(row_length - 1, 3)
  
  '描画箇所のフォーマット
  Call init_format(2, row_length)
  
  For i = 0 To row_length - 1
    wk_range_1(i, 0) = from_yaml(data(i * PARAMS + 0)) '項目名
    wk_range_1(i, 1) = from_yaml(data(i * PARAMS + 1)) '区分値
    wk_range_1(i, 2) = from_yaml(data(i * PARAMS + 2)) '区分名
    wk_range_1(i, 3) = from_yaml(data(i * PARAMS + 3)) '説明
  Next i
  
  '値セット
  Cells(2, col_no.項目名J).Resize(row_length, 4) = wk_range_1
  
End Sub


'save_data書込
'// FIXME:★★★★Refactering to Backend
Public Sub save_(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call save_item(data, elm)
  Next elm
  
End Sub

'// FIXME:★★★★Refactering to Backend
Private Sub save_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  Call push_array(data, indent(1) & "- " & to_yaml("key", Cells(i, col_no.項目名J).value))
  Call push_array(data, indent(2) & to_yaml("value", Cells(i, col_no.区分値).value))
  Call push_array(data, indent(2) & to_yaml("name", Cells(i, col_no.区分名).value))
  Call push_array(data, indent(2) & to_yaml("description", Cells(i, col_no.説明).value))
End Sub


'APIパラメータの出力
Public Sub api_value(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long:
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(data, indent & Cells(i, col_no.区分値).value)
    '##TODO: enum定義の際シングルクォートが必要かどうか
    'If elm.value = name Then Call push_array(data, indent & "'" & Cells(i, col_no.区分値C).value & "'")
  Next elm
End Sub


'APIパラメータの出力
Public Sub api_desc(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long:
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(data, indent & Cells(i, col_no.区分値).value & " : " & Cells(i, col_no.区分名).value)
  Next elm
End Sub


'DDLパラメータの出力
Public Sub ddl_value(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long
  Dim wk() As String
  
  '一度wkに入れる（最後だけカンマつけたくないため）
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(wk, indent & "'" & Cells(i, col_no.区分値).value & "'")
  Next elm
  
  For i = 0 To UBound(wk) - 1
    Call push_array(data, wk(i) & ",")
  Next i
  Call push_array(data, wk(UBound(wk)))
  
End Sub


'segments.tsv(表示用)出力
Public Sub out_view(ByVal line_sep As Integer)
  Dim data() As String, elm As Range, i As Long
  
  Call push_array(data, array2tabstr("項目名", "区分値", "区分名", "説明"))
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then
      i = elm.Row
      Call push_array(data, array2tabstr( _
        Cells(i, col_no.項目名J).value, _
        Cells(i, col_no.区分値).value, _
        Cells(i, col_no.区分名).value, _
        Cells(i, col_no.説明).value & " "))
    End If
  Next elm
  
  Call text_io.plain_out(Config.VIEW_SEGMENTS, data, line_sep)
End Sub
