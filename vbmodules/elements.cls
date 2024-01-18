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

'カラム番号
Private Enum col_no
  項目名J = 1
  項目名EC = 2
  項目名ES = 3
 ドメイン = 4
  正規表現 = 5
  最小桁数 = 6
  最大桁数 = 7
  最小値 = 8
  最小値含 = 9
  最大値 = 10
  最大値含 = 11
  API例 = 12
  説明 = 13
  API型1 = 14
  API型2 = 15
  種類数 = 16
  区分値数 = 17
End Enum

'デコレーション種類
Private Enum decoration_type
  文字列 = 1
  桁数 = 2
  数値 = 3
  真偽値 = 4
  入力不可 = 9
  入力必須 = 100
End Enum

'セル設定種類
Private Type decolation_cells
  dec_range As Range
  dec_type As decoration_type
  reqired As Boolean
End Type
  
  
Const MAX_COL = 17


Dim val As Variant '変更前の仮値

'シート選択時
Private Sub Worksheet_Activate()
  Dim start_row As Integer: start_row = Cells(rows.Count, "A").End(xlUp).Row
  Call init_format(start_row + 1, 0)
End Sub

'セルの選択イベント
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error GoTo ENDUP
  
  '項目名を選択した場合
  If Target.Column = 項目名J Then
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
  If Target.Column = col_no.項目名J Then
    If val = "" Then Exit Sub
    If Target.value = "" Then Exit Sub
  
    If Target.value <> val Then
      Call derive_elements.replace_element_name(val, Target.value)
      Call segments.replace_element_name(val, Target.value)
    End If
    Exit Sub
  
ENDUP:
  
  'ドメインの変更の場合
  ElseIf Target.Column = col_no.ドメイン Then
    '値リセット
    Call format_domain_type(Target, True)
  End If

End Sub

'フォーマット初期化
Public Sub init_format(ByVal start_row As Long, ByVal rows As Long)
  Application.EnableEvents = False
  
  Dim f_rows As Integer: f_rows = rows + ADDED_ROWS 'フォーマット化行数(ADDED_ROWS分多めにフォーマット)
  Dim cond_format As FormatCondition

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

  '背景色
  Union(Cells(start_row, col_no.項目名ES).Resize(f_rows), Cells(start_row, col_no.API型1).Resize(f_rows, 4)).Interior.Color = RGB(255, 255, 213) 'CREAM
  
  '入力規則
  With Cells(start_row, col_no.ドメイン).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ドメインタイプKEY"
  End With

  '関数
  Cells(start_row, col_no.項目名ES).Resize(f_rows).Formula = "=camel2snake(RC[-1])"
  Cells(start_row, col_no.API型1).Resize(f_rows).Formula = "=iferror(vlookup(RC[-10],ドメインタイプDATA,6,False)&"""","""")"
  Cells(start_row, col_no.API型2).Resize(f_rows).Formula = "=iferror(vlookup(RC[-11],ドメインタイプDATA,7,False)&"""","""")"
  Cells(start_row, col_no.種類数).Resize(f_rows).Formula = "=countif(派生元項目KEY,RC[-15])"
  Cells(start_row, col_no.区分値数).Resize(f_rows).Formula = "=countif(区分値項目KEY,RC[-16])"
  
  '位置
  Union(Cells(start_row, col_no.最小値含).Resize(f_rows), Cells(start_row, col_no.最大値含).Resize(f_rows)).HorizontalAlignment = xlCenter

  '表示形式
  Union(Cells(start_row, col_no.最小値含).Resize(f_rows), Cells(start_row, col_no.最大値含).Resize(f_rows)).NumberFormatLocal = """含"";;;"
  Cells(start_row, col_no.API例).Resize(f_rows).NumberFormatLocal = "@"
  Union(Cells(start_row, col_no.種類数).Resize(f_rows), Cells(start_row, col_no.区分値数).Resize(f_rows)).NumberFormatLocal = "0_;;;"

  
  '条件付き書式
  Set cond_format = Union(Cells(start_row, col_no.項目名EC).Resize(f_rows), Cells(start_row, col_no.ドメイン).Resize(f_rows), _
   Cells(start_row, col_no.API例).Resize(f_rows), Cells(start_row, col_no.説明).Resize(f_rows)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC1<>"""",RC="""")")
  cond_format.Interior.Color = work.Range("入力必須").Interior.Color
  
  'IMEモード
  With Union(Cells(start_row, col_no.項目名J).Resize(f_rows), Cells(start_row, col_no.説明).Resize(f_rows)).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOn
  End With
  With Cells(start_row, col_no.項目名EC).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOff
  End With
  
  
  Application.EnableEvents = True
End Sub


'ドメインタイプの変更イベント
Private Sub format_domain_type(ByVal Target As Range, ByVal is_reset As Boolean)
  Dim ng As decolation_cells:  ng.dec_type = decoration_type.入力不可:  ng.reqired = False
  Dim str As decolation_cells:  str.dec_type = decoration_type.文字列:  str.reqired = False
  Dim str_req As decolation_cells:  str_req.dec_type = decoration_type.文字列:  str_req.reqired = True
  Dim keta As decolation_cells:  keta.dec_type = decoration_type.桁数:  keta.reqired = False
  Dim keta_req As decolation_cells:  keta_req.dec_type = decoration_type.桁数:  keta_req.reqired = True
  Dim num As decolation_cells:  num.dec_type = decoration_type.数値:  num.reqired = True
  Dim bool As decolation_cells:  bool.dec_type = decoration_type.真偽値: bool.reqired = False
  
  Dim elm As Range
  Dim i As Long
  Dim dec As Integer
  
  '初期化（値リセット）
  If is_reset Then Target.Offset(, 1).Resize(, 7).ClearContents
  
  '行ごとにフォーマット対照を選別
  For Each elm In Target
    i = elm.Row
    
    '正規表現
    dec = domain_types.get_val(elm.value, 7, -1)
    If dec > 0 Then
      Set str_req.dec_range = Util.union_range(str_req.dec_range, Cells(i, col_no.正規表現))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.正規表現))
    Else
      Set str.dec_range = Util.union_range(str.dec_range, Cells(i, col_no.正規表現))
    End If
    'タイプ：時間の場合、正規表現を設定
    If elm.value = "時間" Then Cells(i, col_no.正規表現).value = "^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$"
    
    '最小桁数
    dec = domain_types.get_val(elm.value, 8, -1)
    If dec > 0 Then
      Set keta_req.dec_range = Util.union_range(keta_req.dec_range, Cells(i, col_no.最小桁数))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.最小桁数))
    Else
      Set keta.dec_range = Util.union_range(keta.dec_range, Cells(i, col_no.最小桁数))
    End If
    
    '最大桁数
    dec = domain_types.get_val(elm.value, 9, -1)
    If dec > 0 Then
      Set keta_req.dec_range = Util.union_range(keta_req.dec_range, Cells(i, col_no.最大桁数))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.最大桁数))
    End If
    
    '最小値/最大値
    dec = domain_types.get_val(elm.value, 10, -1)
    If dec > 0 Then
      Set num.dec_range = Util.union_range(num.dec_range, Cells(i, col_no.最小値))
      Set num.dec_range = Util.union_range(num.dec_range, Cells(i, col_no.最大値))
      Set bool.dec_range = Util.union_range(bool.dec_range, Cells(i, col_no.最小値含))
      Set bool.dec_range = Util.union_range(bool.dec_range, Cells(i, col_no.最大値含))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.最小値).Resize(, 4))
    End If
        
  Next elm
   
  
  Call decoration(str_req, is_reset)
  Call decoration(str, is_reset)
  Call decoration(keta_req, is_reset)
  Call decoration(keta, is_reset)
  Call decoration(num, is_reset)
  Call decoration(bool, is_reset)
  Call decoration(ng, is_reset)

End Sub


'入力エリアのデコレーション
Private Sub decoration(ByRef Target As decolation_cells, ByVal is_reset As Boolean)
  Dim cond_format As FormatCondition
  
  If Target.dec_range Is Nothing Then Exit Sub
      
      
  If Target.dec_type = decoration_type.文字列 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateInputOnly
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "@"
    
    End With
    
  ElseIf Target.dec_type = decoration_type.桁数 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1000"
      .Validation.ErrorMessage = "0-1000の間で入力可能"
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.数値 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-99999999", Formula2:="99999999"
      .Validation.ErrorMessage = "数値のみ入力可能"
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.真偽値 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1"
      .Validation.ErrorMessage = "1(値を含む)のみ入力可能"
      .Validation.IMEMode = xlIMEModeOff
      If is_reset Then .value = 1
    
    End With
    
  ElseIf Target.dec_type = decoration_type.入力不可 Then
    With Target.dec_range
      .Interior.Color = work.Range("入力不可").Interior.Color
      .Validation.Delete
      .Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="0"
      .Validation.ErrorMessage = "入力不可項目です"
      .Validation.IMEMode = xlIMEModeOff
    
    End With
  End If
  
  '必須
  If Target.reqired Then
    Set cond_format = Target.dec_range.FormatConditions.Add(xlBlanksCondition)
    cond_format.Interior.Color = work.Range("入力必須").Interior.Color
  
  Else
    Target.dec_range.FormatConditions.Delete
  
  End If
End Sub

'項目名を指定して要素値を返す。
Public Function get_val(ByVal name As String, ByVal col As Integer, ByVal def_val) As Variant
  Dim elm As Range: Set elm = Range("項目名KEY").Find(name, LookAt:=xlWhole)
  
  '存在しなければデフォルト値を返却
  If elm Is Nothing Then
    get_val = def_val
  
  Else
    get_val = elm.Offset(, col).value
  
  End If
  
End Function


'■■■■■↓IO　■■■■■

'save_data読込
'// FIXME:★★★★Refactering to Backend
Public Sub load_(ByRef data() As String)
  Const PARAMS = 13
  Dim row_length As Long: row_length = (UBound(data) + 1) / PARAMS
  Dim i As Long
  ReDim wk_range_1(row_length - 1, 1)
  ReDim wk_range_2(row_length - 1, 9)
  
  '描画箇所のフォーマット
  Call init_format(2, row_length)
  
  For i = 0 To row_length - 1
    wk_range_1(i, 0) = from_yaml(data(i * PARAMS + 0))  '名称_和
    wk_range_1(i, 1) = from_yaml(data(i * PARAMS + 1))  '名称_英
    
    wk_range_2(i, 0) = from_yaml(data(i * PARAMS + 2))  'ドメイン
    wk_range_2(i, 1) = from_yaml(data(i * PARAMS + 3))  '正規表現
    wk_range_2(i, 2) = from_yaml(data(i * PARAMS + 4))  '最小桁数
    wk_range_2(i, 3) = from_yaml(data(i * PARAMS + 5))  '最大桁数
    wk_range_2(i, 4) = from_yaml(data(i * PARAMS + 6))  '最小値
    wk_range_2(i, 5) = from_yaml(data(i * PARAMS + 7))  '最小値含む
    wk_range_2(i, 6) = from_yaml(data(i * PARAMS + 8))  '最大値
    wk_range_2(i, 7) = from_yaml(data(i * PARAMS + 9))  '最大値含む
    wk_range_2(i, 8) = from_yaml(data(i * PARAMS + 10))  '例
    wk_range_2(i, 9) = from_yaml(data(i * PARAMS + 11))  '説明
  Next i
  
  '値セット
  Cells(2, col_no.項目名J).Resize(row_length, 2) = wk_range_1
  Cells(2, col_no.ドメイン).Resize(row_length, 10) = wk_range_2
  
  'format
  Call format_domain_type(Cells(2, col_no.ドメイン).Resize(row_length), False)
 
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
  
  Call push_array(data, indent(1) & "- " & to_yaml("name_jp", Cells(i, col_no.項目名J).value))
  Call push_array(data, indent(2) & to_yaml("name_en", Cells(i, col_no.項目名EC).value))
  Call push_array(data, indent(2) & to_yaml("domain", Cells(i, col_no.ドメイン).value))
  Call push_array(data, indent(2) & to_yaml("reg_ex", Cells(i, col_no.正規表現).value))
  Call push_array(data, indent(2) & to_yaml("min_digits", Cells(i, col_no.最小桁数).value))
  Call push_array(data, indent(2) & to_yaml("max_digits", Cells(i, col_no.最大桁数).value))
  Call push_array(data, indent(2) & to_yaml("min_value", Cells(i, col_no.最小値).value))
  Call push_array(data, indent(2) & to_yaml("min_value_boundory", Cells(i, col_no.最小値含).value))
  Call push_array(data, indent(2) & to_yaml("max_value", Cells(i, col_no.最大値).value))
  Call push_array(data, indent(2) & to_yaml("max_value_boundory", Cells(i, col_no.最大値含).value))
  Call push_array(data, indent(2) & to_yaml("example", Cells(i, col_no.API例).value))
  Call push_array(data, indent(2) & to_yaml("description", Cells(i, col_no.説明).value))
  Call push_array(data, "")
End Sub


'table_elements書込
Public Sub out_table(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call table_item(data, elm)
  Next elm
  
End Sub

Private Sub table_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  Call push_array(data, "- " & to_yaml("name_jp", Cells(i, col_no.項目名J).value))
  Call push_array(data, indent(1) & to_yaml("name_en", Cells(i, col_no.項目名ES).value))
  If Cells(i, col_no.ドメイン).value = "区分値" Then
    Call push_array(data, indent(1) & to_yaml("db_model", Cells(i, col_no.項目名ES).value))
  Else
    Call push_array(data, indent(1) & to_yaml("db_model", domain_types.get_db_model(Cells(i, col_no.ドメイン).value, Cells(i, col_no.最大桁数).value)))
  End If
  Call push_array(data, indent(1) & to_yaml("constraint", db_constraint(elm, Cells(i, col_no.項目名ES).value)))
  Call push_array(data, indent(1) & to_yaml("must_not_null", must_not_null(Cells(i, col_no.ドメイン).value)))
  Call push_array(data, indent(1) & to_yaml("is_string_default", is_string_default(Cells(i, col_no.ドメイン).value)))
  Call push_array(data, indent(1) & to_yaml("description", Cells(i, col_no.説明).value))
  Call push_array(data, indent(1) & to_yaml("is_origin", "true"))
  Call push_array(data, indent(1) & to_yaml("origin", ""))
  Call push_array(data, "")
End Sub


'api_elements書込
Public Sub out_api(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call api_item(data, elm)
  Next elm
  
End Sub

Private Sub api_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  Call push_array(data, "# " & Cells(i, col_no.項目名J).value)
  Call push_array(data, "# : " & Cells(i, col_no.説明).value)
  Call push_array(data, Cells(i, col_no.項目名EC).value & ":")
  Call push_array(data, indent(1) & to_yaml("type", Cells(i, col_no.API型1).value))
  If Cells(i, col_no.API型2).value <> "" Then Call push_array(data, indent(1) & to_yaml("format", Cells(i, col_no.API型2).value))
  If Cells(i, col_no.正規表現).value <> "" Then Call push_array(data, indent(1) & to_yaml("pattern", Cells(i, col_no.正規表現).value))
  If Cells(i, col_no.最小桁数).value <> "" Then Call push_array(data, indent(1) & to_yaml("minLength", Cells(i, col_no.最小桁数).value))
  If Cells(i, col_no.最大桁数).value <> "" Then Call push_array(data, indent(1) & to_yaml("maxLength", Cells(i, col_no.最大桁数).value))
  If Cells(i, col_no.最小値).value <> "" Then Call push_array(data, indent(1) & to_yaml("minimum", Cells(i, col_no.最小値).value))
  If Cells(i, col_no.最大値).value <> "" Then Call push_array(data, indent(1) & to_yaml("maximum", Cells(i, col_no.最大値).value))
  If Cells(i, col_no.最小値含).value <> "" Then Call push_array(data, indent(1) & "exclusiveMinimum: true")
  If Cells(i, col_no.最大値含).value <> "" Then Call push_array(data, indent(1) & "exclusiveMaximum: true")
  If Cells(i, col_no.ドメイン).value = "区分値" Then
    Call push_array(data, indent(1) & "enum:")
    Call segments.api_value(data, Cells(i, col_no.項目名J).value, indent(2) & "- ")
  End If
  Call push_array(data, indent(1) & to_yaml("description", Cells(i, col_no.説明).value))
  If Cells(i, col_no.API型1).value = "string" Then
    Call push_array(data, indent(1) & to_yaml("example", "'" & Cells(i, col_no.API例).value) & "'")
  Else
    Call push_array(data, indent(1) & to_yaml("example", Cells(i, col_no.API例).value))
  End If
  Call push_array(data, "")
End Sub


'type_ddl書込
Public Sub out_type_ddl(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    Dim i As Long: i = elm.Row
    If Cells(i, col_no.ドメイン).value = "区分値" Then
      Call push_array(data, "-- " & Cells(i, col_no.項目名J).value)
      Call push_array(data, "DROP TYPE IF EXISTS " & Cells(i, col_no.項目名ES).value & ";")
      Call push_array(data, "CREATE TYPE " & Cells(i, col_no.項目名ES).value & " AS enum (")
      Call segments.ddl_value(data, Cells(i, col_no.項目名J).value, indent(1))
      Call push_array(data, ");")
      Call push_array(data, "")
    End If
  Next elm
  
End Sub


'elements.tsv(表示用)出力
Public Sub out_view(ByVal line_sep As Integer)
  Dim data() As String, elm As Range, i As Long
  
  Call push_array(data, array2tabstr("名称(JP)", "名称(EN)", "ドメイン", "API(type)", "API(format)", "正規表現", "最小桁数", "最大桁数", "最小値", "最小値含む", "最大値", "最大値含む", "API(example)", "説明"))
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then
      i = elm.Row
      Call push_array(data, array2tabstr( _
        Cells(i, col_no.項目名J).value, _
        Cells(i, col_no.項目名EC).value, _
        Cells(i, col_no.ドメイン).value, _
        Cells(i, col_no.API型1).value, _
        Cells(i, col_no.API型2).value, _
        Cells(i, col_no.正規表現).value, _
        Cells(i, col_no.最小桁数).value, _
        Cells(i, col_no.最大桁数).value, _
        Cells(i, col_no.最小値).value, _
        change_boundary(Cells(i, col_no.最小値含).value), _
        Cells(i, col_no.最大値).value, _
        change_boundary(Cells(i, col_no.最大値含).value), _
        Cells(i, col_no.API例).value, _
        Cells(i, col_no.説明).value & " "))
    End If
  Next elm
  
  Call text_io.plain_out(Config.VIEW_ELEMENTS, data, line_sep)
End Sub

Private Function change_boundary(ByVal val As Variant) As Variant
  If val = 1 Then change_boundary = "含" Else change_boundary = "-"
End Function

'DB制約文字列を構成
Private Function db_constraint(Target As Range, field_name As String) As String
  Dim le As String, ge As String
  
  If Target.Offset(, 4).value <> "" Then '正規表現の場合
    db_constraint = "(" & field_name & " ~* '" & Target.Offset(, 4).value & "')"
  
  ElseIf Target.Offset(, 5).value <> "" Then '最低文字列長の場合
    db_constraint = "(LENGTH(" & field_name & ") >= " & Target.Offset(, 5).value & ")"
  
  ElseIf Target.Offset(, 7).value <> "" Then '数値範囲の場合
    If Target.Offset(, 8).value = 1 Then le = " >= " Else le = " > "
    If Target.Offset(, 10).value = 1 Then ge = " <= " Else ge = " < "
    db_constraint = "(" & field_name & le & Target.Offset(, 7).value & " AND " & field_name & ge & Target.Offset(, 9).value & ")"
  
  Else
    db_constraint = ""
  End If
End Function

'DB制約文字列を構成(別名用)
Public Function db_constraint_by_name(name As String, field_name As String) As String
  Dim elm As Range: Set elm = Range("項目名KEY").Find(name, LookAt:=xlWhole)
  
  '存在しなければデフォルト値を返却
  If elm Is Nothing Then
    db_constraint_by_name = ""
  
  Else
    db_constraint_by_name = db_constraint(elm, field_name)
  
  End If
End Function

