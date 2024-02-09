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
  最大値 = 9
  API例 = 10
  説明 = 11
  種類数 = 12
  区分値数 = 13
End Enum

Const MAX_COL = 13

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

Dim val As Variant '変更前の仮値

'シート選択時
Private Sub Worksheet_Activate()
  '// フォーマット
  Call areaFormat
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

'// データLoad(From yaml)
Public Sub load(ByVal data As String)
  Dim records() As String, items() As String
  Dim i As Long, rowLength As Long

  '// 初期化
  UsedRange.Offset(1).Clear
  Columns(col_no.API例).NumberFormatLocal = "@"

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
  Cells(2, col_no.項目名J).Resize(rowLength, 2) = wkRange1
  Cells(2, col_no.ドメイン).Resize(rowLength, 8) = wkRange2

  '// フォーマット
  Call format2
End Sub

'// フォーマット TODO:リネーム
Public Sub format2()
  Dim row_length As Long: row_length = Cells(rows.Count, "A").End(xlUp).Row - 1
  
  Call areaFormat
  Call format_domain_type(Cells(2, col_no.ドメイン).Resize(row_length), False)
End Sub

'フォーマット初期化
Public Sub areaFormat()
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
  Call Formatter.conditionFormat(Union(Range("ELEM_B"), Range("ELEM_D"), Range("ELEM_J"), Range("ELEM_K")), "=AND(RC1<>"""",RC="""")")
  
  '// IMEモード
  Call Formatter.imeMode(Union(Range("ELEM_A"), Range("ELEM_K")), Formatter.IME_MODE_ON)
  Call Formatter.imeMode(Range("ELEM_B"), Formatter.IME_MODE_OFF)

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
  If is_reset Then Target.Offset(, 1).Resize(, 5).ClearContents
  
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
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.最小値).Resize(, 2))
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
      .Validation.imeMode = xlIMEModeOff
      .NumberFormatLocal = "@"
    
    End With
    
  ElseIf Target.dec_type = decoration_type.桁数 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1000"
      .Validation.ErrorMessage = "0-1000の間で入力可能"
      .Validation.imeMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.数値 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-99999999", Formula2:="99999999"
      .Validation.ErrorMessage = "数値のみ入力可能"
      .Validation.imeMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.真偽値 Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1"
      .Validation.ErrorMessage = "1(値を含む)のみ入力可能"
      .Validation.imeMode = xlIMEModeOff
      If is_reset Then .value = 1
    
    End With
    
  ElseIf Target.dec_type = decoration_type.入力不可 Then
    With Target.dec_range
      .Interior.Color = work.Range("入力不可").Interior.Color
      .Validation.Delete
      .Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="0"
      .Validation.ErrorMessage = "入力不可項目です"
      .Validation.imeMode = xlIMEModeOff
    
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
