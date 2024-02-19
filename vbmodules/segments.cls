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

Const COL_ORIGINAL = 1

'シート選択時
Private Sub Worksheet_Activate()
  Application.ScreenUpdating = False
  Call work.segment_reflesh '区分値リストの最新化
  
  '// フォーマット
  Call areaFormat
  Application.ScreenUpdating = True
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

  '// 入力値が存在しない場合、フォーマットのみ実施
  If data = "" Then
    '// フォーマット
    Call areaFormat
    Exit Sub
  End If
    
  '// 行ごとのデータに分割
  records = Split(data, vbLf)
  rowLength = UBound(records)
  ReDim wkRange1(rowLength - 1, 3)
  
  '// 列ごとの配列に変換
  For i = 0 To rowLength - 1
    items = Split(records(i), vbTab)

    wkRange1(i, 0) = items(0)  '項目名
    wkRange1(i, 1) = items(1)  '区分値
    wkRange1(i, 2) = items(2)  '区分名
    wkRange1(i, 3) = items(3)  '説明
  Next i

  '// 値セット
  Cells(2, COL_ORIGINAL).Resize(rowLength, 4) = wkRange1

  '// フォーマット
  Call areaFormat
End Sub

'// フォーマット
Private Sub areaFormat()
  Application.EnableEvents = False

  '// 枠線
  Call Formatter.borderAndFontsize(Range("SEGM_FULL"))

  '// リスト選択
  Call Formatter.listSelect(Range("SEGM_A"), "区分値リスト")

  '// 条件付き書式
  Call Formatter.condFunction(Union(Range("SEGM_B"), Range("SEGM_C"), Range("SEGM_D")), "=AND(RC1<>"""",RC="""")")

  '// IMEモード
  Call Formatter.imeMode(Union(Range("SEGM_C"), Range("SEGM_D")), Formatter.IME_MODE_ON)
  Call Formatter.imeMode(Range("SEGM_B"), Formatter.IME_MODE_OFF)

  Application.EnableEvents = True
End Sub
