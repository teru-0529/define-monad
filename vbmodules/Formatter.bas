Attribute VB_Name = "Formatter"
Option Explicit

Public Const IME_MODE_ON = True
Public Const IME_MODE_OFF = False

'// 枠線・文字サイズ
Sub borderAndFontsize(ByRef area As Range)
  With area
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(68, 114, 196) 'DEEP_BLUE
    .Borders(xlEdgeTop).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Borders(xlInsideHorizontal).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Font.Size = 9
  End With
End Sub

'// リスト選択
Sub listSelect(ByRef area As Range, ByVal listName As String)
  Dim condFormat As FormatCondition

  '入力規則
  With area.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
  End With

  '条件付き書式
  Call conditionFormat(area, "=AND(RC<>"""",ISNA(MATCH(RC," & listName & ",0)))")
End Sub

'// 条件付き書式
Sub conditionFormat(ByRef area As Range, ByVal formula As String)
  Dim condFormat As FormatCondition

  Set condFormat = area.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
  condFormat.Interior.Color = work.Range("エラー").Interior.Color
End Sub

'// IMEモード
Sub imeMode(ByRef area As Range, ByVal mode As Boolean)
  With area.Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    If mode Then
      .imeMode = xlIMEModeOn
    Else
      .imeMode = xlIMEModeOff
    End If
  End With
End Sub

  
