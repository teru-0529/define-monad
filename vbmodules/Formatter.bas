Attribute VB_Name = "Formatter"
Option Explicit

Public Const IME_MODE_ON = True
Public Const IME_MODE_OFF = False

'// �g���E�����T�C�Y
Sub borderAndFontsize(ByRef area As Range)
  With area
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(68, 114, 196) 'DEEP_BLUE
    .Borders(xlEdgeTop).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Borders(xlInsideHorizontal).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Font.Size = 9
  End With
End Sub

'// ���X�g�I��
Sub listSelect(ByRef area As Range, ByVal listName As String)
  Dim condFormat As FormatCondition

  '���͋K��
  With area.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
  End With

  '�����t������
  Call conditionFormat(area, "=AND(RC<>"""",ISNA(MATCH(RC," & listName & ",0)))")
End Sub

'// �����t������
Sub conditionFormat(ByRef area As Range, ByVal formula As String)
  Dim condFormat As FormatCondition

  Set condFormat = area.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
  condFormat.Interior.Color = work.Range("�G���[").Interior.Color
End Sub

'// IME���[�h
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

  
