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

'�J�����ԍ�
Private Enum col_no
  ���ږ�J = 1
  �敪�l = 2
  �敪�� = 3
  ���� = 4
End Enum

Const MAX_COL = 4


'�V�[�g�I����
Private Sub Worksheet_Activate()
  Call work.segment_reflesh '�敪�l���X�g�̍ŐV��

  Dim start_row As Integer: start_row = Cells(rows.Count, "A").End(xlUp).Row
  Call init_format(start_row + 1, 0)
End Sub


'���ږ��̕ύX�ǐ�
Public Sub replace_element_name(ByVal old_val As String, ByVal new_val As String)
  UsedRange.Resize(, 1).Replace what:=old_val, replacement:=new_val
End Sub


'�t�H�[�}�b�g������
Public Sub init_format(ByVal start_row As Long, ByVal rows As Long)
  Application.EnableEvents = False
  
  Dim f_rows As Integer: f_rows = rows + ADDED_ROWS '�t�H�[�}�b�g���s��(ADDED_ROWS�����߂Ƀt�H�[�}�b�g)
  Dim cond_format_1 As FormatCondition, cond_format_2 As FormatCondition

  '�����s�ȍ~���N���A
  UsedRange.Offset(start_row - 1).Clear
  
  '�g��
  With Cells(start_row, 1).Resize(f_rows, MAX_COL)
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(68, 114, 196) 'DEEP_BLUE
    .Borders(xlEdgeTop).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Borders(xlInsideHorizontal).Color = RGB(189, 215, 238) 'WEAK_BLUE
    .Font.Size = 9
  End With

  '���͋K��
  With Cells(start_row, col_no.���ږ�J).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=�敪�l���X�g"
  End With

  '�����t������
  Set cond_format_1 = Cells(start_row, col_no.���ږ�J).Resize(f_rows).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC<>"""",ISNA(MATCH(RC,�敪�l���X�g,0)))")
  cond_format_1.Interior.Color = work.Range("���͕K�{").Interior.Color
  Set cond_format_2 = Union(Cells(start_row, col_no.�敪�l).Resize(f_rows), _
   Cells(start_row, col_no.�敪��).Resize(f_rows), Cells(start_row, col_no.����).Resize(f_rows)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC1<>"""",RC="""")")
  cond_format_2.Interior.Color = work.Range("���͕K�{").Interior.Color
  
  'IME���[�h
  With Union(Cells(start_row, col_no.�敪��).Resize(f_rows), Cells(start_row, col_no.����).Resize(f_rows)).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOn
  End With
  With Cells(start_row, col_no.�敪�l).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOff
  End With
  
  Application.EnableEvents = True
End Sub


'������������IO�@����������

'save_data�Ǎ�
Public Sub load_(ByRef data() As String)
  Const PARAMS = 5
  Dim row_length As Long: row_length = (UBound(data) + 1) / PARAMS
  Dim i As Long
  ReDim wk_range_1(row_length - 1, 3)
  
  '�`��ӏ��̃t�H�[�}�b�g
  Call init_format(2, row_length)
  
  For i = 0 To row_length - 1
    wk_range_1(i, 0) = from_yaml(data(i * PARAMS + 0)) '���ږ�
    wk_range_1(i, 1) = from_yaml(data(i * PARAMS + 1)) '�敪�l
    wk_range_1(i, 2) = from_yaml(data(i * PARAMS + 2)) '�敪��
    wk_range_1(i, 3) = from_yaml(data(i * PARAMS + 3)) '����
  Next i
  
  '�l�Z�b�g
  Cells(2, col_no.���ږ�J).Resize(row_length, 4) = wk_range_1
  
End Sub


'save_data����
Public Sub save_(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call save_item(data, elm)
  Next elm
  
End Sub

Private Sub save_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  Call push_array(data, indent(1) & "- " & to_yaml("key", Cells(i, col_no.���ږ�J).value))
  Call push_array(data, indent(2) & to_yaml("value", Cells(i, col_no.�敪�l).value))
  Call push_array(data, indent(2) & to_yaml("name", Cells(i, col_no.�敪��).value))
  Call push_array(data, indent(2) & to_yaml("description", Cells(i, col_no.����).value))
  Call push_array(data, "")
End Sub


'API�p�����[�^�̏o��
Public Sub api_value(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long:
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(data, indent & Cells(i, col_no.�敪�l).value)
    '##TODO: enum��`�̍ۃV���O���N�H�[�g���K�v���ǂ���
    'If elm.value = name Then Call push_array(data, indent & "'" & Cells(i, col_no.�敪�lC).value & "'")
  Next elm
End Sub


'API�p�����[�^�̏o��
Public Sub api_desc(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long:
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(data, indent & Cells(i, col_no.�敪�l).value & " : " & Cells(i, col_no.�敪��).value)
  Next elm
End Sub


'DDL�p�����[�^�̏o��
Public Sub ddl_value(ByRef data() As String, ByVal name As String, ByVal indent As String)
  Dim elm As Range, i As Long
  Dim wk() As String
  
  '��xwk�ɓ����i�Ōゾ���J���}�������Ȃ����߁j
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    i = elm.Row
    If elm.value = name Then Call push_array(wk, indent & "'" & Cells(i, col_no.�敪�l).value & "'")
  Next elm
  
  For i = 0 To UBound(wk) - 1
    Call push_array(data, wk(i) & ",")
  Next i
  Call push_array(data, wk(UBound(wk)))
  
End Sub


'segments.tsv(�\���p)�o��
Public Sub out_view(ByVal line_sep As Integer)
  Dim data() As String, elm As Range, i As Long
  
  Call push_array(data, array2tabstr("���ږ�", "�敪�l", "�敪��", "����"))
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then
      i = elm.Row
      Call push_array(data, array2tabstr( _
        Cells(i, col_no.���ږ�J).value, _
        Cells(i, col_no.�敪�l).value, _
        Cells(i, col_no.�敪��).value, _
        Cells(i, col_no.����).value & " "))
    End If
  Next elm
  
  Call text_io.plain_out(settings.view_segments_path(), data, line_sep)
End Sub