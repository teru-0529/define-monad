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

'�J�����ԍ�
Private Enum col_no
  �h�������ږ� = 1
  ���ږ�J = 2
  ���ږ�EC = 3
  ���ږ�ES = 4
  ���� = 5
End Enum

Const MAX_COL = 5


'�V�[�g�I����
Private Sub Worksheet_Activate()
  Call work.element_reflesh '���ڃ��X�g�̍ŐV��

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

  '�w�i�F
  Cells(start_row, col_no.���ږ�ES).Resize(f_rows).Interior.Color = RGB(255, 255, 213) 'CREAM
  
  '���͋K��
  With Cells(start_row, col_no.�h�������ږ�).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=���ڃ��X�g"
  End With

  '�֐�
  Cells(start_row, col_no.���ږ�ES).Resize(f_rows).Formula = "=camel2snake(RC[-1])"
  
  '�����t������
  Set cond_format_1 = Cells(start_row, col_no.�h�������ږ�).Resize(f_rows).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC<>"""",ISNA(MATCH(RC,���ڃ��X�g,0)))")
  cond_format_1.Interior.Color = work.Range("���͕K�{").Interior.Color
  Set cond_format_2 = Union(Cells(start_row, col_no.���ږ�J).Resize(f_rows), Cells(start_row, col_no.���ږ�EC).Resize(f_rows), _
   Cells(start_row, col_no.����).Resize(f_rows)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC1<>"""",RC="""")")
  cond_format_2.Interior.Color = work.Range("���͕K�{").Interior.Color
  
  'IME���[�h
  With Union(Cells(start_row, col_no.���ږ�J).Resize(f_rows), Cells(start_row, col_no.����).Resize(f_rows)).Validation
    .Delete
    .Add Type:=xlValidateInputOnly
    .IMEMode = xlIMEModeOn
  End With
  With Cells(start_row, col_no.���ږ�EC).Resize(f_rows).Validation
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
  ReDim wk_range_1(row_length - 1, 2)
  ReDim wk_range_2(row_length - 1, 0)
  
  '�`��ӏ��̃t�H�[�}�b�g
  Call init_format(2, row_length)
  
  For i = 0 To row_length - 1
    wk_range_1(i, 0) = from_yaml(data(i * PARAMS + 0)) '�h����
    wk_range_1(i, 1) = from_yaml(data(i * PARAMS + 1)) '����_�a
    wk_range_1(i, 2) = from_yaml(data(i * PARAMS + 2)) '����_�p
    
    wk_range_2(i, 0) = from_yaml(data(i * PARAMS + 3)) '����
  Next i
  
  '�l�Z�b�g
  Cells(2, col_no.�h�������ږ�).Resize(row_length, 3) = wk_range_1
  Cells(2, col_no.����).Resize(row_length) = wk_range_2
  
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
  
  Call push_array(data, indent(1) & "- " & to_yaml("origin", Cells(i, col_no.�h�������ږ�).value))
  Call push_array(data, indent(2) & to_yaml("name_jp", Cells(i, col_no.���ږ�J).value))
  Call push_array(data, indent(2) & to_yaml("name_en", Cells(i, col_no.���ږ�EC).value))
  Call push_array(data, indent(2) & to_yaml("description", Cells(i, col_no.����).value))
  Call push_array(data, "")
End Sub


'table_elements����
Public Sub out_table(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call table_item(data, elm)
  Next elm
  
End Sub

Private Sub table_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  '���ڃV�[�g�̔h��������擾
  Dim domain As String: domain = elements.get_val(Cells(i, col_no.�h�������ږ�).value, 3, "")
  Dim maxlen As String: maxlen = elements.get_val(Cells(i, col_no.�h�������ږ�).value, 6, "")
  Dim name_en As String: name_en = elements.get_val(Cells(i, col_no.�h�������ږ�).value, 2, "")
  Dim constraint As String: constraint = elements.db_constraint_by_name(Cells(i, col_no.�h�������ږ�).value, Cells(i, col_no.���ږ�ES).value)
  
  Call push_array(data, "- " & to_yaml("name_jp", Cells(i, col_no.���ږ�J).value))
  Call push_array(data, indent(1) & to_yaml("name_en", Cells(i, col_no.���ږ�ES).value))
  If domain = "�敪�l" Then
    Call push_array(data, indent(1) & to_yaml("db_model", name_en))
  Else
    Call push_array(data, indent(1) & to_yaml("db_model", domain_types.get_db_model(domain, maxlen)))
  End If
  Call push_array(data, indent(1) & to_yaml("constraint", constraint))
  Call push_array(data, indent(1) & to_yaml("must_not_null", must_not_null(domain)))
  Call push_array(data, indent(1) & to_yaml("is_string_default", is_string_default(domain)))
  Call push_array(data, indent(1) & to_yaml("description", Cells(i, col_no.����).value))
  Call push_array(data, indent(1) & to_yaml("is_origin", "false"))
  Call push_array(data, indent(1) & to_yaml("origin", Cells(i, col_no.�h�������ږ�).value))
  Call push_array(data, "")
End Sub


'derive_elements.tsv(�\���p)�o��
Public Sub out_view(ByVal line_sep As Integer)
  Dim data() As String, elm As Range, i As Long
  
  Call push_array(data, array2tabstr("�h����", "����(JP)", "����(EN)", "����"))
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then
      i = elm.Row
      Call push_array(data, array2tabstr( _
        Cells(i, col_no.�h�������ږ�).value, _
        Cells(i, col_no.���ږ�J).value, _
        Cells(i, col_no.���ږ�EC).value, _
        Cells(i, col_no.����).value & " "))
    End If
  Next elm
  
  Call text_io.plain_out(settings.view_derive_elements_path(), data, line_sep)
End Sub





