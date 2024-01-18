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

'�J�����ԍ�
Private Enum col_no
  ���ږ�J = 1
  ���ږ�EC = 2
  ���ږ�ES = 3
 �h���C�� = 4
  ���K�\�� = 5
  �ŏ����� = 6
  �ő包�� = 7
  �ŏ��l = 8
  �ŏ��l�� = 9
  �ő�l = 10
  �ő�l�� = 11
  API�� = 12
  ���� = 13
  API�^1 = 14
  API�^2 = 15
  ��ސ� = 16
  �敪�l�� = 17
End Enum

'�f�R���[�V�������
Private Enum decoration_type
  ������ = 1
  ���� = 2
  ���l = 3
  �^�U�l = 4
  ���͕s�� = 9
  ���͕K�{ = 100
End Enum

'�Z���ݒ���
Private Type decolation_cells
  dec_range As Range
  dec_type As decoration_type
  reqired As Boolean
End Type
  
  
Const MAX_COL = 17


Dim val As Variant '�ύX�O�̉��l

'�V�[�g�I����
Private Sub Worksheet_Activate()
  Dim start_row As Integer: start_row = Cells(rows.Count, "A").End(xlUp).Row
  Call init_format(start_row + 1, 0)
End Sub

'�Z���̑I���C�x���g
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error GoTo ENDUP
  
  '���ږ���I�������ꍇ
  If Target.Column = ���ږ�J Then
    val = ""
    If Target.value = "" Then Exit Sub
    If Target.Cells.Count <> 1 Then Exit Sub
  
    val = Target.value
  End If

ENDUP:
End Sub

'�Z���̕ύX�C�x���g
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo ENDUP
  
  If Target.Cells.Count <> 1 Then Exit Sub
  
  '���ږ��̕ύX�̏ꍇ
  If Target.Column = col_no.���ږ�J Then
    If val = "" Then Exit Sub
    If Target.value = "" Then Exit Sub
  
    If Target.value <> val Then
      Call derive_elements.replace_element_name(val, Target.value)
      Call segments.replace_element_name(val, Target.value)
    End If
    Exit Sub
  
ENDUP:
  
  '�h���C���̕ύX�̏ꍇ
  ElseIf Target.Column = col_no.�h���C�� Then
    '�l���Z�b�g
    Call format_domain_type(Target, True)
  End If

End Sub

'�t�H�[�}�b�g������
Public Sub init_format(ByVal start_row As Long, ByVal rows As Long)
  Application.EnableEvents = False
  
  Dim f_rows As Integer: f_rows = rows + ADDED_ROWS '�t�H�[�}�b�g���s��(ADDED_ROWS�����߂Ƀt�H�[�}�b�g)
  Dim cond_format As FormatCondition

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
  Union(Cells(start_row, col_no.���ږ�ES).Resize(f_rows), Cells(start_row, col_no.API�^1).Resize(f_rows, 4)).Interior.Color = RGB(255, 255, 213) 'CREAM
  
  '���͋K��
  With Cells(start_row, col_no.�h���C��).Resize(f_rows).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=�h���C���^�C�vKEY"
  End With

  '�֐�
  Cells(start_row, col_no.���ږ�ES).Resize(f_rows).Formula = "=camel2snake(RC[-1])"
  Cells(start_row, col_no.API�^1).Resize(f_rows).Formula = "=iferror(vlookup(RC[-10],�h���C���^�C�vDATA,6,False)&"""","""")"
  Cells(start_row, col_no.API�^2).Resize(f_rows).Formula = "=iferror(vlookup(RC[-11],�h���C���^�C�vDATA,7,False)&"""","""")"
  Cells(start_row, col_no.��ސ�).Resize(f_rows).Formula = "=countif(�h��������KEY,RC[-15])"
  Cells(start_row, col_no.�敪�l��).Resize(f_rows).Formula = "=countif(�敪�l����KEY,RC[-16])"
  
  '�ʒu
  Union(Cells(start_row, col_no.�ŏ��l��).Resize(f_rows), Cells(start_row, col_no.�ő�l��).Resize(f_rows)).HorizontalAlignment = xlCenter

  '�\���`��
  Union(Cells(start_row, col_no.�ŏ��l��).Resize(f_rows), Cells(start_row, col_no.�ő�l��).Resize(f_rows)).NumberFormatLocal = """��"";;;"
  Cells(start_row, col_no.API��).Resize(f_rows).NumberFormatLocal = "@"
  Union(Cells(start_row, col_no.��ސ�).Resize(f_rows), Cells(start_row, col_no.�敪�l��).Resize(f_rows)).NumberFormatLocal = "0_;;;"

  
  '�����t������
  Set cond_format = Union(Cells(start_row, col_no.���ږ�EC).Resize(f_rows), Cells(start_row, col_no.�h���C��).Resize(f_rows), _
   Cells(start_row, col_no.API��).Resize(f_rows), Cells(start_row, col_no.����).Resize(f_rows)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC1<>"""",RC="""")")
  cond_format.Interior.Color = work.Range("���͕K�{").Interior.Color
  
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


'�h���C���^�C�v�̕ύX�C�x���g
Private Sub format_domain_type(ByVal Target As Range, ByVal is_reset As Boolean)
  Dim ng As decolation_cells:  ng.dec_type = decoration_type.���͕s��:  ng.reqired = False
  Dim str As decolation_cells:  str.dec_type = decoration_type.������:  str.reqired = False
  Dim str_req As decolation_cells:  str_req.dec_type = decoration_type.������:  str_req.reqired = True
  Dim keta As decolation_cells:  keta.dec_type = decoration_type.����:  keta.reqired = False
  Dim keta_req As decolation_cells:  keta_req.dec_type = decoration_type.����:  keta_req.reqired = True
  Dim num As decolation_cells:  num.dec_type = decoration_type.���l:  num.reqired = True
  Dim bool As decolation_cells:  bool.dec_type = decoration_type.�^�U�l: bool.reqired = False
  
  Dim elm As Range
  Dim i As Long
  Dim dec As Integer
  
  '�������i�l���Z�b�g�j
  If is_reset Then Target.Offset(, 1).Resize(, 7).ClearContents
  
  '�s���ƂɃt�H�[�}�b�g�ΏƂ�I��
  For Each elm In Target
    i = elm.Row
    
    '���K�\��
    dec = domain_types.get_val(elm.value, 7, -1)
    If dec > 0 Then
      Set str_req.dec_range = Util.union_range(str_req.dec_range, Cells(i, col_no.���K�\��))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.���K�\��))
    Else
      Set str.dec_range = Util.union_range(str.dec_range, Cells(i, col_no.���K�\��))
    End If
    '�^�C�v�F���Ԃ̏ꍇ�A���K�\����ݒ�
    If elm.value = "����" Then Cells(i, col_no.���K�\��).value = "^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$"
    
    '�ŏ�����
    dec = domain_types.get_val(elm.value, 8, -1)
    If dec > 0 Then
      Set keta_req.dec_range = Util.union_range(keta_req.dec_range, Cells(i, col_no.�ŏ�����))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.�ŏ�����))
    Else
      Set keta.dec_range = Util.union_range(keta.dec_range, Cells(i, col_no.�ŏ�����))
    End If
    
    '�ő包��
    dec = domain_types.get_val(elm.value, 9, -1)
    If dec > 0 Then
      Set keta_req.dec_range = Util.union_range(keta_req.dec_range, Cells(i, col_no.�ő包��))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.�ő包��))
    End If
    
    '�ŏ��l/�ő�l
    dec = domain_types.get_val(elm.value, 10, -1)
    If dec > 0 Then
      Set num.dec_range = Util.union_range(num.dec_range, Cells(i, col_no.�ŏ��l))
      Set num.dec_range = Util.union_range(num.dec_range, Cells(i, col_no.�ő�l))
      Set bool.dec_range = Util.union_range(bool.dec_range, Cells(i, col_no.�ŏ��l��))
      Set bool.dec_range = Util.union_range(bool.dec_range, Cells(i, col_no.�ő�l��))
    ElseIf dec < 0 Then
      Set ng.dec_range = Util.union_range(ng.dec_range, Cells(i, col_no.�ŏ��l).Resize(, 4))
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


'���̓G���A�̃f�R���[�V����
Private Sub decoration(ByRef Target As decolation_cells, ByVal is_reset As Boolean)
  Dim cond_format As FormatCondition
  
  If Target.dec_range Is Nothing Then Exit Sub
      
      
  If Target.dec_type = decoration_type.������ Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateInputOnly
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "@"
    
    End With
    
  ElseIf Target.dec_type = decoration_type.���� Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1000"
      .Validation.ErrorMessage = "0-1000�̊Ԃœ��͉\"
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.���l Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-99999999", Formula2:="99999999"
      .Validation.ErrorMessage = "���l�̂ݓ��͉\"
      .Validation.IMEMode = xlIMEModeOff
      .NumberFormatLocal = "#,##0_ "
    
    End With
    
  ElseIf Target.dec_type = decoration_type.�^�U�l Then
    With Target.dec_range
      .Interior.ColorIndex = xlNone
      .Validation.Delete
      .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1"
      .Validation.ErrorMessage = "1(�l���܂�)�̂ݓ��͉\"
      .Validation.IMEMode = xlIMEModeOff
      If is_reset Then .value = 1
    
    End With
    
  ElseIf Target.dec_type = decoration_type.���͕s�� Then
    With Target.dec_range
      .Interior.Color = work.Range("���͕s��").Interior.Color
      .Validation.Delete
      .Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="0"
      .Validation.ErrorMessage = "���͕s���ڂł�"
      .Validation.IMEMode = xlIMEModeOff
    
    End With
  End If
  
  '�K�{
  If Target.reqired Then
    Set cond_format = Target.dec_range.FormatConditions.Add(xlBlanksCondition)
    cond_format.Interior.Color = work.Range("���͕K�{").Interior.Color
  
  Else
    Target.dec_range.FormatConditions.Delete
  
  End If
End Sub

'���ږ����w�肵�ėv�f�l��Ԃ��B
Public Function get_val(ByVal name As String, ByVal col As Integer, ByVal def_val) As Variant
  Dim elm As Range: Set elm = Range("���ږ�KEY").Find(name, LookAt:=xlWhole)
  
  '���݂��Ȃ���΃f�t�H���g�l��ԋp
  If elm Is Nothing Then
    get_val = def_val
  
  Else
    get_val = elm.Offset(, col).value
  
  End If
  
End Function


'������������IO�@����������

'save_data�Ǎ�
Public Sub load_(ByRef data() As String)
  Const PARAMS = 13
  Dim row_length As Long: row_length = (UBound(data) + 1) / PARAMS
  Dim i As Long
  ReDim wk_range_1(row_length - 1, 1)
  ReDim wk_range_2(row_length - 1, 9)
  
  '�`��ӏ��̃t�H�[�}�b�g
  Call init_format(2, row_length)
  
  For i = 0 To row_length - 1
    wk_range_1(i, 0) = from_yaml(data(i * PARAMS + 0))  '����_�a
    wk_range_1(i, 1) = from_yaml(data(i * PARAMS + 1))  '����_�p
    
    wk_range_2(i, 0) = from_yaml(data(i * PARAMS + 2))  '�h���C��
    wk_range_2(i, 1) = from_yaml(data(i * PARAMS + 3))  '���K�\��
    wk_range_2(i, 2) = from_yaml(data(i * PARAMS + 4))  '�ŏ�����
    wk_range_2(i, 3) = from_yaml(data(i * PARAMS + 5))  '�ő包��
    wk_range_2(i, 4) = from_yaml(data(i * PARAMS + 6))  '�ŏ��l
    wk_range_2(i, 5) = from_yaml(data(i * PARAMS + 7))  '�ŏ��l�܂�
    wk_range_2(i, 6) = from_yaml(data(i * PARAMS + 8))  '�ő�l
    wk_range_2(i, 7) = from_yaml(data(i * PARAMS + 9))  '�ő�l�܂�
    wk_range_2(i, 8) = from_yaml(data(i * PARAMS + 10))  '��
    wk_range_2(i, 9) = from_yaml(data(i * PARAMS + 11))  '����
  Next i
  
  '�l�Z�b�g
  Cells(2, col_no.���ږ�J).Resize(row_length, 2) = wk_range_1
  Cells(2, col_no.�h���C��).Resize(row_length, 10) = wk_range_2
  
  'format
  Call format_domain_type(Cells(2, col_no.�h���C��).Resize(row_length), False)
 
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
  
  Call push_array(data, indent(1) & "- " & to_yaml("name_jp", Cells(i, col_no.���ږ�J).value))
  Call push_array(data, indent(2) & to_yaml("name_en", Cells(i, col_no.���ږ�EC).value))
  Call push_array(data, indent(2) & to_yaml("domain", Cells(i, col_no.�h���C��).value))
  Call push_array(data, indent(2) & to_yaml("reg_ex", Cells(i, col_no.���K�\��).value))
  Call push_array(data, indent(2) & to_yaml("min_digits", Cells(i, col_no.�ŏ�����).value))
  Call push_array(data, indent(2) & to_yaml("max_digits", Cells(i, col_no.�ő包��).value))
  Call push_array(data, indent(2) & to_yaml("min_value", Cells(i, col_no.�ŏ��l).value))
  Call push_array(data, indent(2) & to_yaml("min_value_boundory", Cells(i, col_no.�ŏ��l��).value))
  Call push_array(data, indent(2) & to_yaml("max_value", Cells(i, col_no.�ő�l).value))
  Call push_array(data, indent(2) & to_yaml("max_value_boundory", Cells(i, col_no.�ő�l��).value))
  Call push_array(data, indent(2) & to_yaml("example", Cells(i, col_no.API��).value))
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
  
  Call push_array(data, "- " & to_yaml("name_jp", Cells(i, col_no.���ږ�J).value))
  Call push_array(data, indent(1) & to_yaml("name_en", Cells(i, col_no.���ږ�ES).value))
  If Cells(i, col_no.�h���C��).value = "�敪�l" Then
    Call push_array(data, indent(1) & to_yaml("db_model", Cells(i, col_no.���ږ�ES).value))
  Else
    Call push_array(data, indent(1) & to_yaml("db_model", domain_types.get_db_model(Cells(i, col_no.�h���C��).value, Cells(i, col_no.�ő包��).value)))
  End If
  Call push_array(data, indent(1) & to_yaml("constraint", db_constraint(elm, Cells(i, col_no.���ږ�ES).value)))
  Call push_array(data, indent(1) & to_yaml("must_not_null", must_not_null(Cells(i, col_no.�h���C��).value)))
  Call push_array(data, indent(1) & to_yaml("is_string_default", is_string_default(Cells(i, col_no.�h���C��).value)))
  Call push_array(data, indent(1) & to_yaml("description", Cells(i, col_no.����).value))
  Call push_array(data, indent(1) & to_yaml("is_origin", "true"))
  Call push_array(data, indent(1) & to_yaml("origin", ""))
  Call push_array(data, "")
End Sub


'api_elements����
Public Sub out_api(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then Call api_item(data, elm)
  Next elm
  
End Sub

Private Sub api_item(ByRef data() As String, ByVal elm As Range)
  Dim i As Long: i = elm.Row
  
  Call push_array(data, "# " & Cells(i, col_no.���ږ�J).value)
  Call push_array(data, "# : " & Cells(i, col_no.����).value)
  Call push_array(data, Cells(i, col_no.���ږ�EC).value & ":")
  Call push_array(data, indent(1) & to_yaml("type", Cells(i, col_no.API�^1).value))
  If Cells(i, col_no.API�^2).value <> "" Then Call push_array(data, indent(1) & to_yaml("format", Cells(i, col_no.API�^2).value))
  If Cells(i, col_no.���K�\��).value <> "" Then Call push_array(data, indent(1) & to_yaml("pattern", Cells(i, col_no.���K�\��).value))
  If Cells(i, col_no.�ŏ�����).value <> "" Then Call push_array(data, indent(1) & to_yaml("minLength", Cells(i, col_no.�ŏ�����).value))
  If Cells(i, col_no.�ő包��).value <> "" Then Call push_array(data, indent(1) & to_yaml("maxLength", Cells(i, col_no.�ő包��).value))
  If Cells(i, col_no.�ŏ��l).value <> "" Then Call push_array(data, indent(1) & to_yaml("minimum", Cells(i, col_no.�ŏ��l).value))
  If Cells(i, col_no.�ő�l).value <> "" Then Call push_array(data, indent(1) & to_yaml("maximum", Cells(i, col_no.�ő�l).value))
  If Cells(i, col_no.�ŏ��l��).value <> "" Then Call push_array(data, indent(1) & "exclusiveMinimum: true")
  If Cells(i, col_no.�ő�l��).value <> "" Then Call push_array(data, indent(1) & "exclusiveMaximum: true")
  If Cells(i, col_no.�h���C��).value = "�敪�l" Then
    Call push_array(data, indent(1) & "enum:")
    Call segments.api_value(data, Cells(i, col_no.���ږ�J).value, indent(2) & "- ")
  End If
  Call push_array(data, indent(1) & to_yaml("description", Cells(i, col_no.����).value))
  If Cells(i, col_no.API�^1).value = "string" Then
    Call push_array(data, indent(1) & to_yaml("example", "'" & Cells(i, col_no.API��).value) & "'")
  Else
    Call push_array(data, indent(1) & to_yaml("example", Cells(i, col_no.API��).value))
  End If
  Call push_array(data, "")
End Sub


'type_ddl����
Public Sub out_type_ddl(ByRef data() As String)
  Dim elm As Range
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    Dim i As Long: i = elm.Row
    If Cells(i, col_no.�h���C��).value = "�敪�l" Then
      Call push_array(data, "-- " & Cells(i, col_no.���ږ�J).value)
      Call push_array(data, "DROP TYPE IF EXISTS " & Cells(i, col_no.���ږ�ES).value & ";")
      Call push_array(data, "CREATE TYPE " & Cells(i, col_no.���ږ�ES).value & " AS enum (")
      Call segments.ddl_value(data, Cells(i, col_no.���ږ�J).value, indent(1))
      Call push_array(data, ");")
      Call push_array(data, "")
    End If
  Next elm
  
End Sub


'elements.tsv(�\���p)�o��
Public Sub out_view(ByVal line_sep As Integer)
  Dim data() As String, elm As Range, i As Long
  
  Call push_array(data, array2tabstr("����(JP)", "����(EN)", "�h���C��", "API(type)", "API(format)", "���K�\��", "�ŏ�����", "�ő包��", "�ŏ��l", "�ŏ��l�܂�", "�ő�l", "�ő�l�܂�", "API(example)", "����"))
  
  For Each elm In UsedRange.Resize(, 1).Offset(1)
    If elm.value <> "" Then
      i = elm.Row
      Call push_array(data, array2tabstr( _
        Cells(i, col_no.���ږ�J).value, _
        Cells(i, col_no.���ږ�EC).value, _
        Cells(i, col_no.�h���C��).value, _
        Cells(i, col_no.API�^1).value, _
        Cells(i, col_no.API�^2).value, _
        Cells(i, col_no.���K�\��).value, _
        Cells(i, col_no.�ŏ�����).value, _
        Cells(i, col_no.�ő包��).value, _
        Cells(i, col_no.�ŏ��l).value, _
        change_boundary(Cells(i, col_no.�ŏ��l��).value), _
        Cells(i, col_no.�ő�l).value, _
        change_boundary(Cells(i, col_no.�ő�l��).value), _
        Cells(i, col_no.API��).value, _
        Cells(i, col_no.����).value & " "))
    End If
  Next elm
  
  Call text_io.plain_out(settings.view_elements_path(), data, line_sep)
End Sub

Private Function change_boundary(ByVal val As Variant) As Variant
  If val = 1 Then change_boundary = "��" Else change_boundary = "-"
End Function

'DB���񕶎�����\��
Private Function db_constraint(Target As Range, field_name As String) As String
  Dim le As String, ge As String
  
  If Target.Offset(, 4).value <> "" Then '���K�\���̏ꍇ
    db_constraint = "(" & field_name & " ~* '" & Target.Offset(, 4).value & "')"
  
  ElseIf Target.Offset(, 5).value <> "" Then '�Œᕶ���񒷂̏ꍇ
    db_constraint = "(LENGTH(" & field_name & ") >= " & Target.Offset(, 5).value & ")"
  
  ElseIf Target.Offset(, 7).value <> "" Then '���l�͈͂̏ꍇ
    If Target.Offset(, 8).value = 1 Then le = " >= " Else le = " > "
    If Target.Offset(, 10).value = 1 Then ge = " <= " Else ge = " < "
    db_constraint = "(" & field_name & le & Target.Offset(, 7).value & " AND " & field_name & ge & Target.Offset(, 9).value & ")"
  
  Else
    db_constraint = ""
  End If
End Function

'DB���񕶎�����\��(�ʖ��p)
Public Function db_constraint_by_name(name As String, field_name As String) As String
  Dim elm As Range: Set elm = Range("���ږ�KEY").Find(name, LookAt:=xlWhole)
  
  '���݂��Ȃ���΃f�t�H���g�l��ԋp
  If elm Is Nothing Then
    db_constraint_by_name = ""
  
  Else
    db_constraint_by_name = db_constraint(elm, field_name)
  
  End If
End Function
