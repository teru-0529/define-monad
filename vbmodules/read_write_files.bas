Attribute VB_Name = "read_write_files"
Option Explicit

'save_data.yaml.����
'// FIXME:��������Refactering to Backend
Public Sub load_(ByVal line_sep As Integer)
  Dim data() As String, row_length As Long, i As Long, temp As Variant
  
  Dim elements_d() As String, derive_elements_d() As String, segments_d() As String
  Dim s1 As Long, s2 As Long
  
  data = text_io.plain_in(Config.SAVE_DATA, line_sep)

  '��؂蕶���̍s�ԍ�
  s1 = get_sep_no(data, "delive_elements :")
  s2 = get_sep_no(data, "segments :")
  
  'elements����
  For i = 3 To (s1 - 1)
    Call push_array(elements_d, data(i))
  Next i
  Call elements.load_(elements_d)
  
  'delive_elements����
  For i = (s1 + 1) To (s2 - 1)
    Call push_array(derive_elements_d, data(i))
  Next i
  Call derive_elements.load_(derive_elements_d)

  'segments����
  For i = (s2 + 1) To UBound(data)
    Call push_array(segments_d, data(i))
  Next i
  Call segments.load_(segments_d)
  
End Sub

'�Z�p���[�g������̍s�ԍ�
'// FIXME:��������Refactering to Backend
Private Function get_sep_no(ByRef data() As String, find_val As String) As Long
  Dim i As Long
  
  For i = 0 To UBound(data)
    If data(i) = find_val Then
      get_sep_no = i
      Exit Function
    End If
  Next i
  
  get_sep_no = -1
End Function


'save_data.yaml.�o��
'// FIXME:��������Refactering to Backend
Public Sub save_(ByVal line_sep As Integer)
  Dim data() As String

  Call push_array(data, "data_type : define_elements")
  Call push_array(data, "version : " & Config.getVersion())

  Call push_array(data, "elements :")
  'elements�o��
  Call elements.save_(data)
  
  Call push_array(data, "delive_elements :")
  'delive_elements�o��
  Call derive_elements.save_(data)
  
  Call push_array(data, "segments :")
  'segments�o��
  Call segments.save_(data)
  
  Call text_io.plain_out(Config.SAVE_DATA, data, line_sep)
End Sub

'table_elements.yaml.�o��
Public Sub out_table_elements(ByVal line_sep As Integer)
  Dim data() As String

  'elements�o��
  Call elements.out_table(data)
  'delive_elements�o��
  Call derive_elements.out_table(data)
  
  Call text_io.plain_out(Config.TABLE_ELEMENTS, data, line_sep)
End Sub


'api_elements.yaml.�o��
Public Sub out_api_elements(ByVal line_sep As Integer)
  Dim data() As String

  'elements�o��
  Call elements.out_api(data)
  
  Call text_io.plain_out(Config.API_ELEMENTS, data, line_sep)
End Sub


'types_ddl.�o��
Public Sub out_type_ddl(ByVal line_sep As Integer)
  Dim data() As String
  Call push_array(data, "-- Enum Type DDL")
  Call push_array(data, "")

  'elements�o��
  Call elements.out_type_ddl(data)
  
  Call text_io.plain_out(Config.TYPES_DDL, data, line_sep)
End Sub

