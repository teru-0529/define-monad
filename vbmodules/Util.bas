Attribute VB_Name = "Util"
Option Explicit

Const NULL_VALUE = "null" '// null������

'// �L�������P�[�X���X�l�[�N�P�[�X�̕ϊ�
Public Function camel2snake(ByVal val As String, Optional ByVal is_upper As Boolean = False) As String
  Dim ret As String
  Dim i As Long
  Dim length As Long

  Dim c0 As String, c1 As String

  length = Len(val)
  ret = Left(val, 1)

  For i = 2 To length
    c0 = Mid(val, i - 1, 1)
    c1 = Mid(val, i, 1)

    If UCase(c1) = c1 Then
      '// �啶��
      If i > 1 And UCase(c0) = c0 Then
        '// �O�̕������啶���Ȃ�A���X�R�Ȃ�
        ret = ret & c1

      Else
        '// �O�̕������������Ȃ�A���X�R����
        ret = ret & "_" & c1

      End If
    Else
      '// ������
      ret = ret & c1

    End If
  Next i

  If is_upper Then
    camel2snake = UCase(ret)
  Else
    camel2snake = LCase(ret)
  End If

End Function

'// �^�C�}�[�̕\��
Public Sub showTime(ByVal sec As Double)
  '// �����b���擾
  Dim hh As String:  hh = Format(Int(sec / 3600), "00")
  Dim mm As String:  mm = Format(Int((sec Mod 3600) / 60), "00")
  Dim ss As String:  ss = Format(Int((sec Mod 3600) Mod 60), "00")

  Dim vSec As Double: vSec = Int(sec * 100) / 100
  '// �f�o�b�N�\��
  Debug.Print "operation-time:(" & hh & ":" & mm & ":" & ss & ") �c" & vSec & "(sec)"
End Sub

'// ********

'�z��v�f�̒ǉ�
Public Sub push_array(ByRef my_array() As String, ByVal item As String)
  If is_array_empty(my_array) Then
    '������
    ReDim Preserve my_array(0)
  Else
    '�v�f���g��
    ReDim Preserve my_array(UBound(my_array) + 1)
  End If
  my_array(UBound(my_array)) = item

End Sub


'�z��̋󔻒�
Private Function is_array_empty(ByRef my_array() As String) As Boolean
On Error GoTo ERROR_
  is_array_empty = IIf(UBound(my_array) >= 0, False, True)
  Exit Function
    
ERROR_:
  If Err.Number = 9 Then
    is_array_empty = True
  Else
    '�z��O�G���[
  End If
End Function


'union���b�p�[
Public Function union_range(ByVal rng1 As Range, ByVal rng2 As Range) As Range
  If rng1 Is Nothing And rng2 Is Nothing Then
    Set union_range = Nothing
  ElseIf rng1 Is Nothing Then
    Set union_range = rng2
  ElseIf rng2 Is Nothing Then
    Set union_range = rng1
  Else
    Set union_range = Union(rng1, rng2)
  End If
End Function


'Array �� String(tab��؂�)�ϊ�
Public Function array2tabstr(ParamArray values() As Variant) As String
  array2tabstr = vbNullString
  Dim val As Variant: For Each val In values
    If 0 < Len(array2tabstr) Then
      array2tabstr = array2tabstr & vbTab
    End If
    array2tabstr = array2tabstr & val
  Next val
End Function


Public Sub Print_Array(ByRef arr As Variant)
  Dim i As Long, d As String
    
  i = UBound(arr)
  For i = LBound(arr) To UBound(arr)
    d = Replace(arr(i), vbTab, "<TAB>")
    Debug.Print i & ": " & d & "<E>"
  Next i
  
End Sub

Public Function indent(ByVal num As Integer) As String
  indent = String(num * YAML_INDENT_SPACE, " ")
End Function

' yaml �p��key-value��������쐬����ivalue����̏ꍇ��null�l�ɕϊ��j
Public Function to_yaml(ByVal key As String, ByVal value As String) As String
  to_yaml = key & ": "
  If value = "" Then to_yaml = to_yaml & NULL_VALUE Else: to_yaml = to_yaml & value
End Function


' yaml �p��key-value�����񂩂�value�𒊏o����ivalue��null�l�̏ꍇ�͋󕶎��ɕϊ��j
Public Function from_yaml(ByVal yaml As String) As String
  Dim val As String
  
  'yaml�`���ł͂Ȃ��ꍇ�̓u�����N
  If InStr(yaml, ":") = 0 Then
    from_yaml = ""
    Exit Function
  End If
  
  val = Trim(Mid(yaml, InStr(yaml, ":") + 1))
  If val = NULL_VALUE Then from_yaml = "" Else from_yaml = val
End Function

'�f�t�H���g�l�������񂩂ǂ���
Public Function is_string_default(ByVal val As String) As String
  If val = "NO" Or val = "�敪�l" Or val = "�R�[�h�l" Or val = "������" Or val = "�e�L�X�g" Then
    is_string_default = "true"
  Else
    is_string_default = "false"
  End If
End Function

'NULL���񂪕K�{���ǂ���
Public Function must_not_null(ByVal val As String) As String
  If val = "�敪�l" Or val = "��/�t���O" Then
    must_not_null = "true"
  Else
    must_not_null = "false"
  End If
End Function
