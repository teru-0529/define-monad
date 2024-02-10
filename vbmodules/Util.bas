Attribute VB_Name = "Util"
Option Explicit

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
  Dim hh As String:  hh = format(Int(sec / 3600), "00")
  Dim mm As String:  mm = format(Int((sec Mod 3600) / 60), "00")
  Dim ss As String:  ss = format(Int((sec Mod 3600) Mod 60), "00")

  Dim vSec As Double: vSec = Int(sec * 100) / 100
  '// �f�o�b�N�\��
  Debug.Print "operation-time:(" & hh & ":" & mm & ":" & ss & ") �c" & vSec & "(sec)"
End Sub

'// �z��v�f�̒ǉ�
Public Sub push_array(ByRef my_array() As String, ByVal item As String)
  If is_array_empty(my_array) Then
    '// ������
    ReDim Preserve my_array(0)
  Else
    '// �v�f���g��
    ReDim Preserve my_array(UBound(my_array) + 1)
  End If
  my_array(UBound(my_array)) = item

End Sub

'// �z��̋󔻒�
Private Function is_array_empty(ByRef my_array() As String) As Boolean
On Error GoTo ERROR_
  is_array_empty = IIf(UBound(my_array) >= 0, False, True)
  Exit Function

ERROR_:
  If Err.Number = 9 Then
    is_array_empty = True
  Else
    '// �z��O�G���[
  End If
End Function

'// union���b�p�[
Public Function unionRange(ByVal rng1 As Range, ByVal rng2 As Range) As Range
  If rng1 Is Nothing And rng2 Is Nothing Then
    Set unionRange = Nothing
  ElseIf rng1 Is Nothing Then
    Set unionRange = rng2
  ElseIf rng2 Is Nothing Then
    Set unionRange = rng1
  Else
    Set unionRange = Union(rng1, rng2)
  End If
End Function
