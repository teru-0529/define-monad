VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "work"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'���ڈꗗ�iD��j�̏����ŐV������
Public Sub element_reflesh()
  Dim dist As Range, src As Range
  Set dist = Range("D2"): dist.value = "NO_DATA"
  
  Range("���ڃ��X�g").ClearContents
  
  For Each src In elements.UsedRange.Offset(1).Resize(, 1) '���ڃV�[�g��1��ڂ��Ώ�
    If src.value <> "" Then
      dist.value = src.value
      Set dist = dist.Offset(1)
    End If
  Next src

End Sub


'�敪�l�ꗗ�iF��j�̏����ŐV������
Public Sub segment_reflesh()
  Dim dist As Range, src As Range
  Set dist = Range("F2"): dist.value = "NO_DATA"
  
  Range("�敪�l���X�g").ClearContents
  
  For Each src In elements.UsedRange.Offset(1).Resize(, 1) '���ڃV�[�g��1��ڂ��Ώ�
    If src.value <> "" And src.Offset(, 3).value = "�敪�l" Then
      dist.value = src.value
      Set dist = dist.Offset(1)
    End If
  Next src
  
End Sub
