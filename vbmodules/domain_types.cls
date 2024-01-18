VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "domain_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'���ږ����w�肵�ėv�f�l��Ԃ��B
Public Function get_val(ByVal name As String, ByVal col As Integer, ByVal def_val) As Variant
  Dim elm As Range: Set elm = Range("�h���C���^�C�vKEY").Find(name, LookAt:=xlWhole)
  
  '���݂��Ȃ���΃f�t�H���g�l��ԋp
  If elm Is Nothing Then
    get_val = def_val
  
  Else
    get_val = elm.Offset(, col).value
  
  End If
  
End Function

'���ږ��A�������w�肵��DB�̃f�[�^�^��Ԃ��B
Public Function get_db_model(ByVal name As String, ByVal maxlen As String) As String
  Dim elm As Range: Set elm = Range("�h���C���^�C�vKEY").Find(name, LookAt:=xlWhole)
  
  '���݂��Ȃ���΃u�����N��ԋp
  If elm Is Nothing Then
    get_db_model = ""
  
  ElseIf elm.value = "����" Then
    get_db_model = "varchar(5)"
  
  ElseIf elm.Offset(, 4).value = "varchar" Then
    get_db_model = elm.Offset(, 4).value & "(" & maxlen & ")"
  
  Else
    get_db_model = elm.Offset(, 4).value
    
  End If
  
End Function
