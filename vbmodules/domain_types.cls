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


'項目名を指定して要素値を返す。
Public Function get_val(ByVal name As String, ByVal col As Integer, ByVal def_val) As Variant
  Dim elm As Range: Set elm = Range("ドメインタイプKEY").Find(name, LookAt:=xlWhole)
  
  '存在しなければデフォルト値を返却
  If elm Is Nothing Then
    get_val = def_val
  
  Else
    get_val = elm.Offset(, col).value
  
  End If
  
End Function

'項目名、桁数を指定してDBのデータ型を返す。
Public Function get_db_model(ByVal name As String, ByVal maxlen As String) As String
  Dim elm As Range: Set elm = Range("ドメインタイプKEY").Find(name, LookAt:=xlWhole)
  
  '存在しなければブランクを返却
  If elm Is Nothing Then
    get_db_model = ""
  
  ElseIf elm.value = "時間" Then
    get_db_model = "varchar(5)"
  
  ElseIf elm.Offset(, 4).value = "varchar" Then
    get_db_model = elm.Offset(, 4).value & "(" & maxlen & ")"
  
  Else
    get_db_model = elm.Offset(, 4).value
    
  End If
  
End Function

