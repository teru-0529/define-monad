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

'項目一覧（D列）の情報を最新化する
Public Sub element_reflesh()
  Dim dist As Range, src As Range
  Set dist = Range("D2"): dist.value = "NO_DATA"
  
  Range("項目リスト").ClearContents
  
  For Each src In elements.UsedRange.Offset(1).Resize(, 1) '項目シートの1列目が対照
    If src.value <> "" Then
      dist.value = src.value
      Set dist = dist.Offset(1)
    End If
  Next src

End Sub


'区分値一覧（F列）の情報を最新化する
Public Sub segment_reflesh()
  Dim dist As Range, src As Range
  Set dist = Range("F2"): dist.value = "NO_DATA"
  
  Range("区分値リスト").ClearContents
  
  For Each src In elements.UsedRange.Offset(1).Resize(, 1) '項目シートの1列目が対照
    If src.value <> "" And src.Offset(, 3).value = "区分値" Then
      dist.value = src.value
      Set dist = dist.Offset(1)
    End If
  Next src
  
End Sub

