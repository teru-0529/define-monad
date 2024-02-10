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

'// 管理状態
Const STATUS_REQUIRED = 1
Const STATUS_ANY = 0
Const STATUS_UNABLED = -1

Dim DOMAIN_STATUS As Variant
Dim DIC_DOMAIN As Object

'// 設定
Private Sub setup()
  Dim domain As Range

  DOMAIN_STATUS = Range("ドメインタイプDATA")
  Set DIC_DOMAIN = CreateObject("Scripting.Dictionary")
  
  For Each domain In Range("ドメインタイプKEY")
    DIC_DOMAIN.Add domain.value, domain.Row - 1
  Next domain
End Sub

'// 入力セルのフォーマット分類
Public Sub devideRange(ByRef base As Range, ByRef devided As DecolationRanges)
  Dim rowNo As Long

  '// セットアップ(連想配列の準備)
  Call setup
  
  '// キーがなければ分配せずに終了
  If Not DIC_DOMAIN.Exists(base.value) Then Exit Sub
  
  rowNo = DIC_DOMAIN(base.value)
  
  '// 正規表現
  If DOMAIN_STATUS(rowNo, 1) = STATUS_REQUIRED Then
    Set devided.stringRangeR = Util.union_range(devided.stringRangeR, base.Offset(, 1))
  ElseIf DOMAIN_STATUS(rowNo, 1) = STATUS_ANY Then
    Set devided.stringRange = Util.union_range(devided.stringRange, base.Offset(, 1))
  ElseIf DOMAIN_STATUS(rowNo, 1) = STATUS_UNABLED Then
    Set devided.unabledRange = Util.union_range(devided.unabledRange, base.Offset(, 1))
  End If

  '// 最小桁数
  If DOMAIN_STATUS(rowNo, 2) = STATUS_REQUIRED Then
    Set devided.lengthRangeR = Util.union_range(devided.lengthRangeR, base.Offset(, 2))
  ElseIf DOMAIN_STATUS(rowNo, 2) = STATUS_ANY Then
    Set devided.lengthRange = Util.union_range(devided.lengthRange, base.Offset(, 2))
  ElseIf DOMAIN_STATUS(rowNo, 2) = STATUS_UNABLED Then
    Set devided.unabledRange = Util.union_range(devided.unabledRange, base.Offset(, 2))
  End If
  
  '// 最大桁数
  If DOMAIN_STATUS(rowNo, 3) = STATUS_REQUIRED Then
    Set devided.lengthRangeR = Util.union_range(devided.lengthRangeR, base.Offset(, 3))
  ElseIf DOMAIN_STATUS(rowNo, 3) = STATUS_ANY Then
    Set devided.lengthRange = Util.union_range(devided.lengthRange, base.Offset(, 3))
  ElseIf DOMAIN_STATUS(rowNo, 3) = STATUS_UNABLED Then
    Set devided.unabledRange = Util.union_range(devided.unabledRange, base.Offset(, 3))
  End If
  
  '// 最小最大値
  If DOMAIN_STATUS(rowNo, 4) = STATUS_REQUIRED Then
    Set devided.numericRangeR = Util.union_range(devided.numericRangeR, base.Offset(, 4))
    Set devided.numericRangeR = Util.union_range(devided.numericRangeR, base.Offset(, 5))
  ElseIf DOMAIN_STATUS(rowNo, 4) = STATUS_ANY Then
    Set devided.numericRange = Util.union_range(devided.numericRange, base.Offset(, 4))
    Set devided.numericRange = Util.union_range(devided.numericRange, base.Offset(, 5))
  ElseIf DOMAIN_STATUS(rowNo, 4) = STATUS_UNABLED Then
    Set devided.unabledRange = Util.union_range(devided.unabledRange, base.Offset(, 4))
    Set devided.unabledRange = Util.union_range(devided.unabledRange, base.Offset(, 5))
  End If
End Sub
