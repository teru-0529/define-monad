Attribute VB_Name = "Domain"
Option Explicit

'// ドメイン種類
  Public Const DOM_UUID = "UUID"
  Public Const DOM_SEQUENCE = "NOKEY"
  Public Const DOM_ID = "ID"
  Public Const DOM_ENUM = "区分値"
  Public Const DOM_CODE = "コード値"
  Public Const DOM_BOOL = "可否/フラグ"
  Public Const DOM_DATETIME = "日時"
  Public Const DOM_DATE = "日付"
  Public Const DOM_TIME = "時間"
  Public Const DOM_INTEGER = "整数"
  Public Const DOM_NUMBER = "実数"
  Public Const DOM_STRING = "文字列"
  Public Const DOM_TEXT = "テキスト"

'// 管理項目
  Public Const ITEM_REG_EX = "正規表現"
  Public Const ITEM_MIN_DIGITS = "最小桁数"
  Public Const ITEM_MAX_DIGITS = "最大桁数"
  Public Const ITEM_VALUE_RANGE = "最小最大値"

'// 管理状態
  Public Const STATUS_REQUIRED = 1
  Public Const STATUS_ANY = 0
  Public Const STATUS_UNABLED = -1

'// 管理項目の状態を返す
Public Function getStatus(ByVal Domain As String, ByVal item As String) As Long
  If Domain = DOM_UUID Then
    getStatus = STATUS_UNABLED

  ElseIf Domain = DOM_SEQUENCE Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_ID Then
    getStatus = STATUS_UNABLED

  ElseIf Domain = DOM_ENUM Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_CODE Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_BOOL Then
    getStatus = STATUS_UNABLED

  ElseIf Domain = DOM_DATETIME Then
    getStatus = STATUS_UNABLED

  ElseIf Domain = DOM_DATE Then
    getStatus = STATUS_UNABLED

  ElseIf Domain = DOM_TIME Then
    If item = ITEM_REG_EX Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_INTEGER Then
    If item = ITEM_VALUE_RANGE Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_NUMBER Then
    If item = ITEM_VALUE_RANGE Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_STRING Then
    If item = ITEM_MAX_DIGITS Then
      getStatus = STATUS_REQUIRED
    ElseIf item = ITEM_REG_EX Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf Domain = DOM_TEXT Then
    If item = ITEM_REG_EX Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  Else
    getStatus = STATUS_UNABLED

  End If
End Function
