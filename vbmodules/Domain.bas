Attribute VB_Name = "Domain"
Option Explicit

'// �h���C�����
  Public Const DOM_UUID = "UUID"
  Public Const DOM_SEQUENCE = "NOKEY"
  Public Const DOM_ID = "ID"
  Public Const DOM_ENUM = "�敪�l"
  Public Const DOM_CODE = "�R�[�h�l"
  Public Const DOM_BOOL = "��/�t���O"
  Public Const DOM_DATETIME = "����"
  Public Const DOM_DATE = "���t"
  Public Const DOM_TIME = "����"
  Public Const DOM_INTEGER = "����"
  Public Const DOM_NUMBER = "����"
  Public Const DOM_STRING = "������"
  Public Const DOM_TEXT = "�e�L�X�g"

'// �Ǘ�����
  Public Const ITEM_REG_EX = "���K�\��"
  Public Const ITEM_MIN_DIGITS = "�ŏ�����"
  Public Const ITEM_MAX_DIGITS = "�ő包��"
  Public Const ITEM_VALUE_RANGE = "�ŏ��ő�l"

'// �Ǘ����
  Public Const STATUS_REQUIRED = 1
  Public Const STATUS_ANY = 0
  Public Const STATUS_UNABLED = -1
  Public Const STATUS_NONE = -99


'// �Ǘ����ڂ̏�Ԃ�Ԃ�
Public Function getStatus(ByVal domain As String, ByVal item As String) As Long
  If domain = DOM_UUID Then
    getStatus = STATUS_UNABLED

  ElseIf domain = DOM_SEQUENCE Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_ID Then
    getStatus = STATUS_UNABLED

  ElseIf domain = DOM_ENUM Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_CODE Then
    If item = ITEM_MAX_DIGITS Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_BOOL Then
    getStatus = STATUS_UNABLED

  ElseIf domain = DOM_DATETIME Then
    getStatus = STATUS_UNABLED

  ElseIf domain = DOM_DATE Then
    getStatus = STATUS_UNABLED

  ElseIf domain = DOM_TIME Then
    If item = ITEM_REG_EX Then
      getStatus = STATUS_REQUIRED
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_INTEGER Then
    If item = ITEM_VALUE_RANGE Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_NUMBER Then
    If item = ITEM_VALUE_RANGE Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_STRING Then
    If item = ITEM_MAX_DIGITS Then
      getStatus = STATUS_REQUIRED
    ElseIf item = ITEM_REG_EX Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  ElseIf domain = DOM_TEXT Then
    If item = ITEM_REG_EX Or item = ITEM_MIN_DIGITS Then
      getStatus = STATUS_ANY
    Else
      getStatus = STATUS_UNABLED
    End If

  Else
    getStatus = STATUS_NONE

  End If
End Function
