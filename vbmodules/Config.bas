Attribute VB_Name = "Config"
Option Explicit

'// �J�����[�h�̏ꍇ��True�i�����[�X����False�ŏo�ׂ���j
Public Const IS_DEVELOP_MODE = True

'// �����[�X���ɕύX����
Const VERSION = "3.0.0-rc.0"
Const RELEASE_DATE = "2024/01/16"

'// �o�[�W������񕶎���
Public Function getVersion() As String
  getVersion = "version: " & VERSION & " (releasedAt: " & RELEASE_DATE & ")"
End Function
