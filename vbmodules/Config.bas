Attribute VB_Name = "Config"
Option Explicit

'// �J�����[�h�̏ꍇ��True�i�����[�X����False�ŏo�ׂ���j
Public Const IS_DEVELOP_MODE = True

'// �o�[�W�������擾�iFull�j
Public Function getFullVersion() As String
  getFullVersion = Process.outerExec("version -F")
End Function

'// �o�[�W�������擾
Public Function getVersion() As String
  getVersion = Process.outerExec("version")
End Function
