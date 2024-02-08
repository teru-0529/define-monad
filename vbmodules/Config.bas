Attribute VB_Name = "Config"
Option Explicit

'// �J�����[�h�̏ꍇ��True�i�����[�X����False�ŏo�ׂ���j
Public Const IS_DEVELOP_MODE = True

'// �t�H�[�}�b�g�ǉ��s���i�����̐ݒ���s���ǉ��s���j
Public Const ADDED_ROWS = 30

'// ini�t�@�C���Ǎ��ݗp�֐���`
Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal section As String, _
    ByVal key As Any, _
    ByVal default As String, _
    ByVal value As String, _
    ByVal Size As Long, _
    ByVal ini_path As String _
) As Long

'// ���o�̓p�X
Public SAVE_DATA As String
Public TABLE_ELEMENTS As String
Public API_ELEMENTS As String
Public TYPES_DDL As String

Public VIEW_DIR As String

'// �o�[�W�������擾�iFull�j
Public Function getFullVersion() As String
  getFullVersion = Process.outerExec("version -F")
End Function

'// �o�[�W�������擾
Public Function getVersion() As String
  getVersion = Process.outerExec("version")
End Function

'// ./vba.ini��������擾����Public�ϐ��ɐݒ肷��
Public Sub init(ByVal iniFile As String)
  Dim startTime As Double: startTime = Timer

  Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
  Dim iniPath As String: iniPath = FSO.BuildPath(ThisWorkbook.path, iniFile)

  Debug.Print "|----|---- configuration setup start ----|----|"

  SAVE_DATA = getIniValue("Path", "saveData", iniPath)
  Debug.Print "[config] SAVE_DATA: " & SAVE_DATA

  TABLE_ELEMENTS = getIniValue("Path", "tableElements", iniPath)
  Debug.Print "[config] TABLE_ELEMENTS: " & TABLE_ELEMENTS

  API_ELEMENTS = getIniValue("Path", "apiElements", iniPath)
  Debug.Print "[config] API_ELEMENTS: " & API_ELEMENTS

  TYPES_DDL = getIniValue("Path", "typesDDL", iniPath)
  Debug.Print "[config] TYPES_DDL: " & TYPES_DDL

  VIEW_DIR = getIniValue("Path", "viewDir", iniPath)
  Debug.Print "[config] VIEW_DIR: " & VIEW_DIR

  Set FSO = Nothing
  Call Util.showTime(Timer - startTime)
  Debug.Print "|----|---- configuration setup end ----|----|"
End Sub

Private Function getIniValue(ByVal base As String, ByVal key As String, ByVal path As String) As String
  Const TEMP_LENGTH = 255
  Dim temp As String: temp = Space(TEMP_LENGTH)

  Call GetPrivateProfileString(base, key, "N/A", temp, TEMP_LENGTH, path)
  getIniValue = Trim(Left(temp, InStr(temp, vbNullChar) - 1))

  '// ��΃p�X���擾
  getIniValue = CreateObject("Scripting.FileSystemObject").BuildPath(ThisWorkbook.path, getIniValue)
End Function

