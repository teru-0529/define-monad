VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim start_time As Double

Private Sub Workbook_Open()
  start_time = Timer

  Application.ScreenUpdating = False
  '��������������������������������������������������������������������
  '�Ɩ����W�b�N
  Call work.element_reflesh '���ڃ��X�g�̍ŐV��
  Call work.segment_reflesh '�敪�l���X�g�̍ŐV��
  
  Call read_write_files.load_(adLF)
  '��������������������������������������������������������������������
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Call showTime(Timer - start_time)
  MsgBox Config.getVersion()
End Sub

' �ۑ���
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  start_time = Timer

  '// �N���C�A���g�o���f�[�V����
  Call Process.validate

End Sub

' �ۑ���
Private Sub Workbook_AfterSave(ByVal Success As Boolean)

  Application.ScreenUpdating = False
  '��������������������������������������������������������������������
  '�Ɩ����W�b�N
  Call read_write_files.save_(adLF)
  Call read_write_files.out_table_elements(adLF)
  Call read_write_files.out_api_elements(adLF)
  Call read_write_files.out_type_ddl(adLF)
  
  Call elements.out_view(adLF)
  Call derive_elements.out_view(adLF)
  Call segments.out_view(adLF)
  '��������������������������������������������������������������������
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Call showTime(Timer - start_time)
  MsgBox "Git�Ǘ��p�t�@�C���A�e�[�u���v�f/OpenAPI�v�f���쐬���܂����B"

  '// �J�����[�h�̏ꍇ�A���W���[�����o��
  If IS_DEVELOP_MODE Then Call ExportModule.SaveModule
  Debug.Print Process.outerExec("version -F")
  Debug.Print Config.getVersion

End Sub