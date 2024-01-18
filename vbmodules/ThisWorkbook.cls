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
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  '業務ロジック
  Call work.element_reflesh '項目リストの最新化
  Call work.segment_reflesh '区分値リストの最新化
  
  Call read_write_files.load_(adLF)
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Call showTime(Timer - start_time)
  MsgBox Config.getFullVersion()
End Sub

' 保存時
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  start_time = Timer

  '// クライアントバリデーション
  Call Process.validate

End Sub

' 保存時
Private Sub Workbook_AfterSave(ByVal Success As Boolean)

  Application.ScreenUpdating = False
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  '業務ロジック
  Call read_write_files.save_(adLF)
  Call read_write_files.out_table_elements(adLF)
  Call read_write_files.out_api_elements(adLF)
  Call read_write_files.out_type_ddl(adLF)
  
  Call elements.out_view(adLF)
  Call derive_elements.out_view(adLF)
  Call segments.out_view(adLF)
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Call showTime(Timer - start_time)
  MsgBox "Git管理用ファイル、テーブル要素/OpenAPI要素を作成しました。"

  '// 開発モードの場合、モジュールを出力
  If IS_DEVELOP_MODE Then Call ExportModule.SaveModule
  Debug.Print Config.getVersion
  Debug.Print Config.getFullVersion

End Sub
