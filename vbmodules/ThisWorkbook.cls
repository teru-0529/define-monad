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

  '// パス情報の取得
  Call Config.init("vba.ini")

  Application.ScreenUpdating = False
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  '// 読込み
  Call elements.load(outerExec("load -S elements -D " & SAVE_DATA))
  Call derive_elements.load(outerExec("load -S derive-elements -D " & SAVE_DATA))
  Call segments.load(outerExec("load -S segments -D " & SAVE_DATA))
  
  '// 項目リスト/区分値リスト最新化
  Call work.element_reflesh
  Call work.segment_reflesh
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

  '// パス情報の取得
  Call Config.init("vba.ini")

  Application.ScreenUpdating = False
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  '// Yamlファイルの出力
  Call outerExec("save -D " & SAVE_DATA & " -E " & ThisWorkbook.FullName)

  '// テーブル定義用要素を保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("output -T db-element -D " & SAVE_DATA & " -F " & TABLE_ELEMENTS)

  '// OpenAPI定義用要素を保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("output -T api-element -D " & SAVE_DATA & " -F " & API_ELEMENTS)

  '// DB用Enum定義を保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("output -T type-ddl -D " & SAVE_DATA & " -F " & TYPES_DDL)

  '// view用のtsvファイルを保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("view -D " & SAVE_DATA & " -O " & VIEW_DIR)
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Call showTime(Timer - start_time)
  MsgBox "Git管理用ファイル、テーブル要素/OpenAPI要素を作成しました。"

  '// 開発モードの場合、モジュールを出力
  If IS_DEVELOP_MODE Then Call ExportModule.SaveModule
  Debug.Print Config.getFullVersion
End Sub
