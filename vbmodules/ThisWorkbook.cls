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

Dim startTime As Double

Private Sub Workbook_Open()
  startTime = Timer

  '// 入力ファイルの存在チェック
  If Not Util.existFile(Config.absPath(CLI_FILE)) Then
    MsgBox "バックエンド処理用CLIファイルが存在しません [" & CLI_FILE & "]" & _
    vbCrLf & vbCrLf & ThisWorkbook.name & " と同じディレクトリに " & CLI_FILE & " を格納してください"
    '// 保存せず終了する
    ThisWorkbook.Close SaveChanges:=False
  End If
  
  If Not Util.existFile(Config.absPath(".env")) Then
    MsgBox "vba設定ファイルが存在しません [.env]" & _
    vbCrLf & vbCrLf & ThisWorkbook.name & " と同じディレクトリに .env を格納してください"
    '// 保存せず終了する
    ThisWorkbook.Close SaveChanges:=False
  End If
  
  '// .env読込み
  Call Config.getEnv

  Application.ScreenUpdating = False
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  If Not Util.existFile(SAVE_DATA) Then
    MsgBox "入出力ファイルが存在しないため空シートを生成します" & vbCrLf & "[" & SAVE_DATA & "]"
  End If
  
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

  Debug.Print "[operationMode] " & Config.OPERATION_MODE
  Debug.Print Config.FULL_VERSION
  Call showTime(Timer - startTime)
  MsgBox Config.FULL_VERSION
End Sub

' 保存時
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  startTime = Timer

  '// クライアントバリデーション
  Call Process.validate

End Sub

' 保存時
Private Sub Workbook_AfterSave(ByVal Success As Boolean)

  '// .env読込み
  Call Config.getEnv

  Application.ScreenUpdating = False
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  '// Yamlファイルの出力
  Call outerExec("save -D " & SAVE_DATA & " -E " & ThisWorkbook.FullName)

  '// OpenAPI定義用要素を保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("output -T api-element -D " & SAVE_DATA & " -F " & API_ELEMENTS)

  '// DB用Enum定義を保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("output -T type-ddl -D " & SAVE_DATA & " -F " & TYPES_DDL)

  '// view用のtsvファイルを保存したyamlファイルから生成(golang呼び出し)
  Call outerExec("view -D " & SAVE_DATA & " -O " & VIEW_DIR)
  '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  Application.ScreenUpdating = True
  Application.StatusBar = False

  Debug.Print "[operationMode] " & Config.OPERATION_MODE
  '// 開発モードの場合、モジュールを出力
  If Config.isDevelop() Then Call ExportModule.SaveModule
  Debug.Print Config.FULL_VERSION
  Call showTime(Timer - startTime)
  MsgBox "Git管理用ファイル、テーブル要素/OpenAPI要素を作成しました。"

End Sub
