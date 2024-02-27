Attribute VB_Name = "Config"
Option Explicit

Const CHAR_SET = "UTF-8"

Public SAVE_DATA As String
Public API_ELEMENTS As String
Public TYPES_DDL As String
Public VIEW_DIR As String

Public FULL_VERSION As String
Public OPERATION_MODE As String
Public Const CLI_FILE = "define-monad.exe"

'// 開発モードの場合にTrue
Public Function isDevelop() As Boolean
  isDevelop = OPERATION_MODE = "develop"
End Function

Public Function absPath(ByVal path) As String
  '// 絶対パスを取得
  absPath = CreateObject("Scripting.FileSystemObject").BuildPath(ThisWorkbook.path, path)
End Function

'// .envから情報を取得してPublic変数に設定する
Public Sub getEnv()
  Const ENV_FILE = ".env"
  
  Dim adoSt As Object
  Dim line As String, v() As String
  Dim dic As New Dictionary
  Dim startTime As Double: startTime = Timer
  
  Set adoSt = CreateObject("ADODB.Stream")

  With adoSt
    .Type = adTypeText
    .Charset = CHAR_SET
    .LineSeparator = adLF
    .Open
    Call .LoadFromFile(absPath(ENV_FILE))
    
    Do While Not (.EOS)
      line = .ReadText(adReadLine)
      
      If Len(line) = 0 Then GoTo CONTINUE
      If Left(line, 1) = "#" Then GoTo CONTINUE
      
      '// データ行
      If InStr(1, line, "=") > 0 Then
        v = Split(line, "=")
        '// Dictionaryに登録
        Call dic.Add(v(0), v(1))
      End If

CONTINUE:
    Loop
    
    .Close
  End With
  Set adoSt = Nothing

  Debug.Print "|----|---- configuration setup start ----|----|"
  SAVE_DATA = absPath(dic.item("saveData"))
  Debug.Print "[config] SAVE_DATA: " & SAVE_DATA

  API_ELEMENTS = absPath(dic.item("apiElements"))
  Debug.Print "[config] API_ELEMENTS: " & API_ELEMENTS

  TYPES_DDL = absPath(dic.item("typesDDL"))
  Debug.Print "[config] TYPES_DDL: " & TYPES_DDL

  VIEW_DIR = absPath(dic.item("viewDir"))
  Debug.Print "[config] VIEW_DIR: " & VIEW_DIR

  OPERATION_MODE = dic.item("mode")
  Debug.Print "[config] OPERATION_MODE: " & OPERATION_MODE
  
  FULL_VERSION = Process.outerExec("version -F")
  Debug.Print "[config] FULL_VERSION: " & FULL_VERSION
  
  Call Util.showTime(Timer - startTime)
  Debug.Print "|----|---- configuration setup end ----|----|"
End Sub
