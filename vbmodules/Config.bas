Attribute VB_Name = "Config"
Option Explicit

'// 開発モードの場合にTrue（リリース時はFalseで出荷する）
Public Const IS_DEVELOP_MODE = True

'// バージョン情報取得（Full）
Public Function getFullVersion() As String
  getFullVersion = Process.outerExec("version -F")
End Function

'// バージョン情報取得
Public Function getVersion() As String
  getVersion = Process.outerExec("version")
End Function
