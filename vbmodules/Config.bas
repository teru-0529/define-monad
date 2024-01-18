Attribute VB_Name = "Config"
Option Explicit

'// 開発モードの場合にTrue（リリース時はFalseで出荷する）
Public Const IS_DEVELOP_MODE = True

'// リリース時に変更する
Const VERSION = "3.0.0-rc.0"
Const RELEASE_DATE = "2024/01/16"

'// バージョン情報文字列
Public Function getVersion() As String
  getVersion = "version: " & VERSION & " (releasedAt: " & RELEASE_DATE & ")"
End Function
