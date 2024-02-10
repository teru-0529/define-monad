Attribute VB_Name = "Util"
Option Explicit

'// キャメルケース→スネークケースの変換
Public Function camel2snake(ByVal val As String, Optional ByVal is_upper As Boolean = False) As String
  Dim ret As String
  Dim i As Long
  Dim length As Long

  Dim c0 As String, c1 As String

  length = Len(val)
  ret = Left(val, 1)

  For i = 2 To length
    c0 = Mid(val, i - 1, 1)
    c1 = Mid(val, i, 1)

    If UCase(c1) = c1 Then
      '// 大文字
      If i > 1 And UCase(c0) = c0 Then
        '// 前の文字が大文字ならアンスコなし
        ret = ret & c1

      Else
        '// 前の文字が小文字ならアンスコあり
        ret = ret & "_" & c1

      End If
    Else
      '// 小文字
      ret = ret & c1

    End If
  Next i

  If is_upper Then
    camel2snake = UCase(ret)
  Else
    camel2snake = LCase(ret)
  End If

End Function

'// タイマーの表示
Public Sub showTime(ByVal sec As Double)
  '// 時分秒を取得
  Dim hh As String:  hh = format(Int(sec / 3600), "00")
  Dim mm As String:  mm = format(Int((sec Mod 3600) / 60), "00")
  Dim ss As String:  ss = format(Int((sec Mod 3600) Mod 60), "00")

  Dim vSec As Double: vSec = Int(sec * 100) / 100
  '// デバック表示
  Debug.Print "operation-time:(" & hh & ":" & mm & ":" & ss & ") …" & vSec & "(sec)"
End Sub

'// 配列要素の追加
Public Sub push_array(ByRef my_array() As String, ByVal item As String)
  If is_array_empty(my_array) Then
    '// 初期化
    ReDim Preserve my_array(0)
  Else
    '// 要素数拡張
    ReDim Preserve my_array(UBound(my_array) + 1)
  End If
  my_array(UBound(my_array)) = item

End Sub

'// 配列の空判定
Private Function is_array_empty(ByRef my_array() As String) As Boolean
On Error GoTo ERROR_
  is_array_empty = IIf(UBound(my_array) >= 0, False, True)
  Exit Function

ERROR_:
  If Err.Number = 9 Then
    is_array_empty = True
  Else
    '// 想定外エラー
  End If
End Function

'// unionラッパー
Public Function unionRange(ByVal rng1 As Range, ByVal rng2 As Range) As Range
  If rng1 Is Nothing And rng2 Is Nothing Then
    Set unionRange = Nothing
  ElseIf rng1 Is Nothing Then
    Set unionRange = rng2
  ElseIf rng2 Is Nothing Then
    Set unionRange = rng1
  Else
    Set unionRange = Union(rng1, rng2)
  End If
End Function
