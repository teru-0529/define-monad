VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function save_data_path() As String
  save_data_path = Range("save_data_path")
End Function

Public Function table_elements_path() As String
  table_elements_path = Range("table_elements_path")
End Function

Public Function api_elements_path() As String
  api_elements_path = Range("api_elements_path")
End Function

Public Function types_ddl_path() As String
  types_ddl_path = Range("types_ddl_path")
End Function


Public Function view_elements_path() As String
  view_elements_path = Range("view_elements_path")
End Function

Public Function view_derive_elements_path() As String
  view_derive_elements_path = Range("view_derive_elements_path")
End Function

Public Function view_segments_path() As String
  view_segments_path = Range("view_segments_path")
End Function



