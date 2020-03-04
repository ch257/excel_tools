VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function GetValue(ByVal dict As Scripting.Dictionary, ByVal key As String) As String
  If dict.Exists(key) Then
    GetValue = dict(key)
  ElseIf dict.Exists("default") Then
    GetValue = dict("default")
  Else
    GetValue = ""
  End If
End Function
    
