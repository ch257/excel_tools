VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function ParseLine(ByVal line As String) As String()
  Dim retVals(2) As String
  Dim sq1_pos, sq2_pos, eq_pos, sc_pos As Long
  
  retVals(0) = ""
  retVals(1) = ""
  retVals(2) = ""
  
  sc_pos = InStr(line, ";")
  If sc_pos > 0 Then
    line = Mid(line, 1, sc_pos - 1)
  End If
  sq1_pos = InStr(line, "[")
  sq2_pos = InStr(line, "]")
  If sq1_pos > 0 And sq2_pos > 0 Then
    retVals(0) = Replace(Trim(Mid(line, sq1_pos + 1, sq2_pos - sq1_pos - 1)), vbTab, "")
  Else
    eq_pos = InStr(line, "=")
    If eq_pos > 0 Then
      retVals(1) = Replace(Trim(Mid(line, 1, eq_pos - 1)), vbTab, "")
      retVals(2) = Replace(Trim(Mid(line, eq_pos + 1)), vbTab, "")
    End If
  End If
  
  ParseLine = retVals
End Function

Sub ReadSettings(ByVal filePath As String, settings As Scripting.Dictionary)
  Dim RW_File As New RWFile
  Dim line As String
  Dim lineParts() As String
  Dim mem_section As String
  Dim emptyDict As New Scripting.Dictionary
  Dim k As Variant
  
  Call RW_File.OpenFile(filePath, "r")
  If My_Err.errOccured Then
    Exit Sub
  End If
  
  While Not RW_File.endOfFile
    line = RW_File.ReadLine
    If line <> "" Then
      lineParts = ParseLine(line)
      If lineParts(0) <> "" Then
        If settings.Exists(lineParts(0)) Then
          mem_section = lineParts(0)
        Else
          Set emptyDict = Nothing
          settings.Add lineParts(0), emptyDict
          mem_section = lineParts(0)
        End If
      End If
      If lineParts(1) <> "" And lineParts(2) <> "" Then
        settings(mem_section)(lineParts(1)) = lineParts(2)
      End If
      
    End If
  Wend
  
  Call RW_File.CloseFile
End Sub

Private Sub LinkSubSettings(ByRef settings As Scripting.Dictionary, ByRef settingsLnk As Scripting.Dictionary, toDelete As String)
  Dim k As Variant
  
  For Each k In settings.keys
    If TypeName(settings(k)) = "Dictionary" Then
      Call LinkSubSettings(settings(k), settingsLnk, toDelete)
    Else
      If settingsLnk.Exists(settings(k)) Then
        toDelete = toDelete & settings(k) & ","
        Set settings(k) = settingsLnk(settings(k))
        Call LinkSubSettings(settings(k), settingsLnk, toDelete)
      End If
    End If
  Next k
End Sub

Sub ComposeSettings(ByRef settings As Scripting.Dictionary)
  Dim k As Variant
  Dim toDelete As String
  
  toDelete = ""
  Call LinkSubSettings(settings, settings, toDelete)
  
  If Len(toDelete) > 0 Then
    For Each k In Split(Mid(toDelete, 1, Len(toDelete) - 1), ",")
      settings.Remove k
    Next k
  End If
End Sub
  
Private Sub UnlinkSubSettings(ParentK As String, settings As Scripting.Dictionary, settingsLnk As Scripting.Dictionary)
  Dim k As Variant
  
  For Each k In settings
    If TypeName(settings(k)) = "Dictionary" Then
      Call UnlinkSubSettings(CStr(k), settings(k), settingsLnk)
      If ParentK <> "" Then
        settingsLnk.Add ParentK & "_" & k, settings(k)
        settings.Remove k
        settings.Add k, ParentK & "_" & k
      End If
    Else
    End If
  Next k
End Sub

Sub DecomposeSettings(settings As Scripting.Dictionary)
  Dim k As Variant
  Call UnlinkSubSettings("", settings, settings)
End Sub

Sub WriteSettings(settings As Scripting.Dictionary, filePath As String)
  Dim RW_File As New RWFile
  Dim str As String
  Dim k, k1 As Variant
  
  Call DecomposeSettings(settings)
  
  str = ""
  For Each k In settings
    If TypeName(settings(k)) = "Dictionary" Then
      str = str & "[" & k & "]" & vbCrLf
      For Each k1 In settings(k)
        str = str & vbTab & k1 & " = " & settings(k)(k1) & vbCrLf
      Next k1
      str = str & vbCrLf
    End If
  Next k
  
  Call RW_File.OpenFile(filePath, "w")
  If My_Err.errOccured Then
    Exit Sub
  End If
  
  Call RW_File.WriteLine(str)
  Call RW_File.CloseFile
End Sub

Sub PrintSettings(settings As Scripting.Dictionary, offset As String)
  Dim str As String
  Call SettingsToJSON(settings, offset, offset, str)
  Debug.Print "{" & vbCrLf & str & "}" & vbCrLf
End Sub

Private Function EscapeJSON(str As String) As String
  EscapeJSON = Replace(str, "\", "\\")
End Function

Private Sub SettingsToJSON(settings As Scripting.Dictionary, offset As String, ByVal offset_m As String, str As String)
  Dim k As Variant
  For Each k In settings
    If TypeName(settings(k)) = "Dictionary" Then
      str = str & offset & """" & k & """" & ": {" & vbCrLf
      offset = offset & offset_m
      Call SettingsToJSON(settings(k), offset, offset_m, str)
      offset = Mid(offset, 1, Len(offset) - Len(offset_m))
      If str <> "" Then
        str = Mid(str, 1, Len(str) - 3) & vbCrLf
      End If
      str = str & offset & "}," & vbCrLf
    Else
      str = str & offset & """" & k & """" & ": " & """" & EscapeJSON(settings(k)) & """" & "," & vbCrLf
    End If
  Next k
  
  If offset = offset_m Then
    str = Mid(str, 1, Len(str) - 3) & vbCrLf
  End If
End Sub

Function GetValue(ByVal dict As Scripting.Dictionary, ByVal key As String) As String
  If dict.Exists(key) Then
    GetValue = dict(key)
  ElseIf dict.Exists("default") Then
    GetValue = dict("default")
  Else
    GetValue = ""
  End If
End Function
    
