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

Sub ReadSettings(filePath As String, settings As Scripting.Dictionary)
  Dim RW_File As New RWFile
  Dim line As String
  Dim lineParts() As String
  Dim mem_section As String
  Dim emptyDict As New Scripting.Dictionary
  
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

Sub ComposeSettings(settings As Scripting.Dictionary, settingsLnk As Scripting.Dictionary)
  Dim k As Variant
  'Dim emptyDict As New Scripting.Dictionary
  
  For Each k In settings.keys
    If TypeName(settings(k)) = "Dictionary" Then
      Call ComposeSettings(settings(k), settingsLnk)
    Else
      If settingsLnk.Exists(settings(k)) Then
        Set settings(k) = settingsLnk(settings(k))
      End If
    End If
  Next
End Sub

Sub DecomposeSettings(settings As Scripting.Dictionary, line As String)
  Dim k As Variant
  
  For Each k In settings
    If TypeName(settings(k)) = "Dictionary" Then
      line = line & vbTab & k & " = " & "3zz_ds_" & k & vbCrLf
      Call DecomposeSettings(settings(k), line)
    Else
      line = line & vbTab & k & " = " & settings(k) & vbCrLf
    End If
  Next k
End Sub

Sub WriteSettings(filePath As String, settings As Scripting.Dictionary)
  Dim RW_File As New RWFile
  Dim line As String
  
  
  Call RW_File.OpenFile(filePath, "w")
  If My_Err.errOccured Then
    Exit Sub
  End If
  
  line = "[3zz_ds]" & vbCrLf
  line2
  Call DecomposeSettings(settings, line)
  Debug.Print line
  
'  line = "[" & "3zz_ds" & "]"
'  For Each sec In settings
'    If TypeName(sec) = "Dictionary" Then
'
'    Else
'      line = "[" & sec & "]"
'    End If
'    RW_File.WriteLine line
'  Next sec
  
  
  
  Call RW_File.CloseFile
  
End Sub
