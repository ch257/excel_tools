VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public endOfFile As Boolean
Private fileNum As String
Private line As String

Public Sub OpenFile(file_Path As String, mode As String)
  On Error GoTo err
  fileNum = FreeFile()
  If mode = "w" Then
    Open file_Path For Output As #fileNum
    Exit Sub
  ElseIf mode = "a" Then
    Open file_Path For Append As #fileNum
    Exit Sub
  Else
    Open file_Path For Input As #fileNum
    endOfFile = False
    Line Input #fileNum, line
    If EOF(fileNum) Then
        endOfFile = True
    End If
    Exit Sub
  End If

err:
  My_Err.RiseError ("Cannot open file " & file_Path)
End Sub

Function ReadLine() As String
  ReadLine = line
  
  If EOF(fileNum) Then
    endOfFile = True
    Exit Function
  End If
      
  Line Input #fileNum, line
    
End Function

Sub WriteLine(line As String)
  Print #fileNum, line
End Sub

 Sub CloseFile()
  Close #fileNum
End Sub

Sub CreateFolder(ByVal folderPath As String)
  Dim FSO As Scripting.FileSystemObject
  Dim folders() As String
  Dim fld As Variant
  Dim currTreeFolder As String
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  folders = Split(folderPath, "\")
  currTreeFolder = ""
  For Each fld In folders
    currTreeFolder = currTreeFolder & fld & "\"
    If Not FSO.FolderExists(currTreeFolder) Then
      MkDir currTreeFolder
    End If
  Next fld
End Sub

Sub ClearFolder(ByVal folderPath As String)
  If Dir(folderPath) <> "" Then
    Kill folderPath & "*.*"
  End If
End Sub
