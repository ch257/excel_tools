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

Function GetFolderFileList(folderPath As String) As String()
  Dim oFSO As Object
  Dim oFolder As Object
  Dim oFile As Object
  Dim fileName As String
  Dim fileNumber, fileCnt As Integer
  Dim fileList() As String
   
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(folderPath)
  
  fileNumber = oFolder.Files.Count
  If fileNumber > 0 Then
    ReDim fileList(fileNumber)
    fileCnt = 1
    For Each oFile In oFolder.Files
      fileName = oFile.Name
      fileList(fileCnt) = fileName
      fileCnt = fileCnt + 1
    Next oFile
  Else
    ReDim fileList(0)
  End If
  
  GetFolderFileList = fileList
End Function

Private Sub CompleteDataStoreFileList(dataStorePath As String, layer As Integer, takeFrom() As String, fileList As String)
  Dim oFSO As Object
  Dim oFolder As Object
  Dim oFile As Object
  Dim fsObj As Object
  Dim fileNumber, subFolderNumber As Long
  Dim subFolderName, fileName As Variant
  Dim allowFolderList As String
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(dataStorePath)
  
  fileNumber = oFolder.Files.Count
  subFolderNumber = oFolder.SubFolders.Count
  
  If fileNumber > 0 Then
    For Each fsObj In oFolder.Files
      fileList = fileList & ";" & dataStorePath & fsObj.Name
    Next fsObj
  End If
  If subFolderNumber > 0 Then
    allowFolderList = takeFrom(layer)
    For Each fsObj In oFolder.SubFolders
      If allowFolderList = "" Or InStr(allowFolderList, fsObj.Name) > 0 Then
        Call CompleteDataStoreFileList(dataStorePath & fsObj.Name & "\", layer + 1, takeFrom, fileList)
      End If
    Next fsObj
  End If
End Sub

Function GetDataStoreFileList(dataStorePath As String, getLine As String) As String()
  Dim takeFrom() As String
  Dim fileList As String
  
  takeFrom = Split(getLine, ":")
  fileList = ""
  Call CompleteDataStoreFileList(dataStorePath, 0, takeFrom, fileList)
  
  GetDataStoreFileList = Split(fileList, ";")
End Function

Function GetFileName(ByVal filePath As String)
  Dim sepPos As Long
  
  sepPos = InStr(filePath, "\")
  While sepPos > 0
    filePath = Mid(filePath, sepPos + 1)
    sepPos = InStr(filePath, "\")
  Wend
  GetFileName = filePath
End Function



