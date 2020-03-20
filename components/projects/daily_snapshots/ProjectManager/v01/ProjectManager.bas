Attribute VB_Name = "ProjectManager"
'Add following references:
  'Microsoft Visual Basic for Applications Extensibility 5.3
  'Microsoft Scripting Runtime

Option Explicit
Private includedComponents() As String

Sub ExportButton()
  Call ExportComponents
End Sub

Sub RemoveButton()
  'Call RemoveComponents
End Sub

Sub ImportButton()
  Call ImportComponents
End Sub

Sub Init()
  Dim thisWbFolder As String
  Dim cmpCount As Long
  Dim cmpFilePath As Variant
  Dim includedComponentsString As String
  
  thisWbFolder = ThisWorkbook.Path & "\"
  includedComponentsString = ""
  includedComponentsString = includedComponentsString & "..\..\components\projects\daily_snapshots\ProjectManager\v01\ProjectManager.bas"
  includedComponentsString = includedComponentsString & "," & "..\..\components\projects\daily_snapshots\Main\v01\Main.bas"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\CommonMethods\v01\CommonMethods.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\Errors\v01\Errors.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\RWFile\v01\RWFile.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\RWini\v01\RWini.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\DataSet\v01\DataSet.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\DataSetIterator\v01\DataSetIterator.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\TypeConvertor\v01\TypeConvertor.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\DataSetTools\v01\DataSetTools.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\common\ChartPlotter\v01\ChartPlotter.cls"

  includedComponentsString = includedComponentsString & "," & "..\..\components\exchange\Buffer\v01\Buffer.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\exchange\ZigZagAbsolute\v01\ZigZagAbsolute.cls"
  includedComponentsString = includedComponentsString & "," & "..\..\components\exchange\ExchangeMethods\v01\ExchangeMethods.cls"
  
  includedComponents = Split(includedComponentsString, ",")
  
  cmpCount = 0
  For Each cmpFilePath In includedComponents
    includedComponents(cmpCount) = thisWbFolder + cmpFilePath
    cmpCount = cmpCount + 1
  Next cmpFilePath
  
End Sub

Sub RemoveComponents()
  Dim cmpComponents As VBIDE.VBComponents
  Dim cmpComponent As VBIDE.VBComponent
  Dim exportFolder, exportClsFolder, exportFrmFolder, exportBasFolder As String

  Set cmpComponents = ThisWorkbook.VBProject.VBComponents
  For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
    If cmpComponent.Name <> "ProjectManager" Then
      Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
          cmpComponents.Remove cmpComponent
        Case vbext_ct_MSForm
          cmpComponents.Remove cmpComponent
        Case vbext_ct_StdModule
          cmpComponents.Remove cmpComponent
        Case vbext_ct_Document
          ''' This is a worksheet or workbook object.
          ''' Don't try to delete.
      End Select
    End If
  Next cmpComponent

End Sub

Sub ImportComponents()
  'Dim thisWbFolder As String
  Dim cmpFilePath As Variant
  Dim cmpComponents As VBIDE.VBComponents
  Dim cmpComponent As VBIDE.VBComponent
  Dim objFSO As Scripting.FileSystemObject
  Set objFSO = New Scripting.FileSystemObject
  
  Call Init
  Set cmpComponents = ThisWorkbook.VBProject.VBComponents
  For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
    If cmpComponent.Name <> "ProjectManager" And Mid(cmpComponent.Name, 1, 6) <> "Module" Then
      Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
          cmpComponents.Remove cmpComponent
        Case vbext_ct_MSForm
          cmpComponents.Remove cmpComponent
        Case vbext_ct_StdModule
          cmpComponents.Remove cmpComponent
        Case vbext_ct_Document
          ''' This is a worksheet or workbook object.
          ''' Don't try to delete.
      End Select
    End If
  Next cmpComponent
  
  Dim cnt As Long
  For Each cmpFilePath In includedComponents
    If cmpFilePath <> Empty And InStr(cmpFilePath, "ProjectManager") = 0 Then
      'cmpFilePath = thisWbFolder & cmpFilePath
      If Not objFSO.FileExists(cmpFilePath) Then
        MsgBox "Can't open file """ & cmpFilePath & """"
        Exit Sub
      End If
      cmpComponents.Import cmpFilePath
    End If
  Next cmpFilePath
  
End Sub

Sub ExportComponents()
  Dim exportFolder, exportClsFolder, exportFrmFolder, exportBasFolder As String
  Dim cmpName, cmpFileName, cmpFileFolder As String
  Dim cmpFilePath As Variant
  Dim cmpComponent As VBIDE.VBComponent
  Dim bExport As Boolean
  Dim objFSO As Scripting.FileSystemObject
  Set objFSO = New Scripting.FileSystemObject
  
  Call Init
  For Each cmpFilePath In includedComponents
    cmpName = objFSO.GetBaseName(cmpFilePath)
    cmpFileName = objFSO.GetFileName(cmpFilePath)
    cmpFileFolder = Mid(cmpFilePath, 1, Len(cmpFilePath) - Len(cmpFileName))
    CreateFolder cmpFileFolder
    ClearFolder cmpFileFolder
    
    For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
      If Mid(cmpComponent.Name, 1, 6) <> "Module" Then
        If cmpName = cmpComponent.Name Then
          bExport = True
          Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
              cmpFileName = cmpFileFolder & cmpFileName
            Case vbext_ct_MSForm
              cmpFileName = cmpFileFolder & cmpFileName
            Case vbext_ct_StdModule
              cmpFileName = cmpFileFolder & cmpFileName
            Case vbext_ct_Document
              ''' This is a worksheet or workbook object.
              ''' Don't try to export.
              bExport = False
          End Select
          
          If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export cmpFileName
          End If
        End If
        
      End If
    Next cmpComponent
  Next cmpFilePath
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



