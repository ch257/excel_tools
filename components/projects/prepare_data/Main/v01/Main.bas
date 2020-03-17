Attribute VB_Name = "Main"
Option Explicit
Public My_Err As New Errors
Dim settings As Scripting.Dictionary

Sub Init()
  Dim thisWbFolder As String
  Dim cmpCount As Long
  Dim iniFile As Variant
  Dim iniFiles() As String
  Dim iniFilesString As String
  Dim RW_ini As New RWini
  
  Set My_Err = New Errors
  
  thisWbFolder = ThisWorkbook.Path & "\"
  iniFilesString = ""
  iniFilesString = iniFilesString & "settings\main.ini"
  iniFilesString = iniFilesString & "," & "settings\zz_pack_ds.ini"
  iniFilesString = iniFilesString & "," & "settings\tick_ds.ini"
  iniFiles = Split(iniFilesString, ",")
  
  cmpCount = 0
  Set settings = New Scripting.Dictionary
  
  For Each iniFile In iniFiles
    iniFiles(cmpCount) = thisWbFolder + iniFile
    Call RW_ini.ReadSettings(iniFiles(cmpCount), settings)
    Call RW_ini.ComposeSettings(settings)

    If My_Err.errOccured Then
      Exit Sub
    End If
    cmpCount = cmpCount + 1
  Next iniFile
End Sub

Sub Run()
  Dim thisWbFolder As String
  Dim tickCSV As New DataSet
  Dim zzBase As New DataSet
  Dim Ex_Meth As New ExchangeMethods
  Dim CM As New CommonMethods
  Dim inputFileFolder As String
  Dim outputFileFolder As String
  Dim zzBase_filePath As String
  Dim RW_File As New RWFile
  Dim tickFileList() As String
  Dim preparedFileList() As String
  Dim getLine As String
  Dim cnt As Long
  Dim fileName As String
  
  thisWbFolder = ThisWorkbook.Path & "\"
  
  Call Init
  If Not My_Err.errOccured Then
    'Call RW_Ini.PrintSettings(settings, "  ")
    inputFileFolder = thisWbFolder & settings("input")("file_folder")
    outputFileFolder = thisWbFolder & settings("output")("file_folder")
    Call RW_File.CreateFolder(outputFileFolder)
    getLine = settings("input")("get_line")
    tickFileList = RW_File.GetDataStoreFileList(inputFileFolder, getLine)
    preparedFileList = RW_File.GetFolderFileList(outputFileFolder)
    
    For cnt = 1 To UBound(tickFileList) - LBound(tickFileList)
      fileName = RW_File.GetFileName(tickFileList(cnt))
      If Not CM.InStringArray(preparedFileList, fileName) Then
        Debug.Print fileName
        Call tickCSV.ReadFromFile(tickFileList(cnt), settings("data_sets")("tick_ds"))
        Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
        Call Ex_Meth.TicksToZZ(tickCSV, 10, zzBase)
        zzBase_filePath = outputFileFolder & fileName
        Call zzBase.WriteToFile(zzBase_filePath, settings("data_sets")("zz_pack_ds"))
      End If
    Next cnt
  End If
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub


