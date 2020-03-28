Attribute VB_Name = "Main"
Option Explicit
Public My_Err As New Errors
Dim settings As Scripting.Dictionary

Private Sub Init(iniFilesString As String)
  Dim thisWbFolder As String
  Dim cmpCount As Long
  Dim iniFile As Variant
  Dim iniFiles() As String
  Dim RW_Ini As New RWini
  
  Set My_Err = New Errors
  Set settings = New Scripting.Dictionary
  
  thisWbFolder = ThisWorkbook.Path & "\"
  
  iniFiles = Split(iniFilesString, ",")
  
  cmpCount = 0
  Set settings = New Scripting.Dictionary
  
  For Each iniFile In iniFiles
    iniFiles(cmpCount) = thisWbFolder + iniFile
    Call RW_Ini.ReadSettings(iniFiles(cmpCount), settings)
    Call RW_Ini.ComposeSettings(settings)

    If My_Err.errOccured Then
      Exit Sub
    End If
    cmpCount = cmpCount + 1
  Next iniFile
  'Call RW_Ini.PrintSettings(settings, "  ")
End Sub

Private Sub PrepareForDailySnapshots()
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
  Dim zzPackMinMovings() As String
  Dim zzPackMinMoving As Variant
  Dim iniFilesString As String
  
  thisWbFolder = ThisWorkbook.Path & "\"
  
  iniFilesString = ""
  iniFilesString = iniFilesString & "settings\main.ini"
  iniFilesString = iniFilesString & "," & "settings\zz_pack_ds.ini"
  iniFilesString = iniFilesString & "," & "settings\tick_ds.ini"
  
  Call Init(iniFilesString)
  If My_Err.errOccured Then
    Exit Sub
  End If
  
  inputFileFolder = thisWbFolder & settings("input")("file_folder")
  outputFileFolder = thisWbFolder & settings("output")("file_folder")
  getLine = settings("input")("get_line")
  tickFileList = RW_File.GetDataStoreFileList(inputFileFolder, getLine)
  
  zzPackMinMovings = Split(settings("parameters")("zz_pack_min_movings"), ",")
  For Each zzPackMinMoving In zzPackMinMovings
    
    outputFileFolder = outputFileFolder & zzPackMinMoving
    Call RW_File.CreateFolder(outputFileFolder)
    preparedFileList = RW_File.GetFolderFileList(outputFileFolder)
    
    For cnt = 1 To UBound(tickFileList) - LBound(tickFileList)
      fileName = RW_File.GetFileName(tickFileList(cnt))
      If Not CM.InStringArray(preparedFileList, fileName) Then
        Debug.Print fileName
        Call tickCSV.ReadFromFile(tickFileList(cnt), settings("data_sets")("tick_ds"))
        Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
        Call Ex_Meth.TicksToZZ(tickCSV, CInt(zzPackMinMoving), zzBase)
        zzBase_filePath = outputFileFolder & "\" & fileName
        Call zzBase.WriteToFile(zzBase_filePath, settings("data_sets")("zz_pack_ds"))
        If My_Err.errOccured Then
          Exit Sub
        End If
      End If
    Next cnt
    
    outputFileFolder = Mid(outputFileFolder, 1, Len(outputFileFolder) - Len(zzPackMinMoving))
  Next zzPackMinMoving
End Sub

Sub Run()
  Call PrepareForDailySnapshots
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub


