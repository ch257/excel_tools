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
  Dim RW_Ini As New RWini
  
  Set My_Err = New Errors
  
  thisWbFolder = ThisWorkbook.Path & "\"
  iniFilesString = ""
  iniFilesString = iniFilesString & "settings\main.ini"
  iniFilesString = iniFilesString & "," & "settings\zz_pack_ds.ini"
  iniFilesString = iniFilesString & "," & "settings\tick_ds.ini"
  iniFilesString = iniFilesString & "," & "settings\base_first_join_rules.ini"
  iniFilesString = iniFilesString & "," & "settings\base_first_second_join_rules.ini"
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
  Call RW_Ini.PrintSettings(settings, "  ")
End Sub

Sub Run()
  Dim thisWbFolder As String
  Dim zzTick As New DataSet
  Dim zzTick_filePath As String
  Dim zzBase As New DataSet
  Dim zzFirst As New DataSet
  Dim zzSecond As New DataSet
'  Dim BaseFirstJoned As New DataSet
'  Dim BaseFirstJoned_Settings As New Scripting.Dictionary
'  Dim BaseFirstSecondJoned As New DataSet
'  Dim BaseFirstSecondJoned_Settings As New Scripting.Dictionary
  Dim Ex_Meth As New ExchangeMethods
  Dim zzBase_MinMoving, zzFirst_MinMoving, zzSecond_MinMoving As Integer
'  Dim DS_Tools As New DataSetTools
  
  'zzBase_filePath, zzFirst_filePath, zzSecond_filePath, joinedDS_filePath As String
  
  
  thisWbFolder = ThisWorkbook.Path & "\"
    
  Call Init
  If Not My_Err.errOccured Then
    zzTick_filePath = thisWbFolder & settings("input")("file_folder") & settings("input")("zz_tick_file_name")
    zzBase_MinMoving = settings("parameters")("zz_base_min_moving")
    zzFirst_MinMoving = settings("parameters")("zz_frst_min_moving")
    zzSecond_MinMoving = settings("parameters")("zz_second_min_moving")
  
    Call zzTick.ReadFromFile(zzTick_filePath, settings("data_sets")("zz_pack_ds"))
  End If
  If Not My_Err.errOccured Then
    Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
    Call zzFirst.Create(settings("data_sets")("zz_pack_ds"))
    Call zzSecond.Create(settings("data_sets")("zz_pack_ds"))
    Call Ex_Meth.ZZToZZ(zzTick, CInt(zzBase_MinMoving), zzBase)
    Call Ex_Meth.ZZToZZ(zzBase, CInt(zzFirst_MinMoving), zzFirst)
    Call Ex_Meth.ZZToZZ(zzBase, CInt(zzSecond_MinMoving), zzSecond)
'
'    Call DS_Tools.FullJoin( _
'      zzBase, settings("data_sets")("zz_pack_ds"), _
'      zzFirst, settings("data_sets")("zz_pack_ds"), _
'      BaseFirstJoned, BaseFirstJoned_Settings, _
'      settings("data_sets")("base_first_join_rules"))
'
'    Call DS_Tools.FullJoin( _
'      BaseFirstJoned, BaseFirstJoned_Settings, _
'      zzSecond, settings("data_sets")("zz_pack_ds"), _
'      BaseFirstSecondJoned, BaseFirstSecondJoned_Settings, _
'      settings("data_sets")("base_first_second_join_rules"))
'
'    zzBase_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_base_file_name")
'    zzFirst_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_first_file_name")
'    zzSecond_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_second_file_name")
'    joinedDS_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("joined_ds_file_name")
'    Call zzBase.WriteToFile(zzBase_filePath, settings("data_sets")("zz_pack_ds"))
'    Call zzFirst.WriteToFile(zzFirst_filePath, settings("data_sets")("zz_pack_ds"))
'    Call zzSecond.WriteToFile(zzSecond_filePath, settings("data_sets")("zz_pack_ds"))
'    Call BaseFirstSecondJoned.WriteToFile(joinedDS_filePath, BaseFirstSecondJoned_Settings)
'
''    Dim RW_Ini As New RWini
''    Dim iniFilePath As String
''    thisWbFolder = ThisWorkbook.Path & "\"
''    iniFilePath = thisWbFolder & "settings\daily_snapshots\3zz_ds.ini"
''    Call RW_Ini.WriteSettings(iniFilePath, BaseFirstSecondJoned_Settings)
  End If
  
  
  If Not My_Err.errOccured Then
'    Dim zz_filePath, zz_ds_settings_file, zz_ds_plot_settings_file As String
'    Dim zz_ds_settings As New Scripting.Dictionary
'    Dim RW_Ini As New RWini
'    Dim zzCSV As New DataSet
'    Dim ChPl As New ChartPlotter
'
'    zz_filePath = thisWbFolder & "data\daily_snapshots\output\3zz.txt"
'    zz_ds_settings_file = thisWbFolder & "settings\daily_snapshots\3zz_ds.ini"
'    zz_ds_plot_settings_file = thisWbFolder & "settings\daily_snapshots\3zz_ds_plot.ini"
'    Call RW_Ini.ReadSettings(zz_ds_settings_file, zz_ds_settings)
'    Call RW_Ini.ReadSettings(zz_ds_plot_settings_file, zz_ds_settings)
'    Call RW_Ini.ComposeSettings(zz_ds_settings, zz_ds_settings)
'    Call zzCSV.ReadFromFile(zz_filePath, zz_ds_settings("3zz_ds"))
'
'    Call ChPl.PlotChart(zzCSV, zz_ds_settings("3zz_ds"), "PM")
'    Dim exportFileFolder As String
'    Dim RW_File As New RWFile
'
'    exportFileFolder = ThisWorkbook.Path & "\" & settings("output")("file_folder") & settings("output")("img_subfolder")
'    Call RW_File.CreateFolder(exportFileFolder)
'    Call RW_File.ClearFolder(exportFileFolder)
'    Call ChPl.ExportCharts("001", exportFileFolder)
  End If
  
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub


Sub test(dataStorePath As String, layer As Integer, takeFrom() As String)
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
      Debug.Print layer, dataStorePath & fsObj.Name
    Next fsObj
  End If
  If subFolderNumber > 0 Then
    allowFolderList = takeFrom(layer)
    For Each fsObj In oFolder.SubFolders
      If allowFolderList = "" Or InStr(allowFolderList, fsObj.Name) > 0 Then
        Call test(dataStorePath & fsObj.Name & "\", layer + 1, takeFrom)
      End If
    Next fsObj
  End If
End Sub
