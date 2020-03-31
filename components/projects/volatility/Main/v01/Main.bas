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
  iniFilesString = iniFilesString & "," & "settings\hist_ds.ini"
  iniFilesString = iniFilesString & "," & "settings\volatility_ds.ini"
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

Sub Run()
   'VBA Settings – Speed Up Code:
  'If your goal is to speed up your code,
  'you should also consider adjusting these other settings:
  'Disabling Screenupdating can make a huge difference in speed:
  Application.ScreenUpdating = False

  'Turning off the Status Bar will also make a small difference:
  Application.DisplayStatusBar = False

  'If your workbook contains events you should
  'also disable events at the start of your procedures
  '(to speed up code and to prevent endless loops!):
  Application.EnableEvents = False

  'Last, your VBA code can be slowed down when Excel
  'tries to re-calculate page breaks (Note: not all procedures will be impacted).
  'To turn off DisplayPageBreaks use this line of code:
  ActiveSheet.DisplayPageBreaks = False
  
  Application.DisplayAlerts = False
    
  Dim thisWbFolder As String
  Dim zzTick As New DataSet
  Dim zzTick_fileFolder As String
  Dim zzTick_filePath As String
  Dim zzTick_fileList() As String
  Dim SelectedZZTick As New DataSet
  Dim zzBase As New DataSet
  Dim Ex_Meth As New ExchangeMethods
  Dim zzBase_MinMoving As Integer
  Dim DS_Tools As New DataSetTools
  Dim outputFileFolder As String
  Dim outputFileName As String
  Dim RW_File As New RWFile
  Dim cnt As Integer
  Dim CM As New CommonMethods
  'Dim SM As New StatMethods
  Dim volatilityDS As New DataSet
  
  thisWbFolder = ThisWorkbook.Path & "\"
    
  Call Init
  If Not My_Err.errOccured Then
    zzTick_fileFolder = thisWbFolder & settings("input")("file_folder")
    zzBase_MinMoving = settings("parameters")("zz_base_min_moving")
    
    outputFileFolder = ThisWorkbook.Path & "\" & settings("output")("file_folder")
    outputFileName = ThisWorkbook.Path & "\" & settings("output")("file_name")
    Call RW_File.CreateFolder(outputFileFolder)
    Call RW_File.ClearFolder(outputFileFolder)
    
    Call volatilityDS.Create(settings("data_sets")("volatility_ds"))
    
    zzTick_fileList = RW_File.GetFolderFileList(zzTick_fileFolder)
    For cnt = 1 To UBound(zzTick_fileList) - LBound(zzTick_fileList)
      
      zzTick_filePath = zzTick_fileFolder & zzTick_fileList(cnt)
      Call zzTick.ReadFromFile(zzTick_filePath, settings("data_sets")("zz_pack_ds"))
      If My_Err.errOccured Then
        Exit For
      End If
      Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
      Call Ex_Meth.ZZToZZ(zzTick, CInt(zzBase_MinMoving), zzBase)
      Call DS_Tools.SelectBetween(zzBase, settings("data_sets")("zz_pack_ds"), "<TIME>", 100000, 110000, SelectedZZTick, settings("data_sets")("zz_pack_ds"))
            
      Call volatilityDS.SetCell("<FILENAME>", cnt - 1, zzTick_fileList(cnt))
      Call volatilityDS.SetCell("<SELECTED_LENGTH>", cnt - 1, SelectedZZTick.rowsCount)
      Call volatilityDS.SetCell("<TOTAL_LENGTH>", cnt - 1, zzBase.rowsCount)
      
    Next cnt
    
    volatilityDS.rowsCount = cnt - 1
    Call volatilityDS.WriteToFile(outputFileFolder & "volatility.txt", settings("data_sets")("zz_pack_ds"))
  End If
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub

