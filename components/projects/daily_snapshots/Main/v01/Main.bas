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
  iniFilesString = iniFilesString & "," & "settings\3zz_ds_plot.ini"
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

Private Function CalcZZBaseMinMoving(a As Double, b As Double, c As Double, zzVolatilityMovings As Double) As Integer
  Dim d As Double
  d = b * b - 4 * a * (c - zzVolatilityMovings)
  CalcZZBaseMinMoving = CInt((-b + Sqr(d)) / (2 * a))
End Function

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
  Dim zzVolatility As New DataSet
  Dim SelectedZZ As New DataSet
  Dim zzBase As New DataSet
  Dim zzFirst As New DataSet
  Dim zzSecond As New DataSet
  Dim BaseFirstJoned As New DataSet
  Dim BaseFirstJoned_Settings As New Scripting.Dictionary
  Dim BaseFirstSecondJoned As New DataSet
  Dim BaseFirstSecondJoned_Settings As New Scripting.Dictionary
  Dim Ex_Meth As New ExchangeMethods
  Dim zzBase_MinMoving, zzFirst_MinMoving, zzSecond_MinMoving As Integer
  Dim DS_Tools As New DataSetTools
  Dim ChPl As New ChartPlotter
  Dim exportFileFolder As String
  Dim RW_File As New RWFile
  Dim cnt As Integer
  Dim RW_Ini As New RWini
  Dim CM As New CommonMethods
  Dim memPlotSettings As New Scripting.Dictionary
  Dim a As Double
  Dim b As Double
  Dim c As Double
  
  thisWbFolder = ThisWorkbook.Path & "\"
    
  Call Init
  If Not My_Err.errOccured Then
    zzTick_fileFolder = thisWbFolder & settings("input")("file_folder")
    a = CDbl(settings("parameters")("a"))
    b = CDbl(settings("parameters")("b"))
    c = CDbl(settings("parameters")("c"))
    
    exportFileFolder = ThisWorkbook.Path & "\" & settings("output")("file_folder") & settings("output")("img_subfolder")
    Call RW_File.CreateFolder(exportFileFolder)
    Call RW_File.ClearFolder(exportFileFolder)
    
    Call CM.CopyDict(settings("plot_settings")("3zz_ds_plot"), memPlotSettings)
    
    zzTick_fileList = RW_File.GetFolderFileList(zzTick_fileFolder)
    For cnt = 1 To UBound(zzTick_fileList) - LBound(zzTick_fileList)
      zzTick_filePath = zzTick_fileFolder & zzTick_fileList(cnt)
      Call zzTick.ReadFromFile(zzTick_filePath, settings("data_sets")("zz_pack_ds"))
      If My_Err.errOccured Then
        Exit For
      End If
      
      Call DS_Tools.SelectBetween(zzTick, settings("data_sets")("zz_pack_ds"), "<TIME>", 100500, 110000, SelectedZZ, settings("data_sets")("zz_pack_ds"))
      Call zzVolatility.Create(settings("data_sets")("zz_pack_ds"))
      Call Ex_Meth.ZZToZZ(SelectedZZ, 10, zzVolatility)
      
      zzBase_MinMoving = CalcZZBaseMinMoving(a, b, c, zzVolatility.rowsCount)
      zzFirst_MinMoving = zzBase_MinMoving * 5
      zzSecond_MinMoving = zzBase_MinMoving * 15
      
      Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
      Call zzFirst.Create(settings("data_sets")("zz_pack_ds"))
      Call zzSecond.Create(settings("data_sets")("zz_pack_ds"))
      Call Ex_Meth.ZZToZZ(zzTick, CInt(zzBase_MinMoving), zzBase)
      Call Ex_Meth.ZZToZZ(zzTick, CInt(zzFirst_MinMoving), zzFirst)
      Call Ex_Meth.ZZToZZ(zzTick, CInt(zzSecond_MinMoving), zzSecond)
      
      Set BaseFirstJoned_Settings = Nothing
      Call DS_Tools.FullJoin( _
        zzBase, settings("data_sets")("zz_pack_ds"), _
        zzFirst, settings("data_sets")("zz_pack_ds"), _
        BaseFirstJoned, BaseFirstJoned_Settings, _
        settings("data_sets")("base_first_join_rules"))
      
      Set BaseFirstSecondJoned_Settings = Nothing
      Call DS_Tools.FullJoin( _
        BaseFirstJoned, BaseFirstJoned_Settings, _
        zzSecond, settings("data_sets")("zz_pack_ds"), _
        BaseFirstSecondJoned, BaseFirstSecondJoned_Settings, _
        settings("data_sets")("base_first_second_join_rules"))
        
      Call CM.CopyDict(memPlotSettings, settings("plot_settings")("3zz_ds_plot"))
      'Call RW_Ini.PrintSettings(memPlotSettings, "  ")
      Call ChPl.PlotChart(BaseFirstSecondJoned, settings("plot_settings")("3zz_ds_plot"), "PM")
      Call ChPl.ExportCharts(Format(cnt, "000"), exportFileFolder)
    Next cnt
  End If
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub

