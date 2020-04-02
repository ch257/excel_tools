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

Private Sub CalcZZ()
  Dim thisWbFolder As String
  Dim zzTick As New DataSet
  Dim zzTick_fileFolder As String
  Dim zzTick_filePath As String
  Dim zzTick_fileList() As String
  Dim SelectedZZTick As New DataSet
  Dim zzBase As New DataSet
  Dim Ex_Meth As New ExchangeMethods
  Dim zzMinMoving As Integer
  
  Dim zzMinMoving_Start As Integer
  Dim zzMinMoving_Stop As Integer
  Dim zzMinMoving_Step As Integer
  
  Dim DS_Tools As New DataSetTools
  Dim RW_File As New RWFile
  Dim cnt As Integer
  Dim CM As New CommonMethods
  'Dim SM As New StatMethods
  Dim volatilityDS As New DataSet
  
  thisWbFolder = ThisWorkbook.Path & "\"
    
  zzTick_fileFolder = thisWbFolder & settings("input")("file_folder")
  zzMinMoving_Start = settings("parameters")("zz_min_moving_start")
  zzMinMoving_Stop = settings("parameters")("zz_min_moving_stop")
  zzMinMoving_Step = settings("parameters")("zz_min_moving_step")
  
  For zzMinMoving = zzMinMoving_Start To zzMinMoving_Stop Step zzMinMoving_Step
    Call volatilityDS.Create(settings("data_sets")("volatility_ds"))
  
    zzTick_fileList = RW_File.GetFolderFileList(zzTick_fileFolder)
    For cnt = 1 To UBound(zzTick_fileList) - LBound(zzTick_fileList)
      
      zzTick_filePath = zzTick_fileFolder & zzTick_fileList(cnt)
      Call zzTick.ReadFromFile(zzTick_filePath, settings("data_sets")("zz_pack_ds"))
      If My_Err.errOccured Then
        Exit For
      End If
      Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
      Call Ex_Meth.ZZToZZ(zzTick, CInt(zzMinMoving), zzBase)
      Call DS_Tools.SelectBetween(zzBase, settings("data_sets")("zz_pack_ds"), "<TIME>", 100000, 110000, SelectedZZTick, settings("data_sets")("zz_pack_ds"))
            
      Call volatilityDS.SetCell("<FILENAME>", cnt - 1, zzTick_fileList(cnt))
      Call volatilityDS.SetCell("<SELECTED_LENGTH>", cnt - 1, SelectedZZTick.rowsCount)
      Call volatilityDS.SetCell("<TOTAL_LENGTH>", cnt - 1, zzBase.rowsCount)
      
'        cnt = cnt + 1
'        Exit For
    Next cnt
    
    volatilityDS.rowsCount = cnt - 1
    Call volatilityDS.WriteToSheet(CStr(zzMinMoving), settings("data_sets")("zz_pack_ds"))
  Next zzMinMoving
End Sub

Private Sub CalcLinear(XColNumber, YColNumber, startRowNumber, rowsNumber)
  Dim colLetters() As String
  
  colLetters = Split("0,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z", ",")
  
  'ThisWorkbook.Sheets("SUM").Range("A1:C5").FormulaArray = "=LINEST(" & colLetters(YColNumber) & startRowNumber & ":" & colLetters(YColNumber) & rowsNumber & "," & colLetters(XColNumber) & startRowNumber & ":" & colLetters(XColNumber) & rowsNumber & "^COLUMN(A:B),1,1)"
  ThisWorkbook.Sheets("SUM").Range("A1:B5").FormulaArray = "=LINEST(" & colLetters(YColNumber) & startRowNumber & ":" & colLetters(YColNumber) & rowsNumber & "," & colLetters(XColNumber) & startRowNumber & ":" & colLetters(XColNumber) & rowsNumber & ",1,1)"

  ThisWorkbook.Sheets("SUM").Cells(1, YColNumber) = ThisWorkbook.Sheets("SUM").Cells(1, 1)
  ThisWorkbook.Sheets("SUM").Cells(2, YColNumber) = ThisWorkbook.Sheets("SUM").Cells(1, 2)
  ThisWorkbook.Sheets("SUM").Cells(3, YColNumber) = ThisWorkbook.Sheets("SUM").Cells(1, 3)
  ThisWorkbook.Sheets("SUM").Cells(4, YColNumber) = ThisWorkbook.Sheets("SUM").Cells(3, 1)
  ThisWorkbook.Sheets("SUM").Cells(6, YColNumber).Formula = "=(" & colLetters(XColNumber) & CStr(startRowNumber - 3) & "-" & colLetters(YColNumber) & CStr(startRowNumber - 7) & ")/" & colLetters(YColNumber) & CStr(startRowNumber - 8)
  
  ThisWorkbook.Sheets("SUM").Range("A1:C5").ClearContents
End Sub

Private Sub CalcSUM()
  Dim zzMinMoving_Start As Integer
  Dim zzMinMoving_Stop As Integer
  Dim zzMinMoving_Step As Integer
  Dim zzMinMoving As Integer
  Dim rowCount As Long
  Dim rowsCount As Long
  Dim sheetCount As Integer
  Dim startColNumber As Integer
  Dim startRowNumber As Integer
  
  Dim colLetters() As String
  colLetters = Split("0,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z", ",")
    
  zzMinMoving_Start = settings("parameters")("zz_min_moving_start")
  zzMinMoving_Stop = settings("parameters")("zz_min_moving_stop")
  zzMinMoving_Step = settings("parameters")("zz_min_moving_step")
  
  Dim CM As New CommonMethods
  
  Call CM.CreateSheet("SUM")
  Call CM.ClearSheet("SUM")
  
  startColNumber = 2
  startRowNumber = 8
  
  ThisWorkbook.Sheets("SUM").Cells(1, startColNumber + 2) = "a"
  ThisWorkbook.Sheets("SUM").Cells(2, startColNumber + 2) = "b"
  ThisWorkbook.Sheets("SUM").Cells(3, startColNumber + 2) = "c"
  ThisWorkbook.Sheets("SUM").Cells(4, startColNumber + 2) = "R2"
  ThisWorkbook.Sheets("SUM").Cells(6, startColNumber + 2) = 1000
  sheetCount = 1
  zzMinMoving = zzMinMoving_Start
  rowCount = startRowNumber + 1
  While ThisWorkbook.Sheets(CStr(zzMinMoving)).Cells(rowCount, 1) <> Empty
    ThisWorkbook.Sheets("SUM").Cells(rowCount, startColNumber + 1) = ThisWorkbook.Sheets(CStr(zzMinMoving)).Cells(rowCount, 1)
    ThisWorkbook.Sheets("SUM").Cells(rowCount, startColNumber + 2) = ThisWorkbook.Sheets(CStr(zzMinMoving)).Cells(rowCount, 2)
    rowCount = rowCount + 1
  Wend
  rowsCount = rowCount - 1
  
  sheetCount = 1
  For zzMinMoving = zzMinMoving_Start To zzMinMoving_Stop Step zzMinMoving_Step
    rowCount = startRowNumber
    ThisWorkbook.Sheets("SUM").Cells(rowCount, startColNumber + sheetCount + 2) = zzMinMoving
    rowCount = rowCount + 1
    While ThisWorkbook.Sheets(CStr(zzMinMoving)).Cells(rowCount, 1) <> Empty
      ThisWorkbook.Sheets("SUM").Cells(rowCount, startColNumber + sheetCount + 2) = ThisWorkbook.Sheets(CStr(zzMinMoving)).Cells(rowCount, 3)
      rowCount = rowCount + 1
    Wend
    Call CalcLinear(startColNumber + 2, startColNumber + sheetCount + 2, startRowNumber + 1, rowsCount)
    sheetCount = sheetCount + 1
  Next zzMinMoving
  
  ThisWorkbook.Sheets("SUM").Range("A1:C5").FormulaArray = "=LINEST(" & _
    colLetters(startColNumber + 3) & CStr(startRowNumber - 2) & ":" & _
    colLetters(startColNumber + sheetCount + 1) & CStr(startRowNumber - 2) & "," & _
    colLetters(startColNumber + 3) & CStr(startRowNumber) & ":" & _
    colLetters(startColNumber + sheetCount + 1) & CStr(startRowNumber) & "^{1;2},1,1)"
  
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
    
  
  Call Init
  If Not My_Err.errOccured Then
    'Call CalcZZ
    Call CalcSUM
  End If
  
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub

