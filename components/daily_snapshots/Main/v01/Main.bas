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
  iniFilesString = "" & _
    "settings\daily_snapshots\main.ini" & "," & _
    "settings\daily_snapshots\zz_pack_ds.ini" & "," & _
    "settings\daily_snapshots\tick_ds.ini"
  iniFiles = Split(iniFilesString, ",")
  
  cmpCount = 0
  Set settings = New Scripting.Dictionary
  For Each iniFile In iniFiles
    iniFiles(cmpCount) = thisWbFolder + iniFile
    Call RW_Ini.ReadSettings(iniFiles(cmpCount), settings)
    
    If My_Err.errOccured Then
      Exit Sub
    End If
    cmpCount = cmpCount + 1
  Next iniFile
  Call RW_Ini.ComposeSettings(settings, settings)
End Sub

Sub Run()
  Dim thisWbFolder As String
  Dim tickCSV As New DataSet
  Dim zzBase As New DataSet
  Dim zzFirst As New DataSet
  Dim zzSecond As New DataSet
  Dim Ex_Meth As New ExchangeMethods
  Dim tick_filePath, zzBase_filePath, zzFirst_filePath, zzSecond_filePath As String
  
  
  thisWbFolder = ThisWorkbook.Path & "\"
  
  Call Init
  If Not My_Err.errOccured Then
    tick_filePath = thisWbFolder & settings("input")("file_folder") & settings("input")("tick_file_name")
    Call tickCSV.ReadFromFile(tick_filePath, settings("data_sets")("tick_ds"))
  End If
  If Not My_Err.errOccured Then
    Call zzBase.Create(settings("data_sets")("zz_pack_ds"))
    Call zzFirst.Create(settings("data_sets")("zz_pack_ds"))
    Call zzSecond.Create(settings("data_sets")("zz_pack_ds"))
    Call Ex_Meth.TicksToZZ(tickCSV, 10, zzBase)
    Call Ex_Meth.ZZToZZ(zzBase, 50, zzFirst)
    Call Ex_Meth.ZZToZZ(zzFirst, 100, zzSecond)
    
 
    zzBase_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_base_file_name")
    zzFirst_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_first_file_name")
    zzSecond_filePath = thisWbFolder & settings("output")("file_folder") & settings("output")("zz_second_file_name")
    Call zzBase.WriteToFile(zzBase_filePath, settings("data_sets")("zz_pack_ds"))
    Call zzFirst.WriteToFile(zzFirst_filePath, settings("data_sets")("zz_pack_ds"))
    Call zzSecond.WriteToFile(zzSecond_filePath, settings("data_sets")("zz_pack_ds"))
  End If
  
  
  If Not My_Err.errOccured Then
  End If
  
  If My_Err.errOccured Then
    Debug.Print My_Err.errMessage
    Exit Sub
  End If
  Debug.Print "OK!"
End Sub
