VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExchangeMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function AggTickVol(DS As DataSet, startIndex As Long, stopIndex As Long) As Long()
  Dim ret(2) As Long
  Dim DS_i As New DataSetIterator
  Dim bVol, sVol, vol As Long
  
  Call DS_i.InitRange(DS, "<DATE>,<TIME>,<LAST>,<VOL>,<OPER>", startIndex, stopIndex)
  bVol = 0
  sVol = 0
  vol = 0
  While Not DS_i.EndOfDS
    If DS.row("<OPER>") = "B" Then
      bVol = bVol + DS.row("<VOL>")
    End If
    If DS.row("<OPER>") = "S" Then
      sVol = sVol + DS.row("<VOL>")
    End If
    vol = vol + DS.row("<VOL>")
    Call DS_i.NextRangeRow
  Wend
  
  ret(0) = bVol
  ret(1) = sVol
  ret(2) = vol
  
  AggTickVol = ret
  
End Function

Sub TicksToZZ(DS As DataSet, MinMoving As Long, zzDS As DataSet)
  Dim DS_i As New DataSetIterator
  Dim ZZA As New ZigZagAbsolute
  
  Dim zzDSRowCount, lastRowCount As Long
  Dim vols() As Long
  
  Call ZZA.Init(MinMoving)
    zzDSRowCount = 0
    Call DS_i.Init(DS, "<DATE>,<TIME>,<LAST>,<VOL>,<OPER>")
    
    Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS_i.rowCount)
    Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.row("<DATE>"))
    Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.row("<TIME>"))
    Call zzDS.SetCell("<LAST>", zzDSRowCount, DS.row("<LAST>"))
    Call zzDS.SetCell("<B_VOL>", zzDSRowCount, 0)
    Call zzDS.SetCell("<S_VOL>", zzDSRowCount, 0)
    Call zzDS.SetCell("<VOL>", zzDSRowCount, 0)
    zzDSRowCount = zzDSRowCount + 1
    lastRowCount = DS_i.rowCount
    
    While Not DS_i.EndOfDS
      'DS.row ("<DATA>"), DS.row("<TIME>"), DS.row("<LAST>"), DS.row("<VOL>"
      Call ZZA.findExpectedExtrems(DS_i.rowCount, DS.row("<LAST>"), DS.row("<LAST>"))
      Call ZZA.confirmExpectedExtrems
      If ZZA.isMax Then
        vols = AggTickVol(DS, lastRowCount, DS_i.rowCount)
        
        Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, ZZA.maxIndex)
        Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.GetCell("<DATE>", ZZA.maxIndex))
        Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.GetCell("<TIME>", ZZA.maxIndex))
        Call zzDS.SetCell("<LAST>", zzDSRowCount, ZZA.maxVal)
        Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
        Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
        Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
        lastRowCount = DS_i.rowCount
        zzDSRowCount = zzDSRowCount + 1
      End If
      
      If ZZA.isMin Then
        vols = AggTickVol(DS, lastRowCount, DS_i.rowCount)
        Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, ZZA.minIndex)
        Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.GetCell("<DATE>", ZZA.minIndex))
        Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.GetCell("<TIME>", ZZA.minIndex))
        Call zzDS.SetCell("<LAST>", zzDSRowCount, ZZA.minVal)
        Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
        Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
        Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
        lastRowCount = DS_i.rowCount
        zzDSRowCount = zzDSRowCount + 1
      End If
      
      Call DS_i.NextRow
    Wend
    vols = AggTickVol(DS, lastRowCount, DS_i.rowCount)
    Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS_i.rowCount)
    Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.row("<DATE>"))
    Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.row("<TIME>"))
    Call zzDS.SetCell("<LAST>", zzDSRowCount, DS.row("<LAST>"))
    Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
    Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
    Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
    zzDS.rowsCount = zzDSRowCount + 1
  
End Sub

Function AggZZVol(DS As DataSet, startIndex As Long, stopIndex As Long) As Long()
  Dim ret(2) As Long
  Dim DS_i As New DataSetIterator
  Dim bVol, sVol, vol As Long
  
  Call DS_i.InitRange(DS, "<MAIN_INDEX>,<DATE>,<TIME>,<LAST>,<B_VOL>,<S_VOL>,<VOL>", startIndex, stopIndex)
  bVol = 0
  sVol = 0
  vol = 0
  While Not DS_i.EndOfDS
    bVol = bVol + DS.row("<B_VOL>")
    sVol = sVol + DS.row("<S_VOL>")
    vol = vol + DS.row("<VOL>")
    Call DS_i.NextRangeRow
  Wend
  
  ret(0) = bVol
  ret(1) = sVol
  ret(2) = vol
  
  AggZZVol = ret
  
End Function

Sub ZZToZZ(DS As DataSet, MinMoving As Long, zzDS As DataSet)
  Dim DS_i As New DataSetIterator
  Dim ZZA As New ZigZagAbsolute
  
  Dim zzDSRowCount, lastRowCount As Long
  Dim vols() As Long
  
  Call ZZA.Init(MinMoving)
    zzDSRowCount = 0
    Call DS_i.Init(DS, "<MAIN_INDEX>,<DATE>,<TIME>,<LAST>,<B_VOL>,<S_VOL>,<VOL>")
    
    Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS.row("<MAIN_INDEX>"))
    Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.row("<DATE>"))
    Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.row("<TIME>"))
    Call zzDS.SetCell("<LAST>", zzDSRowCount, DS.row("<LAST>"))
    Call zzDS.SetCell("<B_VOL>", zzDSRowCount, 0)
    Call zzDS.SetCell("<S_VOL>", zzDSRowCount, 0)
    Call zzDS.SetCell("<VOL>", zzDSRowCount, 0)
    zzDSRowCount = zzDSRowCount + 1
    lastRowCount = DS_i.rowCount
    
    While Not DS_i.EndOfDS
      'DS.row ("<DATA>"), DS.row("<TIME>"), DS.row("<LAST>"), DS.row("<VOL>"
      Call ZZA.findExpectedExtrems(DS_i.rowCount, DS.row("<LAST>"), DS.row("<LAST>"))
      Call ZZA.confirmExpectedExtrems
      If ZZA.isMax Then
        vols = AggZZVol(DS, lastRowCount, DS_i.rowCount)
        
        Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS.row("<MAIN_INDEX>"))
        Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.GetCell("<DATE>", ZZA.maxIndex))
        Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.GetCell("<TIME>", ZZA.maxIndex))
        Call zzDS.SetCell("<LAST>", zzDSRowCount, ZZA.maxVal)
        Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
        Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
        Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
        lastRowCount = DS_i.rowCount
        zzDSRowCount = zzDSRowCount + 1
      End If
      
      If ZZA.isMin Then
        vols = AggZZVol(DS, lastRowCount, DS_i.rowCount)
        Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS.row("<MAIN_INDEX>"))
        Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.GetCell("<DATE>", ZZA.minIndex))
        Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.GetCell("<TIME>", ZZA.minIndex))
        Call zzDS.SetCell("<LAST>", zzDSRowCount, ZZA.minVal)
        Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
        Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
        Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
        lastRowCount = DS_i.rowCount
        zzDSRowCount = zzDSRowCount + 1
      End If
      
      Call DS_i.NextRow
    Wend
    vols = AggZZVol(DS, lastRowCount, DS_i.rowCount)
    Call zzDS.SetCell("<MAIN_INDEX>", zzDSRowCount, DS.row("<MAIN_INDEX>"))
    Call zzDS.SetCell("<DATE>", zzDSRowCount, DS.row("<DATE>"))
    Call zzDS.SetCell("<TIME>", zzDSRowCount, DS.row("<TIME>"))
    Call zzDS.SetCell("<LAST>", zzDSRowCount, DS.row("<LAST>"))
    Call zzDS.SetCell("<B_VOL>", zzDSRowCount, vols(0))
    Call zzDS.SetCell("<S_VOL>", zzDSRowCount, vols(1))
    Call zzDS.SetCell("<VOL>", zzDSRowCount, vols(2))
    zzDS.rowsCount = zzDSRowCount + 1

End Sub
