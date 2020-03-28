VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub CalcDistribution(Data_Set As DataSet, groupColName As String, aggColName As String, step As Variant, Hist_DS As DataSet)
  Dim sortedIndex() As Long
  'Dim ColNames() As String
  Dim DS_Iterator As New DataSetIterator
  Dim maxVal, minVal, startRange, stopRange, currVal, nextVal, aggVal As Variant
  Dim hdsRowCount As Long
  
'  ReDim ColNames(1)
'  ColNames(0) = groupColName
'  ColNames(1) = aggColName

  Call Data_Set.SortIndexByCol(groupColName, sortedIndex, "asc")
  Call DS_Iterator.InitByIndex(Data_Set, groupColName & "," & aggColName, sortedIndex)
  maxVal = Data_Set.GetCell(groupColName, sortedIndex(Data_Set.rowsCount - 1))
  minVal = Data_Set.GetCell(groupColName, sortedIndex(0))
  startRange = minVal - minVal Mod step
  stopRange = maxVal + step - maxVal Mod step
  
  currVal = startRange
  hdsRowCount = 0
  While currVal < stopRange
    nextVal = currVal + step
    aggVal = 0
    While Data_Set.row(groupColName) < nextVal And Not DS_Iterator.EndOfDS
      aggVal = aggVal + Data_Set.row(aggColName)
      Call DS_Iterator.NextRowByIndex(sortedIndex)
    Wend
    
    Call Hist_DS.SetCell("<STEPS>", hdsRowCount, nextVal)
    Call Hist_DS.SetCell("<VAL>", hdsRowCount, aggVal)
    
    'Debug.Print currVal, nextVal, aggVal
    hdsRowCount = hdsRowCount + 1
    currVal = nextVal
  Wend

  Hist_DS.rowsCount = hdsRowCount
End Sub
