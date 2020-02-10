VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSetIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rowCount As Long
Public EndOfDS As Boolean

Private DS As New DataSet
Private colIndex As New Scripting.Dictionary
Private colNames() As String
Private startIndex As Long, stopIndex As Long

Sub Init(Data_Set As DataSet, col_Names As String)
    Dim col_Name As Variant
    Dim colCount As Long
    Set DS = Data_Set
    EndOfDS = False
    rowCount = 0
    
    colCount = 0
    colNames = Split(col_Names, ",")
    Set colIndex = New Scripting.Dictionary
    For Each col_Name In colNames
      colIndex.Add col_Name, colCount
      colCount = colCount + 1
    Next col_Name
    
    If colCount = 0 Then
        Set colIndex = DS.colIndex
    End If
    
    Call DS.CreateRow(colNames)
    
    If DS.rowsCount = 0 Then
      EndOfDS = True
    Else
      Call DS.GetRow(rowCount)
      'rowCount = rowCount + 1
    End If
    
End Sub

Sub InitByIndex(Data_Set As DataSet, col_Names As String, idx() As Long)
    Dim col_Name As Variant
    Dim colCount As Long
    Set DS = Data_Set
    EndOfDS = False
    rowCount = 0
    
    colCount = 0
    colNames = Split(col_Names, ",")
    Set colIndex = New Scripting.Dictionary
    For Each col_Name In colNames
      colIndex.Add col_Name, colCount
      colCount = colCount + 1
    Next col_Name
    
    If colCount = 0 Then
        Set colIndex = DS.colIndex
    End If
    
    Call DS.CreateRow(colNames)
    
    If DS.rowsCount = 0 Then
      EndOfDS = True
    Else
      Call DS.GetRow(idx(rowCount))
      'rowCount = rowCount + 1
    End If
    
End Sub

Sub InitRange(Data_Set As DataSet, col_Names As String, start_Index As Long, stop_Index As Long)
    Dim col_Name As Variant
    Dim colCount As Long
    Set DS = Data_Set
    EndOfDS = False
    rowCount = start_Index
    
    colCount = 0
    colNames = Split(col_Names, ",")
    Set colIndex = New Scripting.Dictionary
    For Each col_Name In colNames
      colIndex.Add col_Name, colCount
      colCount = colCount + 1
    Next col_Name
    
    If colCount = 0 Then
        Set colIndex = DS.colIndex
    End If
    
    Call DS.CreateRow(colNames)
    
    If DS.rowsCount = 0 Then
      EndOfDS = True
    Else
      Call DS.GetRow(rowCount)
      'rowCount = rowCount + 1
    End If
    
    startIndex = start_Index
    stopIndex = stop_Index
    
End Sub

Sub NextRow()
    Dim col_Name As Variant
    rowCount = rowCount + 1
    If rowCount > DS.rowsCount - 1 Then
      EndOfDS = True
      rowCount = rowCount - 1
      Exit Sub
    Else
      Call DS.GetRow(rowCount)
    End If
    
End Sub

Sub NextRowByIndex(idx() As Long)
    Dim col_Name As Variant
    rowCount = rowCount + 1
    If rowCount > DS.rowsCount - 1 Then
      EndOfDS = True
      rowCount = rowCount - 1
      Exit Sub
    Else
      Call DS.GetRow(idx(rowCount))
    End If
    
End Sub

Sub NextRangeRow()
    Dim col_Name As Variant
    rowCount = rowCount + 1
    If rowCount > stopIndex Then
      EndOfDS = True
      rowCount = rowCount - 1
      Exit Sub
    Else
      Call DS.GetRow(rowCount)
    End If
    
End Sub

