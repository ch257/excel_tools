VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rowIndex() As Long
Private colNames() As String
Private table() As Variant
Public row As New Scripting.Dictionary
Public colIndex As New Scripting.Dictionary
Public colsCount As Long
Public rowsCount As Long
Public dataLength As Long

Sub Create(ByVal dsSettings As Scripting.Dictionary)
  Dim colName As Variant
  Dim colCount, rowCount As Long
  Dim emptyArr() As Variant
  
  Set colIndex = New Scripting.Dictionary
  colNames = Split(dsSettings("columns"), ",")
  colCount = 0
  For Each colName In colNames
    colIndex.Add colName, colCount
    colCount = colCount + 1
  Next colName
  colsCount = colCount
  rowsCount = 0
  dataLength = dsSettings("data_set_length")
  
  ReDim table(colsCount - 1)
  ReDim colNames(colsCount - 1)
  ReDim rowIndex(dataLength - 1)
  ReDim emptyArr(dataLength - 1)
  For rowCount = 0 To dataLength - 1
    rowIndex(rowCount) = rowCount
    emptyArr(rowCount) = Null
  Next rowCount
  For Each colName In colIndex
    table(colIndex(colName)) = emptyArr
    colNames(colIndex(colName)) = colName
  Next colName
  
End Sub

'Sub Reset()
'  Dim emptyArr() As Variant
'  Dim colName As Variant
'  Dim rowCount As Long
'
'  ReDim emptyArr(dataLength - 1)
'  For rowCount = 0 To dataLength - 1
'    emptyArr(rowCount) = Null
'  Next rowCount
'  For Each colName In colIndex
'    table(colIndex(colName)) = emptyArr
'  Next colName
'  rowsCount = 0
'End Sub
'
Sub ReadFromFile(ByVal filePath As String, ByVal dsSettings As Scripting.Dictionary)

  Dim RW_File As New RWFile
  Dim line As String
  Dim lineParts() As String
  Dim colName As Variant
  Dim Type_Convertor As New TypeConvertor
  Dim Settings_Reader As New SettingsReader
  Dim rowCount, colCount As Long
  Dim colType As String

  Call RW_File.OpenFile(filePath, "r")
  If My_Err.errOccured Then
    Exit Sub
  End If
  line = RW_File.ReadLine
  dsSettings("columns") = Join(Split(line, dsSettings("column_separator")), ",")
  Call Create(dsSettings)

  rowCount = 0
  While Not RW_File.endOfFile And rowCount < dataLength
    line = RW_File.ReadLine
    If line <> "" Then
      lineParts = Split(line, dsSettings("column_separator"))
      For Each colName In colIndex
        colCount = colIndex(colName)
        'colType = dsSettings("column_type")(colName)
        colType = Settings_Reader.GetValue(dsSettings("column_type"), CStr(colName))
        table(colCount)(rowCount) = Type_Convertor.Cast(lineParts(colCount), colType, dsSettings("decimal"))
      Next colName
      rowCount = rowCount + 1
    End If
  Wend
  rowsCount = rowCount
  Call RW_File.CloseFile
  If Not RW_File.endOfFile Then
    My_Err.RiseError ("No enough data length for file " & filePath)
  End If
End Sub

Sub WriteToFile(ByVal filePath As String, ByVal dsSettings As Scripting.Dictionary)
  Dim RW_File As New RWFile
  Dim line As String
  Dim lineParts() As String
  Dim colName As Variant
  Dim Type_Convertor As New TypeConvertor
  Dim Settings_Reader As New SettingsReader
  Dim rowCount, colCount As Long
  Dim colFormat As String
  
  Call RW_File.OpenFile(filePath, "w")
  If My_Err.errOccured Then
    Exit Sub
  End If
  
  line = ""
  For Each colName In colNames
    colCount = colIndex(colName)
    line = line & dsSettings("column_separator") & colName
  Next colName
  line = Mid(line, 2)
  Call RW_File.WriteLine(line)
  
  rowCount = 0
  While rowCount < rowsCount
    line = ""
    For Each colName In colNames
      colCount = colIndex(colName)
      colFormat = Settings_Reader.GetValue(dsSettings("column_format"), CStr(colName))
      line = line & dsSettings("column_separator") & Type_Convertor.FormatV(table(colCount)(rowCount), colFormat, dsSettings("decimal"))
    Next colName
    line = Mid(line, 2)
    Call RW_File.WriteLine(line)
    rowCount = rowCount + 1
  Wend
  
  
  Call RW_File.CloseFile

End Sub

Function GetCell(ByVal colName As String, ByVal rowCount As Long) As Variant
  GetCell = table(colIndex(colName))(rowCount)
End Function

Sub SetCell(ByVal colName As String, ByVal rowCount As Long, val As Variant)
  table(colIndex(colName))(rowCount) = val
End Sub

Public Sub CreateRow(col_Names() As String)
  Dim colName As Variant
  Set row = New Scripting.Dictionary
  For Each colName In col_Names
    row.Add colName, Null
  Next colName
End Sub

'Sub SetRow(ByVal rowCount As Long, row As Scripting.Dictionary)
'  Dim colName As Variant
'  For Each colName In row
'    table(colIndex(CStr(colName)))(rowCount) = row(CStr(colName))
'  Next colName
'End Sub

Sub GetRow(ByVal rowCount As Long)
  Dim colName As Variant
  For Each colName In row
    row(colName) = table(colIndex(CStr(colName)))(rowCount)
  Next colName
End Sub

'Function GetColNames() As String()
'  GetColNames = ColNames
'End Function

Function GetRowIndex() As Long()
  GetRowIndex = rowIndex
  ReDim Preserve GetRowIndex(rowsCount - 1)
End Function

'Function GetCol(ByVal colName As String) As Variant()
'  GetCol = table(colIndex(colName))
'  ReDim Preserve GetCol(rowsCount - 1)
'End Function
'
'Function GetIndexByColValue(ByVal colName As String, val As Variant) As Long
'  Dim colCount, rowCount As Long
'  Dim col_val As Variant
'  rowCount = 0
'  For Each col_val In table(colIndex(colName))
'    If col_val > val Then
'      GetIndexByColValue = rowCount
'      Exit Function
'    End If
'    rowCount = rowCount + 1
'  Next col_val
'  GetIndexByColValue = -1
'End Function
'
'Sub SortIndexByCol(ByVal colName As String, ByRef sortedIndex() As Long, order As String)
'  Dim colCopy() As Variant
'  Dim i, j, extr_i, tmpIdx As Long
'  Dim tmp As Variant
'
'  sortedIndex = GetRowIndex()
'  colCopy = GetCol(colName)
'
'  If order = "desc" Then
'    For i = 0 To rowsCount - 1 - 1
'      extr_i = i
'      For j = i + 1 To rowsCount - 1
'        If colCopy(extr_i) < colCopy(j) Then
'            extr_i = j
'        End If
'      Next j
'      tmp = colCopy(i)
'      colCopy(i) = colCopy(extr_i)
'      colCopy(extr_i) = tmp
'
'      tmpIdx = sortedIndex(i)
'      sortedIndex(i) = sortedIndex(extr_i)
'      sortedIndex(extr_i) = tmpIdx
'    Next i
'  Else
'    For i = 0 To rowsCount - 1 - 1
'      extr_i = i
'      For j = i + 1 To rowsCount - 1
'        If colCopy(extr_i) > colCopy(j) Then
'            extr_i = j
'        End If
'      Next j
'      tmp = colCopy(i)
'      colCopy(i) = colCopy(extr_i)
'      colCopy(extr_i) = tmp
'
'      tmpIdx = sortedIndex(i)
'      sortedIndex(i) = sortedIndex(extr_i)
'      sortedIndex(extr_i) = tmpIdx
'    Next i
'  End If
'
'End Sub
'
'Sub SortIndexRangeByCol(ByVal colName As String, startIndex As Long, stopIndex As Long, ByRef sortedIndex() As Long, ByVal order As String)
'  Dim colCopy() As Variant
'  Dim i, j, extr_i, tmpIdx As Long
'  Dim tmp As Variant
'
'  sortedIndex = GetRowIndex()
'  colCopy = GetCol(colName)
'
'  If order = "desc" Then
'    For i = startIndex To stopIndex - 1
'      extr_i = i
'      For j = i + 1 To stopIndex
'        If colCopy(extr_i) < colCopy(j) Then
'            extr_i = j
'        End If
'      Next j
'      tmp = colCopy(i)
'      colCopy(i) = colCopy(extr_i)
'      colCopy(extr_i) = tmp
'
'      tmpIdx = sortedIndex(i)
'      sortedIndex(i) = sortedIndex(extr_i)
'      sortedIndex(extr_i) = tmpIdx
'    Next i
'  Else
'    For i = startIndex To stopIndex - 1
'      extr_i = i
'      For j = i + 1 To stopIndex
'        If colCopy(extr_i) > colCopy(j) Then
'            extr_i = j
'        End If
'      Next j
'      tmp = colCopy(i)
'      colCopy(i) = colCopy(extr_i)
'      colCopy(extr_i) = tmp
'
'      tmpIdx = sortedIndex(i)
'      sortedIndex(i) = sortedIndex(extr_i)
'      sortedIndex(extr_i) = tmpIdx
'    Next i
'  End If
'
'End Sub



