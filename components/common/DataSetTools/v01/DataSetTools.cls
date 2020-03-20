VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSetTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub InsertSettings(fromDict As Scripting.Dictionary, JoinRules As Scripting.Dictionary, columns As Scripting.Dictionary, toDict As Scripting.Dictionary)
    Dim k, destK, jrK As Variant
    Dim emptyDict As New Scripting.Dictionary
    
    For Each k In fromDict.keys
        destK = k
        jrK = k
        If k = "column_type" Then
          jrK = "cross_column_type"
        ElseIf k = "column_format" Then
          jrK = "cross_column_format"
        ElseIf columns.Exists(k) Then
          destK = columns(k)
        End If
        
        If TypeName(fromDict(k)) = "Dictionary" Then
            If Not toDict.Exists(destK) Then
              Set emptyDict = Nothing
              toDict.Add destK, emptyDict
            End If
            Call InsertSettings(fromDict(k), JoinRules(jrK), columns, toDict(k))
        Else
            If Not toDict.Exists(destK) Then
                If k = "columns" Then
                  toDict.Add destK, JoinRules("cross_columns")
                Else
                  toDict.Add destK, fromDict(k)
                End If
            End If
        End If
    Next
    If toDict.Exists("columns") Then
      For Each k In columns
        toDict("columns") = toDict("columns") & "," & columns(k)
      Next k
    End If
End Sub

Private Sub CreateJoinedDataSet(DS1 As DataSet, DS1_Settings As Scripting.Dictionary, DS2 As DataSet, DS2_Settings As Scripting.Dictionary, JoinedDS As DataSet, JoinedDS_Settings As Scripting.Dictionary, JoinRules As Scripting.Dictionary)
  Dim tst As Variant
  Call InsertSettings(DS1_Settings, JoinRules, JoinRules("ds1_columns"), JoinedDS_Settings)
  Call InsertSettings(DS2_Settings, JoinRules, JoinRules("ds2_columns"), JoinedDS_Settings)
  Call JoinedDS.Create(JoinedDS_Settings)
End Sub

Private Sub FillCross(ByVal dt As Long, ByVal tm As Long, DS1 As DataSet, DS1_rowCount As Long, DS2 As DataSet, DS2_rowCount As Long, JoinRules As Scripting.Dictionary, JoinedDS As DataSet, JoinedDS_rowCount As Long)
  Dim col As Variant
  
  Call JoinedDS.SetCell("<DATE>", JoinedDS_rowCount, dt)
  Call JoinedDS.SetCell("<TIME>", JoinedDS_rowCount, tm)
  
  For Each col In JoinRules("ds1_columns")
    Call JoinedDS.SetCell(JoinRules("ds1_columns")(col), JoinedDS_rowCount, DS1.GetCell(col, DS1_rowCount))
  Next col
  For Each col In JoinRules("ds2_columns")
    Call JoinedDS.SetCell(JoinRules("ds2_columns")(col), JoinedDS_rowCount, DS2.GetCell(col, DS2_rowCount))
  Next col
  
  JoinedDS_rowCount = JoinedDS_rowCount + 1
End Sub

Private Sub FillLeft(ByVal dt As Long, ByVal tm As Long, DS1 As DataSet, DS1_rowCount As Long, DS2 As DataSet, DS2_rowCount As Long, JoinRules As Scripting.Dictionary, JoinedDS As DataSet, JoinedDS_rowCount As Long)
  Dim col As Variant
  
  Call JoinedDS.SetCell("<DATE>", JoinedDS_rowCount, dt)
  Call JoinedDS.SetCell("<TIME>", JoinedDS_rowCount, tm)
  
  For Each col In JoinRules("ds1_columns")
    Call JoinedDS.SetCell(JoinRules("ds1_columns")(col), JoinedDS_rowCount, DS1.GetCell(col, DS1_rowCount))
  Next col
  For Each col In JoinRules("ds2_columns")
  '  Call JoinedDS.SetCell(JoinRules("ds2_columns")(col), JoinedDS_rowCount, Null)
    Call JoinedDS.SetCell(JoinRules("ds2_columns")(col), JoinedDS_rowCount, Empty)
  Next col
  
  JoinedDS_rowCount = JoinedDS_rowCount + 1
End Sub

Private Sub FillRight(ByVal dt As Long, ByVal tm As Long, DS1 As DataSet, DS1_rowCount As Long, DS2 As DataSet, DS2_rowCount As Long, JoinRules As Scripting.Dictionary, JoinedDS As DataSet, JoinedDS_rowCount As Long)
  Dim col As Variant
  
  Call JoinedDS.SetCell("<DATE>", JoinedDS_rowCount, dt)
  Call JoinedDS.SetCell("<TIME>", JoinedDS_rowCount, tm)
  
  For Each col In JoinRules("ds1_columns")
  '  Call JoinedDS.SetCell(JoinRules("ds1_columns")(col), JoinedDS_rowCount, Null)
    Call JoinedDS.SetCell(JoinRules("ds1_columns")(col), JoinedDS_rowCount, Empty)
  Next col
  For Each col In JoinRules("ds2_columns")
    Call JoinedDS.SetCell(JoinRules("ds2_columns")(col), JoinedDS_rowCount, DS2.GetCell(col, DS2_rowCount))
  Next col
  
  JoinedDS_rowCount = JoinedDS_rowCount + 1
End Sub

Sub FullJoin(DS1 As DataSet, ByVal DS1_Settings As Scripting.Dictionary, DS2 As DataSet, ByVal DS2_Settings As Scripting.Dictionary, JoinedDS As DataSet, JoinedDS_Settings As Scripting.Dictionary, ByVal JoinRules As Scripting.Dictionary)
  Dim DS1_i As New DataSetIterator
  Dim DS2_i As New DataSetIterator
  Dim JoinedDS_rowCount As Long
  Dim exitWhile, DecCount As Boolean
  Dim dt1, dt2 As Long
  Dim tm1, tm2 As Long
    
  Call CreateJoinedDataSet(DS1, DS1_Settings, DS2, DS2_Settings, JoinedDS, JoinedDS_Settings, JoinRules)
  
  Call DS1_i.Init(DS1, DS1_Settings("columns"))
  Call DS2_i.Init(DS2, DS2_Settings("columns"))
  
  JoinedDS_rowCount = 0
  While Not DS1_i.EndOfDS
    dt1 = DS1.GetCell("<DATE>", DS1_i.rowCount)
    tm1 = DS1.GetCell("<TIME>", DS1_i.rowCount)
    
    exitWhile = False
    DecCount = False
    While Not DS2_i.EndOfDS And Not exitWhile
      dt2 = DS2.GetCell("<DATE>", DS2_i.rowCount)
      tm2 = DS2.GetCell("<TIME>", DS2_i.rowCount)
      
      If dt1 = dt2 Then
        If tm1 = tm2 Then
          Call FillCross(dt1, tm1, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
          exitWhile = True
        ElseIf tm1 < tm2 Then
          Call FillLeft(dt1, tm1, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
          'DS2_i.rowCount = DS2_i.rowCount - 1 '!!!!!!!!!!!!!!!!!! -1
          exitWhile = True
          DecCount = True
        Else
          Call FillRight(dt2, tm2, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
        End If
      ElseIf dt1 < dt2 Then
        Call FillLeft(dt2, tm2, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
        'DS2_i.rowCount = DS2_i.rowCount - 1 '!!!!!!!!!!!!!!!!!! -1
        exitWhile = True
        DecCount = True
      Else
        Call FillRight(dt2, tm2, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
      End If
      
      If Not DecCount = True Then
        Call DS2_i.NextRow
      End If
    Wend
    
    Call DS1_i.NextRow
  Wend
  
  While Not DS1_i.EndOfDS
    dt1 = DS2.GetCell("<DATE>", DS1_i.rowCount)
    tm1 = DS2.GetCell("<TIME>", DS1_i.rowCount)
    Call FillLeft(dt1, tm1, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
    
    Call DS1_i.NextRow
  Wend
  
   While Not DS2_i.EndOfDS
    dt2 = DS2.GetCell("<DATE>", DS2_i.rowCount)
    tm2 = DS2.GetCell("<TIME>", DS2_i.rowCount)
    Call FillRight(dt2, tm2, DS1, DS1_i.rowCount, DS2, DS2_i.rowCount, JoinRules, JoinedDS, JoinedDS_rowCount)
    
    Call DS2_i.NextRow
  Wend
  
  JoinedDS.rowsCount = JoinedDS_rowCount '- 1
  
End Sub
