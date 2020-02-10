VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub CopyDict(ByRef fromDict As Scripting.Dictionary, ByRef toDict As Scripting.Dictionary)
    Dim k As Variant
    Dim emptyDict As New Scripting.Dictionary
    For Each k In fromDict.keys
        If TypeName(fromDict(k)) = "Dictionary" Then
            Set emptyDict = Nothing
            toDict.Add k, emptyDict
            Call CopyDict(fromDict(k), toDict(k))
        Else
            If toDict.Exists(k) Then
                toDict(k) = fromDict(k)
            Else
                toDict.Add k, fromDict(k)
            End If
        End If
    Next
End Sub

Sub SplitStrToCollection(line As String, delimeter As String, coll As Collection)
    Dim arr() As String
    Dim itm As Variant
    
    Set coll = New Collection
    arr = Split(line, delimeter)
    For Each itm In arr
        coll.Add itm
    Next itm
End Sub

Sub CreateRowByColNames(col_Names() As String, ByRef row As Scripting.Dictionary)
  Dim colName As Variant
  Set row = New Scripting.Dictionary
  For Each colName In col_Names
    row.Add colName, Null
  Next colName
End Sub

Sub CreateRowByColIndex(col_Index As Scripting.Dictionary, ByRef row As Scripting.Dictionary)
  Dim colName As Variant
  Set row = New Scripting.Dictionary
  For Each colName In col_Index
    row.Add colName, Null
  Next colName
End Sub

