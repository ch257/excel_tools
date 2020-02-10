VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public buffSize As Long
Private buffItems() As Variant
Public isReady As Boolean

Sub Create(size As Long)
    buffSize = size
    ReDim buffItems(size - 1)
    isReady = False
End Sub

Function GetItem(idx As Long) As Variant
    GetItem = buffItems(idx)
End Function

Sub SetItem(ByVal idx As Long, val As Variant)
    buffItems(idx) = val
End Sub

Sub Slide(newItem As Variant)
    Dim i As Long
    For i = 0 To buffSize - 1 - 1
        buffItems(i) = buffItems(i + 1)
    Next i
    buffItems(buffSize - 1) = newItem
    If buffItems(0) <> Empty Then
        isReady = True
    End If
End Sub
