VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeConvertor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Cast(val As Variant, vType As String, dec As String) As Variant
  If val = "" Then
    'Cast = Null
    Cast = Empty
    Exit Function
  End If
  If vType = "num" Then
    val = Replace(val, dec, Application.DecimalSeparator)
    Cast = CDbl(val)
    Exit Function
  End If
  If vType = "int" Then
    Cast = CLng(val)
    Exit Function
  End If
  If vType = "date" Then
    Cast = CLng(val)
    Exit Function
  End If
  If vType = "time" Then
    Cast = CLng(val)
    Exit Function
  End If
  If vType = "str" Then
    Cast = val
    Exit Function
  End If
End Function

Function FormatV(val As Variant, vFormat As String, dec As String) As String
  'Dim s As String
  'If IsNull(val) Then
  If IsEmpty(val) Then
    FormatV = ""
    Exit Function
  End If
  If vFormat = "yyyymmdd" Then
    FormatV = Format(CStr(val))
    Exit Function
  End If
  If vFormat = "hhmmss" Then
    FormatV = Format(CStr(val))
    Exit Function
  End If
  If InStr(vFormat, "0.") > 0 Then
    FormatV = Format(val, vFormat)
    FormatV = Replace(FormatV, ",", dec)
    Exit Function
  End If
  FormatV = Format(val, vFormat)
End Function
