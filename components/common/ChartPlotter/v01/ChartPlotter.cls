VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartPlotter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#If VBA7 Then
'  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
'#Else
'  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
'#End If


Sub DeleteCharts()
  Dim chtObj As ChartObject
  For Each chtObj In ActiveSheet.ChartObjects
    chtObj.Delete
  Next
End Sub

Sub CreateLineChart(AxesParams As Scripting.Dictionary, SeriaValues, SeriaXValues)
  
  Dim objCht As ChartObject
  Set objCht = Sheets(AxesParams("SheetName")).ChartObjects.Add( _
    AxesParams("topXPosition"), _
    AxesParams("topYposition"), _
    AxesParams("chartWidth"), _
    AxesParams("chartHeight") _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = AxesParams("Y_min")
  objCht.Chart.Axes(xlValue).MaximumScale = AxesParams("Y_max")
  
  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.Axes(xlValue).Delete
  objCht.Chart.Axes(xlValue).MajorGridlines.Delete
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = SeriaXValues
  objCht.Chart.DisplayBlanksAs = xlInterpolated
  
  objCht.Chart.Legend.Delete
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.Width = AxesParams("chartWidth") - AxesParams("PlotAreaMargin") - 12
  objCht.Chart.PlotArea.Height = AxesParams("chartHeight") - AxesParams("PlotAreaMargin")
  objCht.Chart.PlotArea.Left = 21 + 6 '+4 from Office chart
  objCht.Chart.PlotArea.Top = 21 '+4 from Office chart
  
  objCht.Chart.PlotArea.Format.Fill.Visible = msoFalse
  Sheets(AxesParams("SheetName")).Shapes(Mid(objCht.Chart.Name, Len(AxesParams("SheetName")) + 2)).Fill.Visible = msoFalse
  Sheets(AxesParams("SheetName")).Shapes(Mid(objCht.Chart.Name, Len(AxesParams("SheetName")) + 2)).line.Visible = msoFalse
  Set objCht = Nothing
End Sub

Sub CreateYAx(AxesParams As Scripting.Dictionary, SeriaXValues)
  
  Dim objCht As ChartObject
  Set objCht = Sheets(AxesParams("SheetName")).ChartObjects.Add( _
    AxesParams("topXPosition") - AxesParams("Y_ax_width"), _
    AxesParams("topYposition"), _
    AxesParams("chartWidth") + AxesParams("Y_ax_width"), _
    AxesParams("chartHeight") _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = AxesParams("Y_min")
  objCht.Chart.Axes(xlValue).MaximumScale = AxesParams("Y_max")
  
  Dim SeriaValues() As Double
  Dim itm As Variant
  Dim cnt As Long
  ReDim SeriaValues(AxesParams("X_max"))
  cnt = AxesParams("X_min")
  For Each itm In SeriaXValues
    SeriaValues(cnt) = AxesParams("Y_min") - AxesParams("PlotAreaMargin")
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = SeriaXValues
  
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlCategory).Delete
'  objCht.Chart.Axes(xlValue).Delete
'  objCht.Chart.Axes(xlValue).MajorGridlines.Delete
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.Width = AxesParams("chartWidth") - AxesParams("PlotAreaMargin") + AxesParams("Y_ax_width")
  objCht.Chart.PlotArea.Height = AxesParams("chartHeight") - AxesParams("PlotAreaMargin")
  objCht.Chart.PlotArea.Left = 21 '+4 from Office chart
  objCht.Chart.PlotArea.Top = 21 '+4 from Office chart
  
  Set objCht = Nothing
End Sub

Sub CreateXAx(AxesParams As Scripting.Dictionary, SeriaXValues)
  
  Dim objCht As ChartObject
  Set objCht = Sheets(AxesParams("SheetName")).ChartObjects.Add( _
    AxesParams("topXPosition"), _
    AxesParams("topYposition"), _
    AxesParams("chartWidth"), _
    AxesParams("chartHeight") + AxesParams("X_ax_height") _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = AxesParams("Y_min")
  objCht.Chart.Axes(xlValue).MaximumScale = AxesParams("Y_max")
  
  Dim SeriaValues() As Double
  Dim itm As Variant
  Dim cnt As Long
  ReDim SeriaValues(AxesParams("X_max"))
  cnt = AxesParams("X_min")
  For Each itm In SeriaXValues
    SeriaValues(cnt) = AxesParams("Y_min") - AxesParams("PlotAreaMargin")
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = SeriaXValues
  
  objCht.Chart.Legend.Delete
  
'  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.Axes(xlValue).Delete
  objCht.Chart.Axes(xlValue).MajorGridlines.Delete

'  ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
'  ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
'  objCht.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
  
  objCht.Chart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.Width = AxesParams("chartWidth") - AxesParams("PlotAreaMargin")
  objCht.Chart.PlotArea.Height = AxesParams("chartHeight") - AxesParams("PlotAreaMargin") + AxesParams("X_ax_height")
  objCht.Chart.PlotArea.Left = 21 '+4 from Office chart
  objCht.Chart.PlotArea.Top = 21 '+4 from Office chart
  
  Set objCht = Nothing
End Sub

Sub CreateXAxis(plotSettings As Scripting.Dictionary, XIndex, SheetName As String)
  Dim objCht As ChartObject
  Dim axes_height, axes_width As Integer
  Dim x_axis_height, y_axis_width As Integer
  Dim axes_bind_position() As String
  
  axes_width = plotSettings("axes_width")
  axes_height = plotSettings("axes_height")
  x_axis_height = plotSettings("x_axis_height")
  y_axis_width = plotSettings("y_axis_width")
  axes_bind_position = Split(plotSettings("axes_bind_position"), ",")
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)), _
    axes_width, _
    axes_height _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = 0
  objCht.Chart.Axes(xlValue).MaximumScale = 1
  
  Dim SeriaValues() As Integer
  Dim itm As Variant
  Dim cnt As Long
  ReDim SeriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    SeriaValues(cnt) = -1
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = XIndex
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlValue).Delete
  objCht.Chart.Axes(xlValue).MajorGridlines.Delete
  
  objCht.Chart.Axes(xlCategory).TickMarkSpacing = 100
  objCht.Chart.Axes(xlCategory).TickLabelSpacing = 100
  objCht.Chart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
  'objCht.Chart.SetElement (msoElementPrimaryCategoryGridLinesMinor)
  objCht.Chart.PlotArea.Format.line.ForeColor.ObjectThemeColor = msoThemeColorText1
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.Width = axes_width - y_axis_width - 5
  objCht.Chart.PlotArea.Height = axes_height - 19
  objCht.Chart.PlotArea.Left = y_axis_width
  objCht.Chart.PlotArea.Top = 9
  
  'objCht.Chart.PlotArea.Format.line.Visible = msoFalse
  
  With objCht.Chart.Axes(xlCategory).MajorGridlines.Format.line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorBackground1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = -0.25
    .Transparency = 0
    .DashStyle = msoLineDash
  End With

End Sub

Sub CreateSubplotYAxis(plotSettings As Scripting.Dictionary, XIndex, SheetName As String, plotSettingsShift As Scripting.Dictionary)
  Dim objCht As ChartObject
  Dim subplot_axes_height, axes_width As Integer
  Dim x_axis_height, y_axis_width As Integer
  Dim axes_bind_position() As String
  
  axes_width = plotSettings("axes_width")
  subplot_axes_height = plotSettingsShift("subplot_axes_height")
  'x_axis_height = plotSettings("x_axis_height")
  y_axis_width = plotSettings("y_axis_width")
  axes_bind_position = Split(plotSettings("axes_bind_position"), ",")
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)) + plotSettingsShift("axes_bind_position_y"), _
    axes_width, _
    subplot_axes_height _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = 0
  objCht.Chart.Axes(xlValue).MaximumScale = 1
  
  Dim SeriaValues() As Integer
  Dim itm As Variant
  Dim cnt As Long
  ReDim SeriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    SeriaValues(cnt) = -1
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = XIndex
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.PlotArea.Format.line.ForeColor.ObjectThemeColor = msoThemeColorText1
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.Width = axes_width - y_axis_width - 3
'  objCht.Chart.PlotArea.Height = axes_height - 19
  objCht.Chart.PlotArea.Left = y_axis_width - 15
'  objCht.Chart.PlotArea.Top = 9
  
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).Fill.Visible = msoFalse
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  objCht.Chart.PlotArea.Format.Fill.Visible = msoFalse
  objCht.Chart.PlotArea.Format.line.Visible = msoFalse
 
  With objCht.Chart.Axes(xlValue).MajorGridlines.Format.line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorBackground1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = -0.25
    .Transparency = 0
    .DashStyle = msoLineDash
  End With
  
End Sub

Sub CreateAxes(plotSettings As Scripting.Dictionary, XIndex, SheetName As String)
    Dim subplotHeightShare() As String
    Dim heightShare As Variant
    Dim height100 As Integer
    Dim subplot_height_share() As String
    Dim plotSettingsShift As New Scripting.Dictionary
    Dim axes_bind_position_y As Integer
    Dim axes_height, subplot_axes_height, axes_width As Integer
    
    Call CreateXAxis(plotSettings, XIndex, SheetName)
    
    subplot_height_share = Split(plotSettings("subplot_height_share"), ",")
    height100 = 0
    For Each heightShare In subplot_height_share
      height100 = height100 + CInt(heightShare)
    Next heightShare
    
    axes_width = plotSettings("axes_width")
    axes_height = plotSettings("axes_height") - plotSettings("x_axis_height") - 6
    axes_bind_position_y = 6
    For Each heightShare In subplot_height_share
      Set plotSettingsShift = Nothing
      plotSettingsShift.Add "axes_bind_position_y", axes_bind_position_y
      subplot_axes_height = CInt(axes_height * heightShare / height100)
      plotSettingsShift.Add "subplot_axes_height", subplot_axes_height
      
      Call CreateSubplotYAxis(plotSettings, XIndex, SheetName, plotSettingsShift)
      axes_bind_position_y = axes_bind_position_y + subplot_axes_height
    Next heightShare
End Sub

Function GetExtr(seria() As Variant) As Variant()
  Dim extr() As Variant
  ReDim extr(1)
  Dim itm As Variant
  Dim mx, mn As Variant
  
  For Each itm In seria
      mx = itm
      mn = itm
      If itm <> Empty Then
        Exit For
      End If
  Next itm
  For Each itm In seria
    If itm <> Empty Then
      If mx < itm Then
        mx = itm
      End If
      If mn > itm Then
        mn = itm
      End If
    End If
  Next itm
  extr(0) = mn
  extr(1) = mx
  GetExtr = extr
End Function

Sub GetExtrem(DS As DataSet, plotSettings As Scripting.Dictionary)
  Dim series_columns() As String
  Dim seria() As Variant
  Dim col As Variant
  Dim extr() As Variant
  Dim plotSettingsCopy As New Scripting.Dictionary
  Dim CM As New CommonMethods
  Dim SubPlotNumber As String
  Dim mx, mn As Variant
  
  
  Call CM.CopyDict(plotSettings, plotSettingsCopy)
  
  series_columns = Split(plotSettings("series_columns"), ",")
  For Each col In series_columns
    SubPlotNumber = plotSettingsCopy("series_subplot")(col)
    If plotSettings("subplot_y_min")(SubPlotNumber) = "auto" Or plotSettings("subplot_y_max")(SubPlotNumber) = "auto" Then
      seria = DS.GetCol(col)
      extr() = GetExtr(seria)
      If plotSettings("subplot_y_min")(SubPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_min")(SubPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_min")(SubPlotNumber) > mn Then
          plotSettingsCopy("subplot_y_min")(SubPlotNumber) = extr(0)
        End If
      End If
      If plotSettings("subplot_y_max")(SubPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_max")(SubPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_max")(SubPlotNumber) < mx Then
          plotSettingsCopy("subplot_y_max")(SubPlotNumber) = extr(1)
        End If
      End If
    End If
    
  Next col
  
  Call CM.CopyDict(plotSettingsCopy, plotSettings)
  
  Debug.Print plotSettings("subplot_y_min")("1"), plotSettingsCopy("subplot_y_max")("1")
  Debug.Print plotSettings("subplot_y_min")("2"), plotSettingsCopy("subplot_y_max")("2")

End Sub

Sub PlotChart(DS As DataSet, plotSettings As Scripting.Dictionary, SheetName As String)
  Dim seriesColumns() As String
  Dim sCol As Variant
  
  'Dim AxesParams As New Scripting.Dictionary
  
  Call DeleteCharts
  Call GetExtrem(DS, plotSettings)
  
  Call CreateAxes(plotSettings, DS.GetRowIndex(), SheetName)
  
  
  
  
  
  
  
  
'  AxesParams.Add "SheetName", "PM"
'  AxesParams.Add "topXPosition", 100
'  AxesParams.Add "topYposition", 20
'  AxesParams.Add "chartWidth", 500
'  AxesParams.Add "chartHeight", 300
'  AxesParams.Add "Y_min", 63800
'  AxesParams.Add "Y_max", 64200
'  AxesParams.Add "X_min", 0
'  AxesParams.Add "X_max", DS.rowsCount - 1
'  AxesParams.Add "Y_ax_width", 40
'  AxesParams.Add "X_ax_height", 40
'  AxesParams.Add "PlotAreaMargin", 50
  
  
  'Call CreateYAx(AxesParams, DS.GetRowIndex())
'  Call CreateXAx(AxesParams, DS.GetRowIndex())
'  Call CreateLineChart(AxesParams, DS.GetCol("<LAST_BASE>"), DS.GetRowIndex())
'  Call CreateLineChart(AxesParams, DS.GetCol("<LAST_FIRST>"), DS.GetRowIndex())
'  Call CreateLineChart(AxesParams, DS.GetCol("<LAST_SECOND>"), DS.GetRowIndex())
  
'  seriesColumns = Split(plotSettings("series_columns"), ",")
'  For Each sCol In seriesColumns
'
'    Debug.Print sCol
'  Next sCol

End Sub
