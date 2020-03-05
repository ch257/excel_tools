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
  
  objCht.Chart.PlotArea.InsideWidth = axes_width - y_axis_width * 2
  objCht.Chart.PlotArea.InsideHeight = axes_height - x_axis_height - 5
  objCht.Chart.PlotArea.InsideLeft = y_axis_width - 5
  objCht.Chart.PlotArea.InsideTop = 5
  
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
  Dim subPlotNumber As String
  
  axes_width = plotSettings("axes_width")
  subplot_axes_height = plotSettingsShift("subplot_axes_height")
  subPlotNumber = plotSettingsShift("subplot_number")
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
  
  Dim SeriaValues() As Variant
  Dim itm As Variant
  Dim cnt As Long
  ReDim SeriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    SeriaValues(cnt) = plotSettings("subplot_y_min")(subPlotNumber) - plotSettings("subplot_y_max")(subPlotNumber)
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = SeriaValues
  objCht.Chart.SeriesCollection(1).XValues = XIndex
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.PlotArea.Format.line.ForeColor.ObjectThemeColor = msoThemeColorText1
  
  objCht.Chart.Axes(xlValue).MinimumScale = plotSettings("subplot_y_min")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MaximumScale = plotSettings("subplot_y_max")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MajorUnit = plotSettings("subplot_y_step")(subPlotNumber)
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.InsideWidth = axes_width - y_axis_width * 2
  objCht.Chart.PlotArea.InsideHeight = subplot_axes_height - 10
  objCht.Chart.PlotArea.InsideLeft = y_axis_width - 5
  objCht.Chart.PlotArea.InsideTop = 5
  
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
    Dim subPlotNumber As String
    Call CreateXAxis(plotSettings, XIndex, SheetName)
    
    subplot_height_share = Split(plotSettings("subplot_height_share"), ",")
    height100 = 0
    For Each heightShare In subplot_height_share
      height100 = height100 + CInt(heightShare)
    Next heightShare
    
    axes_width = plotSettings("axes_width")
    axes_height = plotSettings("axes_height") - plotSettings("x_axis_height")
    axes_bind_position_y = 5
    subPlotNumber = 0
    For Each heightShare In subplot_height_share
      subPlotNumber = subPlotNumber + 1
      Set plotSettingsShift = Nothing
      plotSettingsShift.Add "axes_bind_position_y", axes_bind_position_y
      subplot_axes_height = CInt(axes_height * heightShare / height100)
      plotSettingsShift.Add "subplot_axes_height", subplot_axes_height
      plotSettingsShift.Add "subplot_number", subPlotNumber
      
      Call CreateSubplotYAxis(plotSettings, XIndex, SheetName, plotSettingsShift)
      axes_bind_position_y = axes_bind_position_y + subplot_axes_height
    Next heightShare
End Sub

Function GetExtr(seria() As Variant) As Double()
  Dim extr() As Double
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
  extr(0) = CDbl(mn)
  extr(1) = CDbl(mx)
  GetExtr = extr
End Function

Sub SetPlotSettingsExtrem(DS As DataSet, DS_Settings As Scripting.Dictionary)
  Dim series_columns() As String
  Dim seria() As Variant
  Dim col As Variant
  Dim extr() As Double
  Dim plotSettingsCopy As New Scripting.Dictionary
  Dim CM As New CommonMethods
  Dim subPlotNumber As String
  Dim Settings_Reader As New SettingsReader
  Dim heightShare As Variant
  Dim subPlotCnt As Integer
  Dim subplotYMin, subplotYMax, subplotYStep As Double
  
  Call CM.CopyDict(DS_Settings("plot_settings"), plotSettingsCopy)
  
  series_columns = Split(DS_Settings("plot_settings")("series_columns"), ",")
  For Each col In series_columns
    subPlotNumber = plotSettingsCopy("series_subplot")(col)
    If DS_Settings("plot_settings")("subplot_y_min")(subPlotNumber) = "auto" Or DS_Settings("plot_settings")("subplot_y_max")(subPlotNumber) = "auto" Then
      seria = DS.GetCol(col)
      extr() = GetExtr(seria)
      If DS_Settings("plot_settings")("subplot_y_min")(subPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_min")(subPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_min")(subPlotNumber) > extr(0) Then
          plotSettingsCopy("subplot_y_min")(subPlotNumber) = extr(0)
        End If
      Else
        plotSettingsCopy("subplot_y_min")(subPlotNumber) = CDbl(plotSettingsCopy("subplot_y_min")(subPlotNumber))
      End If
      If DS_Settings("plot_settings")("subplot_y_max")(subPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_max")(subPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_max")(subPlotNumber) < extr(1) Then
          plotSettingsCopy("subplot_y_max")(subPlotNumber) = extr(1)
        End If
      Else
        plotSettingsCopy("subplot_y_max")(subPlotNumber) = CDbl(plotSettingsCopy("subplot_y_max")(subPlotNumber))
      End If
    End If
  Next col
  
  subPlotCnt = 1
  For Each heightShare In Split(DS_Settings("plot_settings")("subplot_height_share"), ",")
    If DS_Settings("plot_settings")("subplot_y_step")(CStr(subPlotCnt)) = "auto" Then
      plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)) = _
        (plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt)) - _
        plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))) / 5
    Else
      plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)) = CDbl(plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)))
    End If
    
    If DS_Settings("plot_settings")("subplot_y_min")(CStr(subPlotCnt)) = "auto" Then
      subplotYMin = plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))
      subplotYMax = plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt))
      subplotYStep = plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt))
      plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt)) = subplotYMin - subplotYMin Mod subplotYStep
    End If
    
    If DS_Settings("plot_settings")("subplot_y_max")(CStr(subPlotCnt)) = "auto" Then
      subplotYMin = plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))
      subplotYMax = plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt))
      subplotYStep = plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt))
      plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt)) = subplotYMax - subplotYMax Mod subplotYStep + subplotYStep
    End If
    subPlotCnt = subPlotCnt + 1
  Next heightShare
  
  Call CM.CopyDict(plotSettingsCopy, DS_Settings("plot_settings"))
  
'  Debug.Print DS_Settings("plot_settings")("subplot_y_min")("1"), DS_Settings("plot_settings")("subplot_y_max")("1")
'  Debug.Print DS_Settings("plot_settings")("subplot_y_min")("2"), DS_Settings("plot_settings")("subplot_y_max")("2")
'  Debug.Print DS_Settings("plot_settings")("subplot_y_step")("1")
'  Debug.Print DS_Settings("plot_settings")("subplot_y_step")("2")
End Sub

Sub PlotChart(DS As DataSet, ByVal DS_Settings As Scripting.Dictionary, SheetName As String)
  
  Call DeleteCharts
  Call SetPlotSettingsExtrem(DS, DS_Settings)
  
  Call CreateAxes(DS_Settings("plot_settings"), DS.GetRowIndex(), SheetName)
  

End Sub
