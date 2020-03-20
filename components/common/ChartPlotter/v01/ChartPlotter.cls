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

#If VBA7 Then
  'For 64 Bit Systems
'  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
'  Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#Else
  'For 32 Bit Systems
'  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#End If
'Const SM_CXSCREEN = 0
'Const SM_CYSCREEN = 1

Private exportYShift As New Scripting.Dictionary
Private exportChartCnt As Integer

Sub DeleteCharts(SheetName As String)
  Dim chtObj As ChartObject
  For Each chtObj In Sheets(SheetName).ChartObjects
    chtObj.Delete
  Next
  Set exportYShift = Nothing
  exportChartCnt = 1
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
  
  exportYShift.Add exportChartCnt, 0
  exportChartCnt = exportChartCnt + 1
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)), _
    axes_width, _
    axes_height _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.Axes(xlValue).MinimumScale = 0
  objCht.Chart.Axes(xlValue).MaximumScale = 1
  
  Dim seriaValues() As Integer
  Dim itm As Variant
  Dim cnt As Long
  ReDim seriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    seriaValues(cnt) = -1
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = seriaValues
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

Sub CreateSubplotLineChartLayer(plotSettings As Scripting.Dictionary, XIndex, SheetName As String, plotSettingsShift As Scripting.Dictionary)
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
  
  exportYShift.Add exportChartCnt, plotSettingsShift("axes_bind_position_y")
  exportChartCnt = exportChartCnt + 1
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)) + plotSettingsShift("axes_bind_position_y"), _
    axes_width, _
    subplot_axes_height _
  )
  objCht.Chart.ChartType = xlLine
  objCht.Chart.DisplayBlanksAs = xlInterpolated
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  
  Dim seriaValues() As Variant
  Dim itm As Variant
  Dim cnt As Long
  ReDim seriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    seriaValues(cnt) = plotSettings("subplot_y_min")(subPlotNumber) - plotSettings("subplot_y_max")(subPlotNumber)
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = seriaValues
  objCht.Chart.SeriesCollection(1).XValues = XIndex
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.Axes(xlValue).Delete
  objCht.Chart.Axes(xlValue).MajorGridlines.Delete
  objCht.Chart.Axes(xlValue).MinimumScale = plotSettings("subplot_y_min")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MaximumScale = plotSettings("subplot_y_max")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MajorUnit = plotSettings("subplot_y_step")(subPlotNumber)
  
  objCht.Chart.PlotArea.Format.line.ForeColor.ObjectThemeColor = msoThemeColorText1
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.InsideWidth = axes_width - y_axis_width * 2
  objCht.Chart.PlotArea.InsideHeight = subplot_axes_height - 17
  objCht.Chart.PlotArea.InsideLeft = y_axis_width - 5
  objCht.Chart.PlotArea.InsideTop = 5
  
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).Fill.Visible = msoFalse
  'Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  objCht.Chart.PlotArea.Format.Fill.Visible = msoFalse
  objCht.Chart.PlotArea.Format.line.Visible = msoFalse
  
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).Name = "Line" & subPlotNumber
End Sub

Sub CreateSubplotHistChartLayer(plotSettings As Scripting.Dictionary, XIndex, SheetName As String, plotSettingsShift As Scripting.Dictionary)
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
  
  exportYShift.Add exportChartCnt, plotSettingsShift("axes_bind_position_y")
  exportChartCnt = exportChartCnt + 1
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)) + plotSettingsShift("axes_bind_position_y"), _
    axes_width, _
    subplot_axes_height _
  )
  objCht.Chart.ChartType = xlColumnClustered
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  
  Dim seriaValues() As Variant
  Dim itm As Variant
  Dim cnt As Long
  ReDim seriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    seriaValues(cnt) = plotSettings("subplot_y_min")(subPlotNumber) - plotSettings("subplot_y_max")(subPlotNumber)
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = seriaValues
  objCht.Chart.SeriesCollection(1).XValues = XIndex
  objCht.Chart.Legend.Delete
  
  objCht.Chart.Axes(xlCategory).Delete
  objCht.Chart.Axes(xlValue).Delete
  objCht.Chart.Axes(xlValue).MajorGridlines.Delete
  objCht.Chart.Axes(xlValue).MinimumScale = plotSettings("subplot_y_min")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MaximumScale = plotSettings("subplot_y_max")(subPlotNumber)
  objCht.Chart.Axes(xlValue).MajorUnit = plotSettings("subplot_y_step")(subPlotNumber)
  
  objCht.Chart.PlotArea.Format.line.ForeColor.ObjectThemeColor = msoThemeColorText1
  
  Dim temp As Integer 'Office bug
  temp = objCht.Chart.PlotArea.Width 'Office bug
  
  objCht.Chart.PlotArea.InsideWidth = axes_width - y_axis_width * 2
  objCht.Chart.PlotArea.InsideHeight = subplot_axes_height - 17
  objCht.Chart.PlotArea.InsideLeft = y_axis_width - 5
  objCht.Chart.PlotArea.InsideTop = 5
  
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).Fill.Visible = msoFalse
  'Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  objCht.Chart.PlotArea.Format.Fill.Visible = msoFalse
  objCht.Chart.PlotArea.Format.line.Visible = msoFalse
  
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).Name = "Hist" & subPlotNumber
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
  
  exportYShift.Add exportChartCnt, plotSettingsShift("axes_bind_position_y")
  exportChartCnt = exportChartCnt + 1
  
  Set objCht = Sheets(SheetName).ChartObjects.Add( _
    CInt(axes_bind_position(0)), _
    CInt(axes_bind_position(1)) + plotSettingsShift("axes_bind_position_y"), _
    axes_width, _
    subplot_axes_height _
  )
  objCht.Chart.ChartType = xlLine
  Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
  
  Dim seriaValues() As Variant
  Dim itm As Variant
  Dim cnt As Long
  ReDim seriaValues(UBound(XIndex) - LBound(XIndex))
  cnt = 0
  For Each itm In XIndex
    seriaValues(cnt) = plotSettings("subplot_y_min")(subPlotNumber) - plotSettings("subplot_y_max")(subPlotNumber)
    cnt = cnt + 1
  Next itm
  
  objCht.Chart.SeriesCollection.NewSeries
  objCht.Chart.SeriesCollection(1).Values = seriaValues
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
  'Sheets(SheetName).Shapes(Mid(objCht.Chart.Name, Len(SheetName) + 2)).line.Visible = msoFalse
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
      Call CreateSubplotHistChartLayer(plotSettings, XIndex, SheetName, plotSettingsShift)
      Call CreateSubplotLineChartLayer(plotSettings, XIndex, SheetName, plotSettingsShift)
      
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

Sub SetPlotSettingsExtrem(DS As DataSet, plotSettings As Scripting.Dictionary)
  Dim series_columns() As String
  Dim seria() As Variant
  Dim col As Variant
  Dim extr() As Double
  Dim plotSettingsCopy As New Scripting.Dictionary
  Dim CM As New CommonMethods
  Dim subPlotNumber As String
  'Dim Settings_Reader As New SettingsReader
  Dim heightShare As Variant
  Dim subPlotCnt As Integer
  Dim subplotYMin, subplotYMax, subplotYStep As Double
  
  Call CM.CopyDict(plotSettings, plotSettingsCopy)
  
  series_columns = Split(plotSettings("series_columns"), ",")
  For Each col In series_columns
    subPlotNumber = plotSettingsCopy("series_subplot")(col)
    If plotSettings("subplot_y_min")(subPlotNumber) = "auto" Or plotSettings("subplot_y_max")(subPlotNumber) = "auto" Then
      seria = DS.GetCol(col)
      extr() = GetExtr(seria)
      If plotSettings("subplot_y_min")(subPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_min")(subPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_min")(subPlotNumber) > extr(0) Then
          plotSettingsCopy("subplot_y_min")(subPlotNumber) = extr(0)
        End If
      Else
        plotSettingsCopy("subplot_y_min")(subPlotNumber) = CDbl(plotSettingsCopy("subplot_y_min")(subPlotNumber))
      End If
      If plotSettings("subplot_y_max")(subPlotNumber) = "auto" Then
        If plotSettingsCopy("subplot_y_max")(subPlotNumber) = "auto" Or plotSettingsCopy("subplot_y_max")(subPlotNumber) < extr(1) Then
          plotSettingsCopy("subplot_y_max")(subPlotNumber) = extr(1)
        End If
      Else
        plotSettingsCopy("subplot_y_max")(subPlotNumber) = CDbl(plotSettingsCopy("subplot_y_max")(subPlotNumber))
      End If
    End If
  Next col
  
  subPlotCnt = 1
  For Each heightShare In Split(plotSettings("subplot_height_share"), ",")
    If plotSettings("subplot_y_step")(CStr(subPlotCnt)) = "auto" Then
      plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)) = _
        (plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt)) - _
        plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))) / 5
    Else
      plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)) = CDbl(plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt)))
    End If
    
    If plotSettings("subplot_y_min")(CStr(subPlotCnt)) = "auto" Then
      subplotYMin = plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))
      subplotYMax = plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt))
      subplotYStep = plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt))
      plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt)) = subplotYMin - subplotYMin Mod subplotYStep
    End If
    
    If plotSettings("subplot_y_max")(CStr(subPlotCnt)) = "auto" Then
      subplotYMin = plotSettingsCopy("subplot_y_min")(CStr(subPlotCnt))
      subplotYMax = plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt))
      subplotYStep = plotSettingsCopy("subplot_y_step")(CStr(subPlotCnt))
      plotSettingsCopy("subplot_y_max")(CStr(subPlotCnt)) = subplotYMax - subplotYMax Mod subplotYStep + subplotYStep
    End If
    subPlotCnt = subPlotCnt + 1
  Next heightShare
  
  Call CM.CopyDict(plotSettingsCopy, plotSettings)
  
'  Debug.Print plotSettings("subplot_y_min")("1"), plotSettings("subplot_y_max")("1")
'  Debug.Print plotSettings("subplot_y_min")("2"), plotSettings("subplot_y_max")("2")
'  Debug.Print plotSettings("subplot_y_step")("1")
'  Debug.Print plotSettings("subplot_y_step")("2")
End Sub

Sub AddCharts(DS As DataSet, plotSettings As Scripting.Dictionary, SheetName As String)
  Dim objCht As ChartObject
  Dim series_columns() As String
  Dim seriaValues() As Variant
  Dim col As Variant
  Dim subPlotNumber As String
  Dim seriaType As String
  Dim seriaColor() As String
  Dim seriaWeight As Double
  Dim lineSeriesCnt, histSeriesCnt As Integer
  
  series_columns = Split(plotSettings("series_columns"), ",")
  lineSeriesCnt = 1
  histSeriesCnt = 1
  For Each col In series_columns
    subPlotNumber = plotSettings("series_subplot")(col)
    seriaValues = DS.GetCol(col)
    
    seriaType = plotSettings("series_type")(col)
    seriaWeight = CDbl(Replace(plotSettings("series_weight")(col), ".", Application.DecimalSeparator))
    seriaColor = Split(plotSettings("series_color")(col), ",")
    If seriaType = "line" Then
      Set objCht = Sheets(SheetName).ChartObjects("Line" & subPlotNumber)
      objCht.Chart.SeriesCollection.NewSeries
      lineSeriesCnt = lineSeriesCnt + 1
      objCht.Chart.SeriesCollection(lineSeriesCnt).Values = seriaValues
      With objCht.Chart.SeriesCollection(lineSeriesCnt)
        .Format.line.Visible = msoTrue
        .Format.line.Weight = seriaWeight
        .Format.line.ForeColor.RGB = RGB(seriaColor(0), seriaColor(1), seriaColor(2))
      End With
    End If
    If seriaType = "hist" Then
      Set objCht = Sheets(SheetName).ChartObjects("Hist" & subPlotNumber)
      objCht.Chart.SeriesCollection.NewSeries
      histSeriesCnt = histSeriesCnt + 1
      objCht.Chart.SeriesCollection(histSeriesCnt).Values = seriaValues
      With objCht.Chart.SeriesCollection(histSeriesCnt)
        .Format.line.Visible = msoTrue
        .Format.line.Weight = seriaWeight
        .Format.Fill.ForeColor.RGB = RGB(seriaColor(0), seriaColor(1), seriaColor(2))
      End With
    End If
  Next col
End Sub

Sub PlotChart(DS As DataSet, ByVal plotSettings As Scripting.Dictionary, SheetName As String)
  
  Call DeleteCharts(SheetName)
  Call SetPlotSettingsExtrem(DS, plotSettings)
  Call CreateAxes(plotSettings, DS.GetRowIndex(), SheetName)
  Call AddCharts(DS, plotSettings, SheetName)
  
'  ThisWorkbook.Worksheets(SheetName).Shapes.SelectAll
'  Selection.ShapeRange.Group
End Sub

Sub ExportCharts(compositeFileName As String, exportFileFolder)
  Dim RW_File As New RWFile
  Dim chtObj As ChartObject
  Dim chtCnt As Integer
  Dim exportFileName, firstFileName As String
  Dim exportFilePath As String
  Dim chartYPosition As Integer
  
  Call RW_File.OpenFile(exportFileFolder & "composit.bat", "a")
  'RW_File.WriteLine ("@echo off")
  
  chtCnt = 1
  For Each chtObj In ActiveSheet.ChartObjects
    exportFileName = compositeFileName & "_" & Format(chtCnt, "000")
    exportFilePath = exportFileFolder & exportFileName & ".png"
    chtObj.Chart.Export exportFilePath
    If chtCnt = 1 Then
      firstFileName = exportFileName
    ElseIf chtCnt = 2 Then
      chartYPosition = Application.WorksheetFunction.Round(exportYShift(chtCnt) / 72 * 120, 0)
      Call RW_File.WriteLine("magick convert " & firstFileName & ".png " & exportFileName & ".png -geometry +1+" & chartYPosition & " -type TrueColor -composite " & compositeFileName & ".png")
      Call RW_File.WriteLine("del " & firstFileName & ".png ")
      Call RW_File.WriteLine("del " & exportFileName & ".png")
    Else
      chartYPosition = Application.WorksheetFunction.Round(exportYShift(chtCnt) / 72 * 120, 0)
      Call RW_File.WriteLine("magick convert " & compositeFileName & ".png " & exportFileName & ".png -geometry +1+" & chartYPosition & " -type TrueColor -composite " & compositeFileName & ".png")
      Call RW_File.WriteLine("del " & exportFileName & ".png")
    End If
    chtCnt = chtCnt + 1
  Next
  
  'RW_File.WriteLine ("pause")
  Call RW_File.CloseFile
   
End Sub
