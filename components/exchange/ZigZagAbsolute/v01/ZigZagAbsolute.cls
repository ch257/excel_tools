VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZigZagAbsolute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public isMax, isMin As Boolean
Public maxVal, minVal As Variant
Public maxIndex, minIndex As Long

Private minMovingSize
Private lastHighDirection
Private lastLowDirection As String
Private highBuffer As New Buffer
Private lowBuffer As New Buffer
Private highIndexBuffer As New Buffer
Private lowIndexBuffer As New Buffer
Private isExpMax, isExpMin As Variant
Private expMax, expMin  As Variant
Private expMaxIndex, expMinIndex As Long
Private lastFoundIsMax As Boolean
Private lastFoundIsMin As Boolean
Private lastFoundMax, lastFoundMin As Variant
Private lastFoundMaxIndex, lastFoundMinIndex As Long
Private confirmedMax, confirmedMin As Boolean

Sub Init(min_Moving_Size As Variant)
    minMovingSize = min_Moving_Size Or 0
    lastHighDirection = "none"
    lastLowDirection = "none"
    Call highBuffer.Create(2)
    Call lowBuffer.Create(2)
    Call highIndexBuffer.Create(2)
    Call lowIndexBuffer.Create(2)
    lastFoundIsMax = False
    lastFoundIsMin = False
    confirmedMax = False
    confirmedMin = False
End Sub

Private Function isExpectedMax() As Boolean
    If lastHighDirection = "up" Then
        If highBuffer.GetItem(0) > highBuffer.GetItem(1) Then
            lastHighDirection = "down"
            isExpectedMax = True
            Exit Function
        End If
    ElseIf lastHighDirection = "up" Then
        If highBuffer.GetItem(0) = highBuffer.GetItem(1) Then
            lastHighDirection = "up"
        End If
    ElseIf lastHighDirection = "down" Then
        If highBuffer.GetItem(0) < highBuffer.GetItem(1) Then
            lastHighDirection = "up"
        End If
    ElseIf lastHighDirection = "down" Then
        If highBuffer.GetItem(0) = highBuffer.GetItem(1) Then
            lastHighDirection = "down"
        End If
    Else
        If highBuffer.GetItem(0) < highBuffer.GetItem(1) Then
            lastHighDirection = "up"
        ElseIf highBuffer.GetItem(0) > highBuffer.GetItem(1) Then
            lastHighDirection = "down"
        End If
    End If
    isExpectedMax = False
End Function

Private Function isExpectedMin() As Boolean
    If lastLowDirection = "down" Then
        If lowBuffer.GetItem(0) < lowBuffer.GetItem(1) Then
            lastLowDirection = "up"
            isExpectedMin = True
            Exit Function
        End If
    ElseIf lastLowDirection = "down" Then
        If lowBuffer.GetItem(0) = lowBuffer.GetItem(1) Then
            lastLowDirection = "down"
        End If
    ElseIf lastLowDirection = "up" Then
        If lowBuffer.GetItem(0) > lowBuffer.GetItem(1) Then
            lastLowDirection = "down"
        End If
    ElseIf lastLowDirection = "up" Then
        If lowBuffer.GetItem(0) = lowBuffer.GetItem(1) Then
            lastLowDirection = "up"
        End If
    Else
        If lowBuffer.GetItem(0) < lowBuffer.GetItem(1) Then
            lastLowDirection = "up"
        ElseIf lowBuffer.GetItem(0) > lowBuffer.GetItem(1) Then
            lastLowDirection = "down"
        End If
    End If
    isExpectedMin = False
End Function

Sub findExpectedExtrems(bar_cnt, high_value, low_value)
    isExpMax = False
    isExpMin = False
    '------------------ BEGIN find expected max
    If high_value Then
        If Not highBuffer.isReady Then
            Call highBuffer.Slide(high_value)
            Call highIndexBuffer.Slide(bar_cnt)
        Else
            Call highBuffer.Slide(high_value)
            Call highIndexBuffer.Slide(bar_cnt)
            If isExpectedMax() Then
                isExpMax = True
                expMax = highBuffer.GetItem(0)
                expMaxIndex = highIndexBuffer.GetItem(0)
            End If
        End If
    End If
    '------------------ END find expected max
    '------------------ BEGIN find expected min
    If low_value Then
        If Not lowBuffer.isReady Then
            Call lowBuffer.Slide(low_value)
            Call lowIndexBuffer.Slide(bar_cnt)
        Else
            Call lowBuffer.Slide(low_value)
            Call lowIndexBuffer.Slide(bar_cnt)
            If isExpectedMin() Then
                isExpMin = True
                expMin = lowBuffer.GetItem(0)
                expMinIndex = lowIndexBuffer.GetItem(0)
            End If
        End If
    End If
    '------------------ END find expected min
End Sub

Sub confirmExpectedExtrems()
    isMax = False
    isMin = False
    If lastFoundIsMax And isExpMax Then
        If lastFoundMax < expMax Then
            If Not confirmedMax Then
                Call confirmExpMax
            End If
        Else
            If Not confirmedMax Then
                Call confirmLastFoundMax
            End If
        End If
    Else
        If lastFoundIsMax Then
            If Not confirmedMax Then
                Call confirmLastFoundMax
            End If
        End If
    End If
    If isExpMax Then
        If Not confirmedMax Then
            Call confirmExpMax
        End If
        If Not confirmedMax Then
            Call setLastFoundMax
        End If
    End If

    '-----------------------------------------------
    If lastFoundIsMin And isExpMin Then
        If lastFoundMin > expMin Then
            If Not confirmedMin Then
                Call confirmExpMin
            End If
        Else
            If Not confirmedMin Then
                Call confirmLastFoundMin
            End If
        End If
    Else
        If lastFoundIsMin Then
            If Not confirmedMin Then
                Call confirmLastFoundMin
            End If
        End If
    End If
    If isExpMin Then
        If Not confirmedMin Then
            Call confirmExpMin
        End If
        If Not confirmedMin Then
            Call setLastFoundMin
        End If
    End If
End Sub

Private Sub confirmExpMax()
    If expMax - lowBuffer.GetItem(1) >= minMovingSize Then
        isMax = True
        maxVal = expMax
        maxIndex = expMaxIndex
        confirmedMax = True
        confirmedMin = False
        lastFoundIsMax = False
    End If
End Sub

Private Sub confirmExpMin()
    If highBuffer.GetItem(1) - expMin >= minMovingSize Then
        isMin = True
        minVal = expMin
        minIndex = expMinIndex
        confirmedMin = True
        confirmedMax = False
        lastFoundIsMin = False
    End If
End Sub

Private Sub confirmLastFoundMax()
    If lastFoundMax - lowBuffer.GetItem(1) >= minMovingSize Then
        isMax = True
        maxVal = lastFoundMax
        maxIndex = lastFoundMaxIndex
        confirmedMax = True
        confirmedMin = False
        lastFoundIsMax = False
    End If
End Sub

Private Sub confirmLastFoundMin()
    If highBuffer.GetItem(1) - lastFoundMin >= minMovingSize Then
        isMin = True
        minVal = lastFoundMin
        minIndex = lastFoundMinIndex
        confirmedMin = True
        confirmedMax = False
        lastFoundIsMin = False
    End If
End Sub

Private Sub setLastFoundMax()
    If lastFoundIsMax Then
        If lastFoundMax <= expMax Then
            lastFoundMax = expMax
            lastFoundMaxIndex = expMaxIndex
        End If
    Else
        lastFoundMax = expMax
        lastFoundMaxIndex = expMaxIndex
    End If
    lastFoundIsMax = True
End Sub

Private Sub setLastFoundMin()
    If lastFoundIsMin Then
        If lastFoundMin >= expMin Then
            lastFoundMin = expMin
            lastFoundMinIndex = expMinIndex
        End If
    Else
        lastFoundMin = expMin
        lastFoundMinIndex = expMinIndex
    End If
    lastFoundIsMin = True
End Sub


