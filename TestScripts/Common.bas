Attribute VB_Name = "Common"
Sub ShowBorder()
    Set Chart = ActiveSheet.ChartObjects(1).Chart
    If ActiveSheet.cbShowBorders.Value = True Then
        Chart.ChartArea.Border.LineStyle = xlDash
        Chart.PLotArea.Border.LineStyle = xlDot
    Else
        Chart.ChartArea.Border.LineStyle = xlNone
        Chart.PLotArea.Border.LineStyle = xlNone
    End If
End Sub
