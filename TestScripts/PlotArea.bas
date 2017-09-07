Attribute VB_Name = "PLotArea"
Sub PlotArea_Properties()
    Set Chart = ActiveSheet.ChartObjects(1).Chart
    Set Plot = Chart.PLotArea
    Dim Info As String
    Info = vbNullString
    Info = Info & "PlotArea.Name = " + Plot.Name & vbNewLine
    Info = Info & "PlotArea.Position = " + CStr(Plot.Position) & vbNewLine
    Info = Info & vbNewLine
    
    Info = Info & "PlotArea.Left = " + CStr(Plot.Left) & vbNewLine
    Info = Info & "PlotArea.Top = " + CStr(Plot.Top) & vbNewLine
    Info = Info & "PlotArea.Width = " + CStr(Plot.Width) & vbNewLine
    Info = Info & "PlotArea.Height = " + CStr(Plot.Height) & vbNewLine
    Info = Info & vbNewLine
    
    Info = Info & "PlotArea.InsideLeft = " + CStr(Plot.InsideLeft) & vbNewLine
    Info = Info & "PlotArea.InsideTop = " + CStr(Plot.InsideTop) & vbNewLine
    Info = Info & "PlotArea.InsideWidth = " + CStr(Plot.InsideWidth) & vbNewLine
    Info = Info & "PlotArea.InsideHeight = " + CStr(Plot.InsideHeight) & vbNewLine
    Info = Info & vbNewLine
    MsgBox (Info)
End Sub



