Attribute VB_Name = "ChartArea"
Sub ChartArea_Properties()
    Set Chart = ActiveSheet.ChartObjects(1).Chart
    Set ChartArea = Chart.ChartArea
    Dim Info As String
    Info = vbNullString
    Info = Info & "ChartArea.Name = " + ChartArea.Name & vbNewLine
    Info = Info & "ChartArea.Position = " + CStr(ChartArea.Position) & vbNewLine
    Info = Info & vbNewLine
    
    Info = Info & "ChartArea.Left = " + CStr(ChartArea.Left) & vbNewLine
    Info = Info & "ChartArea.Top = " + CStr(ChartArea.Top) & vbNewLine
    Info = Info & "ChartArea.Width = " + CStr(ChartArea.Width) & vbNewLine
    Info = Info & "ChartArea.Height = " + CStr(ChartArea.Height) & vbNewLine
    Info = Info & vbNewLine
    MsgBox (Info)
End Sub




