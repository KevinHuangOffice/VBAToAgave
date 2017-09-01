Attribute VB_Name = "Point"
Sub Point_Test_Reset()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Call Point_Explosion_Reset
End Sub

Rem Point: Display Left,Right,Width,Height properties
Sub Point_Rect_Get()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    Dim Info As String
    Info = vbNullString
    For i = 1 To Points.Count
        Info = Info & "Point(" + CStr(i) + ").Left = " + CStr(Points(i).Left) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").Top = " + CStr(Points(i).Top) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").Width = " + CStr(Points(i).Width) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").Height = " + CStr(Points(i).Height) & vbNewLine
        Info = Info & vbNewLine
    Next i
    MsgBox (Info)
End Sub

Rem Point: Display properties
Sub Point_Properties_Get()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    Dim Info As String
    Info = vbNullString
    For i = 1 To Points.Count
        Info = Info & "Point(" + CStr(i) + ").Explosion = " + CStr(Points(i).Explosion) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").HasDataLabel = " + CStr(Points(i).HasDataLabel) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").InvertIfNegative = " + CStr(Points(i).InvertIfNegative) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").Name = " + CStr(Points(i).Name) & vbNewLine
        Info = Info & "Point(" + CStr(i) + ").PictureType = " + CStr(Points(i).PictureType) & vbNewLine
        Info = Info & vbNewLine
        
    Next i
    MsgBox (Info)
End Sub

Rem Point: Set the Explosion (Only for Pie/Dughut Chart)
Sub Point_Explosion_Reset()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    Points(2).Explosion = 0
End Sub

Rem Point: Set the Explosion (Only for Pie/Dughut Chart)
Sub Point_Explosion_Set()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    Points(2).Explosion = 20
End Sub

Sub cbHas3DEffect_Click()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    If ActiveSheet.cbHas3DEffect.Value = True Then
        Points(2).Has3DEffect = True
    Else
        Points(2).Has3DEffect = False
    End If
End Sub

Sub cbInvertIfNegative_Click()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Set Points = SC(1).Points
    If ActiveSheet.cbInvertIfNegative.Value = True Then
        Points(2).InvertIfNegative = True
    Else
        Points(2).InvertIfNegative = False
    End If
End Sub
