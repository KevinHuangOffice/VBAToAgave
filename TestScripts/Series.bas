Attribute VB_Name = "Series"
Sub Series_Test_Reset()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Call Series_Explosion_Reset
End Sub
Rem Point: Display properties
Sub Series_Properties_Get()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    Dim Info As String
    Info = vbNullString
    For i = 1 To SC.Count
        Info = Info & "Series(" + CStr(i) + ").Explosion = " + CStr(SC(i).Explosion) & vbNewLine
        Info = Info & "Series(" + CStr(i) + ").Has3DEffect = " + CStr(SC(i).Has3DEffect) & vbNewLine
        Info = Info & "Series(" + CStr(i) + ").InvertIfNegative = " + CStr(SC(i).InvertIfNegative) & vbNewLine
        Info = Info & vbNewLine
    Next i
    MsgBox (Info)
End Sub

Sub Series_Explosion_Reset()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    SC(1).Explosion = 0
End Sub

Sub Series_Explosion_Set()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    SC(1).Explosion = 20
End Sub

Sub cbHas3DEffect_Click()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    If ActiveSheet.cbHas3DEffect.Value = True Then
        SC(1).Has3DEffect = True
    Else
        SC(1).Has3DEffect = False
    End If
End Sub

Sub cbInvertIfNegative_Click()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    If ActiveSheet.cbInvertIfNegative.Value = True Then
        SC(1).InvertIfNegative = True
    Else
        SC(1).InvertIfNegative = False
    End If
End Sub

Sub cbShadow_Click()
    Set SC = ActiveSheet.ChartObjects(1).Chart.SeriesCollection
    If ActiveSheet.cbShadow.Value = True Then
        SC(1).Shadow = True
    Else
        SC(1).Shadow = False
    End If
End Sub
