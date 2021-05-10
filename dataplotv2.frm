Private Sub cancelar_Click()
    Unload Me
End Sub

Private Sub filasnum_Change()

End Sub
Private Sub UserForm_Initialize()
Image1.Picture = LoadPicture("C:\Users\Josetv\Documents\logoinst.jpg")
End Sub

Private Sub generar_Click()
'Sub boton()
    Dim LCounter As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim Ccounter As Integer
    i = 0
    j = 0
    k = 0
    l = 0
    m = 0
    Dim pi As Integer
    pi = 1
    Dim n As Long
    Dim Y As String
    Dim filaNum1 As String
    Dim colNum1 As String
    Dim filaNum2 As String
    Dim colNum2 As String

    'ColumnLetter = Me.TextBox1.Value
    'Cells(22, 2).Value = Range(Me.TextBox1.Value & 1).Column
    
    For n = 1 To Len("B3")
    'For n = 1 To Len(Me.lefTopRan.Value)
        'Y = Mid(Me.lefTopRan.Value, n, 1)
        Y = Mid("B3", n, 1)
        If Y Like "[A-Za-z ]" Then
            colNum1 = colNum1 & Y
        ElseIf Y Like "[0-9.]" Then
            filaNum1 = filaNum1 & Y
        End If
    Next n
    Dim filaa As String
    filaa = "AV12"
    For n = 1 To Len(filaa)
    'For n = 1 To Len(Me.rightBottomRan.Value)
        'Y = Mid(Me.rightBottomRan, n, 1)
        Y = Mid(filaa, n, 1)
        If Y Like "[A-Za-z ]" Then
            colNum2 = colNum2 & Y
        ElseIf Y Like "[0-9.]" Then
            filaNum2 = filaNum2 & Y
        End If
    Next n
    
    'Cells(18, 2).Value = colNum1
    'Cells(19, 2).Value = filaNum1
    'Cells(20, 2).Value = colNum2
    'Cells(21, 2).Value = filaNum2

    'Cells(18, 2).Value = Asc(filaNum1) - 96
    'Cells(18, 3).Value = colNum1
    'Cells(19, 2).Value = Asc(filaNum2) - 96
    'Cells(19, 3).Value = colNum2
    
    'For Ccounter = colNum1 To colNum2
    For Ccounter = Range(colNum1 & 1).Column To Range(colNum2 & 1).Column
        For LCounter = filaNum1 To filaNum2
            If Cells(LCounter, Ccounter).Value = "NO PRESENTO" Then
                i = i + 1
            ElseIf Cells(LCounter, Ccounter).Value = "MODIFICAR" Then
                j = j + 1
            ElseIf Cells(LCounter, Ccounter).Value = "ACTUALIZAR" Then
                k = k + 1
            ElseIf Cells(LCounter, Ccounter).Value = "EN PARTE" Then
                l = l + 1
            ElseIf Cells(LCounter, Ccounter).Value = "SI PRESENTA" Or Cells(LCounter, Ccounter).Value = "SÍ PRESENTA" Then
                m = m + 1
            End If
        Next LCounter
        'Cells(filaNum2 + 3, Ccounter).Value = i
        'Cells(Asc(filaNum2) - 96 + 3, Ccounter).Value = i

        Cells(filaNum2 + 2, Ccounter + 1).Value = "Pregunta" & pi
        pi = pi + 1

        Cells(filaNum2 + 3, Ccounter + 1).Value = i
        i = 0

        Cells(filaNum2 + 4, Ccounter + 1).Value = j
        j = 0

        Cells(filaNum2 + 5, Ccounter + 1).Value = k
        k = 0

        Cells(filaNum2 + 6, Ccounter + 1).Value = l
        l = 0

        Cells(filaNum2 + 7, Ccounter + 1).Value = m
        m = 0
    Next Ccounter

    Cells(filaNum2 + 3, Range(colNum1 & 1).Column).Value = "NO PRESENTO"
    Cells(filaNum2 + 4, Range(colNum1 & 1).Column).Value = "MODIFICAR"
    Cells(filaNum2 + 5, Range(colNum1 & 1).Column).Value = "ACTUALIZAR"
    Cells(filaNum2 + 6, Range(colNum1 & 1).Column).Value = "EN PARTE"
    Cells(filaNum2 + 7, Range(colNum1 & 1).Column).Value = "SI PRESENTA"

    'Cells(18, 2).Value = colNum1 & (filaNum2 + 2)
    'Cells(18, 3).Value = Chr(Asc(colNum2) - 32) & (filaNum2 + 2 + 4)

    Dim lefran
    Dim righran
    lefran = colNum1 & (filaNum2 + 2)
    'righran = Chr(Asc(colNum2) - 32) & (filaNum2 + 2 + 5)
    righran = Split(Cells(1, colNum2).Address, "$")(1) & (filaNum2 + 2 + 5)
    'Split(Cells(1, colNum2).Address, "$")(1)
    Cells(23, 2).Value = righran
    Cells(24, 2).Value = lefran
'    Cells(18, 4).Value = Chr(65)
'    Cells(18, 5).Value = Chr(66)
'    Cells(18, 6).Value = Chr(67)
'    Cells(18, 7).Value = Chr(68)
'    Cells(18, 8).Value = Chr(69)
'    Cells(18, 9).Value = Chr(90)

    Dim Chrt As ChartObject
    'Set Chrt = Sheets("Sheet2").ChartObjects.Add(Left:=180, Width:=270, Top:=7, Height:=210)
    Dim DataRng As Range
    Set Chrt = ActiveSheet.ChartObjects.Add(Left:=50, _
                                            Width:=1000, _
                                            Height:=1000, _
                                            Top:=10)

    Set DataRng = Range(lefran, righran)
    Chrt.Chart.SetSourceData Source:=DataRng.CurrentRegion, PlotBy:=xlRows
    'Chrt.Chart.SetElement
    '(msoElementDataLabelOutSideEnd)
    'Chrt.Chart.SetElement (Format(msoElementDataLabelOutSideEnd(1) / 100, "0.00%#"))
    
    'ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowLabelAndPercent, LegendKey _
    '    :=False, HasLeaderLines:=True
    'ActiveChart.SeriesCollectLabels.NumberFormat = "0.0%"ion(1).Data
    'Set Chrt = ActiveChart.SeriesCollection(5).DataLabels.NumberFormat = "0%il\e"
    
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    
    'Chrt.Chart.SetElement (msoElementDataLabelNone)
    'Chrt.Chart.SetElement (msoElementDataLabelShow)
    'Chrt.Chart.SetElement xlDataLabelsShowLabelAndPercent
    
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    'Chrt.Chart.SeriesCollection(1).XValues.Apply
    Dim srs As Series
    
    Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    'Chrt.Chart.SetElement msoElementPrimaryCategoryAxisTitleHorizontal
    'Chrt.Chart.SetElement msoElementPrimaryCategoryAxisTitleBelowAxis
    'Chrt.Chart.SeriesCollection(1).DataLabels.ShowPercentege te
    'Chrt.Chart.ChartObjects(1).At
    'Chrt.Chart.ApplyDataLabels Type:=xlDataLabelsShowPercent
    
    'Chrt.Chart.SetElement (msoElementLegendTop)
    'Chrt.Chart.ApplyDataLabels Type:=xlDataLabelsShowPercent
    'klChrt.Chart.ApplyDa
    'DataLabelsShowPercent
'    ActiveSheet.ChartObjects(1).Activate
'    ActiveChart.SeriesCollection(1).DataLabels.ShowPercentage = True
    'Chrt.Chart.ChartType = xlBarStacked100
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    Chrt.Chart.Axes(xlValue).MajorGridlines.Delete
    Chrt.Chart.ChartType = xlBarStacked100
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    
    Chrt.Chart.SeriesCollection(1).ApplyDataLabels ShowCategoryName:=False, ShowValue:=False, ShowPercentage:=True
    'Chrt.Chart.SeriesCollection(1).ApplyDataLabels Type:=xlDataLabelsShowPercent
    'Chrt.Chart.SeriesCollection(1).DataLabels.Position = xlLabelPositionBestFit
    'Dim i As Integer
    'Dim j As Integer
    Dim s As Double
    Dim v As Variant

    Dim NoDigits As Integer
    Dim PercentFormat As String
    Dim myTxt As String
    NoDigits = 1 'How many digits to round Millions to
    PercentFormat = "0.0%"
    For i = 1 To Chrt.Chart.SeriesCollection.Count
        Chrt.Chart.SeriesCollection(i).HasDataLabels = True
    Next i
     
    For i = 1 To Chrt.Chart.SeriesCollection(1).Points.Count
        s = 0
     
        For j = 1 To Chrt.Chart.SeriesCollection.Count
            v = Chrt.Chart.SeriesCollection(j).Values
            s = s + v(i)
        Next j
     
        For j = 1 To Chrt.Chart.SeriesCollection.Count
            v = Chrt.Chart.SeriesCollection(j).Values
            'myTxt = Round(v(i) / 1000000#, NoDigits) & "M, " & Format(v(i) / s, PercentFormat)
            myTxt = Format(v(i) / s, PercentFormat)
            'myTxt = Round(v(i) / 1000000#, NoDigits) & Format(v(i) / s, PercentFormat)
            Chrt.Chart.SeriesCollection(j).Points(i).DataLabel.Text = myTxt
     
            If v(i) <= 0 Then Chrt.Chart.SeriesCollection(j).Points(i).DataLabel.Delete
        Next j
    Next i
    'Format(msoElementDataLabelOutSideEnd, "0.00%#")
    'Chrt.Chart.ChartType = xl3DColumnStacked
    'Chrt.Chart.ChartType = xl3DColumnStacked100
    
    
    'Chrt.Chart.ChartObjects("Chart 1").Activate
    'Chrt.Chart.SeriesCollection(1).DataLabels.ShowPercentage = True
    'ActiveChart.FullSeriesCollection(1).DataLabels.ShowPercentage = True
    'NumberFormat = "0.0%"
    'ShowPercentage = True
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    
'    With PlotArea
'        Chrt.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0%"
'    End With
    
    'ActiveChart.PlotBy = xlColumns
    'ActiveChart.PlotBy = xlRows
    Chrt.Chart.HasTitle = True
    'Chrt.Chart.Location where:=xlLocationAsNewSheet

    'Chrt.Location xlLocationAsObject
    'Chrt.Chart.Location Name:=Sheet2
    'Chrt.Chart.Location Where:=xlLocationAsObject, Name:="Sheet2"
    'Set Chrt = Chrt.Location(Where:=xlLocationAsObject).Name:="Sheet2"
    'Chrt.Chart.Location Where:=xlLocationAsObject
    'Chrt.Chart Name:=Sheet2
    'Set embeddedchart = Sheets("Sheet1").Shapes.AddChart
    Dim newSheet As Worksheet
    'Dim sourceSheet As Worksheet
    ActiveWorkbook.Worksheets.Add(Before:=Application.Worksheets(1)).Name = "NewSheet"
    Set newSheet = Application.Worksheets("NewSheet")
    Chrt.Chart.Location xlLocationAsObject, "NewSheet"
    'Chrt.Chart.SetElement (msoElementDataLabelOutSideEnd)
    'Chrt.Chart.SetElement = msoElementDataLabelOutSideEnd
    'Set Chrt ¿ Active
    'Chrt.SetElement msoElementDataLabelOutSideEnd
    
    Unload Me
End Sub

'Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'Image1.Picture = LoadPicture("C:\Users\Josetv\Documents\logoinst")
'
'End Sub

Private Sub Label2_Click()

End Sub
