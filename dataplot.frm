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
    
    For n = 1 To Len("F12")
    'For n = 1 To Len(Me.rightBottomRan.Value)
        'Y = Mid(Me.rightBottomRan, n, 1)
        Y = Mid("F12", n, 1)
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
            ElseIf Cells(LCounter, Ccounter).Value = "SI PRESENTA" Then
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

