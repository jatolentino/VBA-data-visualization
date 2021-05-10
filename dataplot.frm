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
    
