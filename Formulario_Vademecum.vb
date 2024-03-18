Private Sub CommandButton1_Click()

    numdatos = Hoja1.Range("B" & Rows.Count).End(xlUp).Row
    LISTA = Clear
    LISTA.RowSource = Clear
    
    y = 0
    
    For fila = 8 To numdatos
    appe = ActiveSheet.Cells(fila, 14).Value
    
    If appe Like "*" & Me.ComboBox2.Value & "*" Then
    
    Me.LISTA.AddItem
    Me.LISTA.List(y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LISTA.List(y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LISTA.List(y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LISTA.List(y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LISTA.List(y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LISTA.List(y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LISTA.List(y, 6) = ActiveSheet.Cells(fila, 7).Value
    Me.LISTA.List(y, 7) = ActiveSheet.Cells(fila, 8).Value
    y = y + 1
    
    End If
    
    Next

End Sub

Private Sub CommandButton2_Click()

    numdatos = Hoja1.Range("B" & Rows.Count).End(xlUp).Row
    LISTA = Clear
    LISTA.RowSource = Clear
    
    y = 0
    
    For fila = 8 To numdatos
    appe = ActiveSheet.Cells(fila, 15).Value
    
    If appe Like "*" & Me.ComboBox2.Value & "*" Then
    
    Me.LISTA.AddItem
    Me.LISTA.List(y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LISTA.List(y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LISTA.List(y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LISTA.List(y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LISTA.List(y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LISTA.List(y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LISTA.List(y, 6) = ActiveSheet.Cells(fila, 7).Value
    Me.LISTA.List(y, 7) = ActiveSheet.Cells(fila, 8).Value
    y = y + 1
    
    End If
    
    Next
    
End Sub

Private Sub CommandButton3_Click()

    numdatos = Hoja1.Range("B" & Rows.Count).End(xlUp).Row
    LISTA = Clear
    LISTA.RowSource = Clear
    
    y = 0
    
    For fila = 8 To numdatos
    appe = ActiveSheet.Cells(fila, 16).Value
    
    If appe Like "*" & Me.ComboBox2.Value & "*" Then
    
    Me.LISTA.AddItem
    Me.LISTA.List(y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LISTA.List(y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LISTA.List(y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LISTA.List(y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LISTA.List(y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LISTA.List(y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LISTA.List(y, 6) = ActiveSheet.Cells(fila, 7).Value
    Me.LISTA.List(y, 7) = ActiveSheet.Cells(fila, 8).Value
    y = y + 1
    
    End If
    
    Next
    
End Sub

Private Sub ExpData_Click()

Dim i As Integer
Dim uf As Integer

With Hoja9

    uf = 5
    
    For i = 0 To Me.LISTA.ListCount - 1
    
        .Cells(uf, 1) = Me.LISTA.List(i, 0)
        .Cells(uf, 2) = Me.LISTA.List(i, 1)
        .Cells(uf, 3) = Me.LISTA.List(i, 2)
        .Cells(uf, 4) = Me.LISTA.List(i, 3)
        .Cells(uf, 5) = Me.LISTA.List(i, 4)
        .Cells(uf, 6) = Me.LISTA.List(i, 5)
        .Cells(uf, 7) = Me.LISTA.List(i, 6)
        .Cells(uf, 8) = Me.LISTA.List(i, 7)
        
        uf = uf + 1
        
    Next i
    
End With

End Sub

Private Sub UserForm_Activate()

Me.LISTA.RowSource = "Tabla_WTW"
Me.LISTA.ColumnCount = 16

End Sub