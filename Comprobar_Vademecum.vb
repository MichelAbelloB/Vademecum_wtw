Sub Vademecum3()

 Dim cont As Long
    Dim ultLinea As Long
    Dim codigo2 As Variant
    Dim codigo As Variant
    Dim rango As Variant
    
    ultLinea = Sheets("WTW").Range("C" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("3").Range("D2:E3454")
    
    For cont = 8 To ultLinea
        codigo = Sheets("WTW").Cells(cont, 3)
        codigo2 = Application.VLookup(codigo, rango, 1, False)
        
        If IsError(codigo2) Then
            codigo2 = "No cruza"
            Else
            codigo2 = "Cruza"
        End If
        
        Sheets("WTW").Cells(cont, 13) = codigo2
        
    Next cont
    
    MsgBox "Buscarv con macro ejecutada exitosamente!", vbInformation, "BuscarV"
    
    
End Sub

Sub Vademecum2()

 Dim cont As Long
    Dim ultLinea As Long
    Dim codigo2 As Variant
    Dim codigo As Variant
    Dim rango As Variant
    
    ultLinea = Sheets("WTW").Range("C" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("2").Range("B3:G4042")
    
    For cont = 8 To ultLinea
        codigo = Sheets("WTW").Cells(cont, 3)
        codigo2 = Application.VLookup(codigo, rango, 1, False)
        
        If IsError(codigo2) Then
            codigo2 = "No cruza"
            Else
            codigo2 = "Cruza"
        End If
        
        Sheets("WTW").Cells(cont, 12) = codigo2
        
        
        
    Next cont
    
    MsgBox "Buscarv con macro ejecutada exitosamente!", vbInformation, "BuscarV"
    
    
End Sub


Sub Vademecum1()

 Dim cont As Long
    Dim ultLinea As Long
    Dim codigo2 As Variant
    Dim codigo As Variant
    Dim rango As Variant
    
    ultLinea = Sheets("WTW").Range("C" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("1").Range("C2:C2397")
    
    For cont = 8 To ultLinea
        codigo = Sheets("WTW").Cells(cont, 3)
        codigo2 = Application.VLookup(codigo, rango, 1, False)

        If IsError(codigo2) Then
            codigo2 = "No cruza"
            Else
            codigo2 = "Cruza"
        End If
        
        Sheets("WTW").Cells(cont, 11) = codigo2
        
        
    Next cont
    
    MsgBox "Buscarv con macro ejecutada exitosamente!", vbInformation, "BuscarV"
    
    
End Sub

Sub MostrarFormulario()
    
    Formulariovalidacion.Show
    
End Sub

Sub Limpiar()

    Range("A5:H1048576").ClearContents
    
End Sub

Sub Limpiar_WTW()

    Range("K8:M1048576").ClearContents
    
End Sub