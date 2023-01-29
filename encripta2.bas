'Esteban Gómez Palomo.
'egp.curso.bbdd@gmail.com
'29/01/2023
'Función para encritar y desencriptar strings para Excel y Access usando VBA.

Function Encripta2(Texto As String) As Variant
Dim lTexto As Long
Dim mTexto() As Long
Dim J As Integer
Dim Resultado As Variant

    If Texto = vbNullString Then
        Encripta2 = "¡ERROR!"
        Exit Function
    End If
    
    lTexto = Len(Texto)
    
    ReDim mTexto(1 To lTexto)
    
    For J = 1 To lTexto Step 1
        mTexto(J) = Asc(Mid(Texto, J, 1))
    Next J
    
    Resultado = ""
    
    For J = 1 To lTexto Step 1
        Resultado = Resultado & Chr(255 - mTexto(J))
    Next J
    
    Encripta2 = Resultado

End Function
