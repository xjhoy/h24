Attribute VB_Name = "modVal"
'Solo TECLEAR NUMEROS
'=================
Function solNum(KeyAscii, cmdAceptar_Click) As Integer
If KeyAscii = 44 Then KeyAscii = 46
'If KeyAscii = 13 Then cmdAceptar_Click Else Exit Function
Dim comma, punto As String
End Function





Function ValidarCampoNumerico(nCampoNumerico) As Boolean

    
    If IsNumeric(nCampoNumerico) = False Then
        ValidarCampoNumerico = False
    Else
        ValidarCampoNumerico = True
    End If
 

End Function
