Attribute VB_Name = "mod1"
Global base As New ADODB.Connection
Global rsUsuarios As New ADODB.Recordset

'Control de acceso
'================
Global bControlAcceso As Boolean

Sub main()

With base
.CursorLocation = adUseClient
.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbh24.mdb;Persist Security Info=False"
frmInicio.Show

End With

End Sub

Sub tbAlmacen()
With rsUsuarios
If .State = 1 Then .Close
.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub


Sub tbUsers()
With rsUsuarios
If .State = 1 Then .Close
.Open "SELECT * FROM tbUser", base, adOpenStatic, adLockOptimistic
End With
End Sub

Sub botones()
    Dim boton As Object
    Set boton = Controls.Add("VB.commandbutton", "Boton")
    boton.Visible = True
    boton.Caption = "Drinky94 ^^"
    boton.Width = 1250
    boton.Height = 250
    boton.Left = 150
    boton.Top = 900
End Sub
