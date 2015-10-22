Attribute VB_Name = "mod1"
Global base As New ADODB.Connection
Global rsUsuarios As New ADODB.Recordset
Global rsSort As New ADODB.Recordset
Global dtgConn As New ADODB.Recordset
Global rsProd As New ADODB.Recordset
Public rsPrint As New ADODB.Recordset
Public rsTurnos As New ADODB.Recordset
Public rsCaja As New ADODB.Recordset
Public rsVentaT As New ADODB.Recordset
Global rsAlmacen As New ADODB.Recordset
Public rsIVA As New ADODB.Recordset
Public rsivaZ As New ADODB.Recordset
Public rsTurnoDetalle As New ADODB.Recordset
Public rsTurnoDetalleT As New ADODB.Recordset
Public rsCountFav As New ADODB.Recordset
Public rsPromo As New ADODB.Recordset
Public rsSortPromo As New ADODB.Recordset
Public rsBuscarProd As New ADODB.Recordset
Public rsCountFav2 As New ADODB.Recordset

'Control de acceso
'================
Global bControlAcceso As Boolean
Global selectId As String
Global SelectProd As String
Global SelectDetalleProv As String
Global sOperacion As String
Global sModidProd As String
Global sModidProdM As String
Global chkturno As Boolean
Public chkadmin As Boolean
Global sIdS As String
Global timerInicioTurno As String
Global fechaInicioTurno As String
Public actHora As String
Public Tventa As Double
Public idTurnos As Long
Public idTurno As String
Public idTurnoD As String
Public chkCaja As Boolean

'========================
'Variables para la caja registradora
'========================
Public Filas As Integer
Public Fila As Integer
Public Tot As Double
Public x As Double
Public xx As Double
Public y As Integer
Public cantT As Integer
Public countFav As Integer

'=====================
'Variables para imprimir
'=====================
Public idTicketPrint As Integer
Public drTotalV As Double
Public drHoraPrint As String

'====================
'Favoritos
'==================
Public X0(24) As String
Public X1(24) As String
Public X2(24) As String
Public X3(24) As String
Public X4(24) As String

'===========================
'DATOS DEL TICKET
'===========================
Public Titulo As String
Public Direc As String
Public telf As String
Public cif As String
Public TituloB As String
Public FechaHora As String
Public idTicket As String
Public Detalles As String
Public Total As String
Public Efectivo As String
Public Devolucion As String
Public TituloC As String
Public TituloIva As String
Public DetalleIVA As String
Public TituloD As String




'Global dato As New ADODB
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Test\Documents\h24\dbh24.mdb;Persist Security Info=False

Sub Main()
'Comprobar que no este abierta la Aplicación.
'============================================
If App.PrevInstance = True Then
          End
End If
With base
.CursorLocation = adUseClient
.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbh24.mdb;Persist Security Info=False"
MDIfrmMadre.Show

End With

'chkturno = False
'chkadmin = False
'chkCaja = False

End Sub
Sub tbSortAlmacen()
With rsSort
If .State = 1 Then .Close
.Open "tbAlmacen", base, adOpenStatic, adLockOptimistic, adCmdTable
'.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbSortPromo()
With rsSortPromo
If .State = 1 Then .Close
.Open "tbAlmacen", base, adOpenStatic, adLockOptimistic, adCmdTable
'.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbAlmacen()
With rsUsuarios
If .State = 1 Then .Close
.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbBuscarProd(bProd As String)
With rsBuscarProd
    If .State = 1 Then .Close
    .Open "SELECT * FROM tbAlmacen WHERE (idProd LIKE  '%" & bProd & "%') OR (nomProd LIKE '%" & bProd & "%')", base, adOpenStatic, adLockOptimistic
End With
End Sub

Sub tbPromo()
With rsPromo
If .State = 1 Then .Close
.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbCaja()
With rsCaja
If .State = 1 Then .Close
.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbAddProd()
With rsAlmacen
If .State = 1 Then .Close
.Open "SELECT * FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
End With
End Sub
Sub tbProv()
With rsUsuarios

If .State = 1 Then .Close
    .Open "SELECT * FROM tbProv", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub dtgConnSort()

With dtgConn
If .State = 1 Then .Close
    .Open "tbProv", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbProdProv()
With rsUsuarios

If .State = 1 Then .Close
    .Open "SELECT * FROM tbProdProv", base, adOpenStatic, adLockOptimistic
End With

End Sub

Sub tbUsers()
With rsUsuarios

If .State = 1 Then .Close
    .Open "SELECT * FROM tbUser", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbCajaProd(sCadena As String)
With rsUsuarios

If .State = 1 Then .Close
    .Open sCadena, base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbFavoritosCount(Favorito As String)
With rsAlmacen

If .State = 1 Then .Close
    .Open "SELECT COUNT( " & Favorito & ")AS Expre FROM tbAlmacen WHERE " & Favorito & " = 1", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbFav(Favorito As String)
With rsAlmacen

If .State = 1 Then .Close
    .Open "SELECT idProd,nomProd," & Favorito & " FROM tbAlmacen WHERE " & Favorito & " = 1", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbTicket()
With rsUsuarios

If .State = 1 Then .Close
    .Open "SELECT * FROM tbTicket ORDER BY  idtbTicket ASC", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbTicketProd()
With rsUsuarios

If .State = 1 Then .Close
    .Open "SELECT * FROM tbTicketProd", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbTicketProdPrint()
With rsPrint

If .State = 1 Then .Close
    .Open "SELECT * FROM tbTicketProd", base, adOpenStatic, adLockOptimistic
End With

End Sub

Sub tbTurnos()

With rsTurnos
    If .State = 1 Then .Close
        .Open "SELECT * FROM tbTurnos ORDER BY id DESC", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbVentaT()
Dim x As Data
Dim fActual As String
With rsVentaT
    If .State = 1 Then .Close
        If idTurnos < 0 Then Exit Sub
            .Open "SELECT SUM(VentaT) AS Expr2 FROM tbTicket WHERE idTurno = " & idTurnos & "", base, adOpenStatic, adLockOptimistic
End With

End Sub

 Sub tbIva(idtb)

 With rsIVA
          
          If .State = 1 Then .Close
              .Open "SELECT SUM(Tuni) AS Cant, ivaProd, Sum(PrecioF) AS sPre, Sum(netoProd) AS sNeto, (sPre-sNeto) AS R  FROM tbTicketProd WHERE idtbTicket = " & idtb & " GROUP BY ivaProd", base, adOpenStatic, adLockOptimistic


End With
End Sub
 Sub tbIvaZ(Turno)

 With rsivaZ
          
          If .State = 1 Then .Close
              .Open "SELECT ivaProd,SUM(netoProd) AS zNeto, SUM(PrecioF) AS zPrecio FROM (tbTicket INNER JOIN tbTicketProd ON tbTicket.idtbTicket = tbTicketProd.idtbTicket) INNER JOIN tbTurnos ON tbTicket.idTurno = tbTurnos.id WHERE tbTurnos.id = " & Turno & " GROUP BY ivaProd ", base, adOpenStatic, adLockOptimistic


End With
End Sub

Sub tbTurnoDetalle(idT As String)
With rsTurnoDetalle
          If .State = 1 Then .Close
          .Open "SELECT * FROM tbTicket WHERE idTurno = " & idT & " ", base, adOpenStatic, adLockOptimistic

End With

End Sub
Sub tbTurnoDetalleT(idT As String)
With rsTurnoDetalleT
          If .State = 1 Then .Close
          .Open "SELECT * FROM tbTicketProd WHERE idtbTicket = " & idT & "  ORDER BY idTProd ASC", base, adOpenStatic, adLockOptimistic
End With

End Sub
Sub tbCounFav(Favorito As String)
With rsCountFav
          If .State = 1 Then .Close
              .Open "SELECT COUNT(Favorito) FROM tbAlmacen WHERE " & Favorito & " = True", base, adOpenStatic, adLockOptimistic

End With

End Sub
Sub tbCountFav2()
With rsCountFav2
    If .State = 1 Then .Close
    .Open "SELECT SUM(favA) As FavA,SUM(favB) As FavB,SUM(favC) As FavC,SUM(favD) As FavD,SUM(favE) As FavE FROM tbAlmacen", base, adOpenStatic, adLockOptimistic
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
