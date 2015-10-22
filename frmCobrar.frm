VERSION 5.00
Begin VB.Form frmCobrar 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Caja"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9390
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   6360
      TabIndex        =   10
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CheckBox chkCobrar 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "C O B R A R"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir ticket"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtefectivo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5400
      TabIndex        =   2
      Top             =   3510
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1320
      TabIndex        =   11
      Top             =   5880
      Width           =   11295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      TabIndex        =   6
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL A PAGAR"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Des As Byte = 10
Const Ca As Byte = 3
Const Pre As Byte = 5
Const TPre As Byte = 9
Private Sub chkDto_Click()
If chkDto.Value = 1 Then
    txtDto.Enabled = True
End If
End Sub

Private Sub cmdApply_Click()
Label1.Caption = CCur(Label1.Caption) - (CCur(Label1.Caption) * CCur(txtDto) / 100)
End Sub

Private Sub chkCobrar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
          Dim canal%
          Dim Impresora
          Impresora = "\\127.0.0.1\" & Printer.DeviceName
          canal = FreeFile
          Open Impresora For Output As #canal
          'Drawer Kick (ESC p)
          Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
          Close #canal
    End If
End Sub

Private Sub chkCobrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkCobrar.Value = 1
End If
End Sub

Private Sub txtefectivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
          Dim canal%
          Dim Impresora
          Impresora = "\\127.0.0.1\" & Printer.DeviceName
          canal = FreeFile
          Open Impresora For Output As #canal
          'Drawer Kick (ESC p)
          Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
          Close #canal
End If
End Sub

Public Sub chkCobrar_Click()
    
    '================================
    'CREAR REGISTRO EN LA TABLA TICKET
    '================================
    Dim codigoProd As Integer
    
    'Guardar en la tabla tbTicket
    tbTicket
    With rsUsuarios
        .Requery
        If .BOF = False Then
            If .EOF = False Then
                .MoveLast
            End If
            codigoProd = !idtbTicket + 1
        Else
            codigoProd = 1
        End If
        
        .Requery
        .AddNew
            !idtbTicket = codigoProd
            !TFecha = Date
            !Thora = Time
            drHoraPrint = Format(Now, "hh:mm:ss")
            !VentaT = CDbl(frmCaja.Label11.Caption)
            drTotalV = !VentaT
            !idTurno = idTurnos
        .Update
        
    End With
    
    'Guardar en la tabla tbTurnos
    tbTurnos
    With rsTurnos
        .Requery
        .MoveFirst
        drTotalV = drTotalV + !Tvendido
        !Tvendido = drTotalV
        .UpdateBatch
    End With
    
    '=======================================
    'AGREGAR PRODUCTOS EN LA TABLA TICKETPROD
    '======================================
            
    Dim i As Integer
    Dim codC As String
    Dim pvp, iva, neto As Double
    
    Dim numTicket As Long
    tbTicketProd
    With frmCaja.fldCaja
        For i = 0 To .Rows - 2
    
            With rsUsuarios
                .Requery
                .AddNew
                
                    frmCaja.fldCaja.Row = i + 1
                    
                    !idtbTicket = codigoProd
                    idTicketPrint = codigoProd
                    
                    !idTProd = i + 1
                    
                    frmCaja.fldCaja.Col = 6
                    codC = frmCaja.fldCaja.Text
                    !idProd = codC
                    
                    frmCaja.fldCaja.Col = 7
                    !idProdManual = frmCaja.fldCaja.Text
                    
                    frmCaja.fldCaja.Col = 2
                    !nomProd = frmCaja.fldCaja.Text
                    
                    frmCaja.fldCaja.Col = 1
                    !Tuni = frmCaja.fldCaja.Text
                    
                    frmCaja.fldCaja.Col = 3
                    
                    !pvpProd = CDec(frmCaja.fldCaja.Text)
                    
                    frmCaja.fldCaja.Col = 4
                    !dtoProd = frmCaja.fldCaja.Text
                    
                    frmCaja.fldCaja.Col = 5
                    pvp = CDec(frmCaja.fldCaja.Text)
                    !PrecioF = pvp
                    
                    frmCaja.fldCaja.Col = 8
                    iva = CDec(frmCaja.fldCaja.Text)
                    !ivaProd = iva
                    
                    neto = CDec(pvp) - ((CDec(pvp) * CDec(iva)) / 100)
                    !netoProd = CDec(neto)
                
                .Update
                
            End With
            
        Next i
    End With
    '===============================
    'DESCONTAR PRODUCTOS DEL ALMACEN
    '===============================
    
    tbAlmacen
    Dim x As String
    With frmCaja.fldCaja
        For i = 0 To frmCaja.fldCaja.Rows - 2
            .Row = i + 1
            .Col = 6
            rsUsuarios.Requery
            rsUsuarios.Find "idProd = '" & .Text & "'"
            .Col = 1
            x = .Text
            With rsUsuarios
                If !promoProd = True Then
                    !uniProd = !uniProd - (Val(frmCaja.fldCaja.Text) * !llevaProd)
                Else
                    !uniProd = !uniProd - Val(frmCaja.fldCaja.Text)
                End If
                .UpdateBatch
            End With
        
        Next i
    End With
    
    
    cmdCancelar.Visible = False
    chkCobrar.Enabled = False
    cmdSalir.Visible = True
    cmdPrint.Visible = True
    
    cmdSalir.SetFocus

End Sub

Private Sub cmdCal_Click()
Dim resta As Currency
If txtefectivo = "" Then
          MsgBox "Debe introducir el efectivo", vbInformation, Me.Caption
          txtefectivo.SetFocus
Else
          resta = CCur(txtefectivo.Text) - CCur(Label1.Caption)
          Label3(1).Caption = "Cambio: "
          
          Label4.Caption = CCur(resta)
          If resta < 0 Then
                    MsgBox "El efectivo no es suficiente", vbExclamation, Me.Caption
                    Exit Sub
          End If
          Label4.Caption = Format(Label4.Caption, "###,##0.00") & "€"
          If txtefectivo <> "" Then
                    chkCobrar.Visible = True
          End If
End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim r As String * Des
Dim s As String * Ca
Dim v As String * Pre
Dim TV As String * TPre
Dim b As String
Dim dato As String
Dim datoB As String
Dim i, u As Long
Dim a, e As Long
Dim Lcalc As Integer

'___________________________________________
Dim canal%
Dim Impresora

'Controlar errores de la impresora.


Impresora = "\\127.0.0.1\" & Printer.DeviceName
canal = FreeFile
Open Impresora For Output As #canal

Print #canal, Chr$(&H1B); "@"; 'INICIO (ESC @)
Print #canal, Chr$(&H1B); "a"; Chr$(1); 'CENTRADO
Print #canal, Chr$(&H1B); "t"; Chr$(19); 'Caracteres en español
Print #canal, Chr$(&H1B); "!"; Chr$(17); 'TAMAÑO LETRA (ESC !)

Print #canal, "Tienda STITCH"; Chr$(&HA);
Print #canal, Chr$(&H1B); "!"; Chr$(1); 'TAMAÑO LETRA (ESC !)

Print #canal, "Pan y Caprichos";
Print #canal, Chr$(&H1B); "d"; Chr$(2); '3 SALTOS DE LINEA(ESC d)

Print #canal, "C/ Cristobal Sanz Nº 86"; Chr$(&HA);
Print #canal, "CP. 03201"; Chr$(&HA);
Print #canal, "Elche - Alicante"; Chr$(&HA);
Print #canal, "Tlf 966109608"; Chr$(&HA);
Print #canal, "CIF 74246484T"; Chr$(&HA);
Print #canal, Chr$(&HA); 'SALTO DE LINA

'==============================
tbTicket
With rsUsuarios
.Find "idtbTicket = '" & idTicketPrint & "'"
'==============================

Print #canal, "Fecha: "; !TFecha; Chr$(&HA);
Print #canal, "Hora: "; !Thora; Chr$(&HA);
Print #canal, Chr$(&H1B); "a"; Chr$(0); 'IZQUIERDA
Print #canal, Chr$(&HA); 'SALTO DE LINEA
Print #canal, "Nº Factura simple:  "; !idtbTicket; Chr$(&HA);

End With

Print #canal, "------------------------------"; Chr$(&HA);

'==============================
With frmCaja.fldCaja
          dato = ""
          For i = 1 To .Cols - 1
                    r = CStr(.TextMatrix(0, i))
                    s = CStr(.TextMatrix(0, i))
                    v = CStr(.TextMatrix(0, i))
                    If i = 1 Then
                              s = CStr(.TextMatrix(0, i))
                              dato = (dato + s + " ")
                    ElseIf i = 2 Then
                              r = CStr(.TextMatrix(0, i))
                              dato = (dato + r + " ")
                    ElseIf i = 3 Then
                              v = CStr(.TextMatrix(0, i))
                              dato = (dato + v + " ")
                    ElseIf i = 5 Then
                              v = CStr(.TextMatrix(0, i))
                              dato = (dato + v + " ")
                    End If

          Next i
'================================

Print #canal, dato; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);

'=================================
          For u = 1 To .Rows - 1
                    dato = ""

                    For i = 1 To .Cols - 1
                              r = CStr(.TextMatrix(u, i))
                              s = CStr(.TextMatrix(u, i))
                              v = CStr(.TextMatrix(u, i))

                              If i = 1 Then

                                        If Len(.TextMatrix(u, i)) >= 3 Then
                                                  dato = (dato & s + " ")
                                        Else
                                                  Lcalc = 3 - Len(.TextMatrix(u, i))
                                                  s = Space(Lcalc) + .TextMatrix(u, i)
                                                  dato = (dato + s + " ")
                                        End If

                              ElseIf i = 5 Then
                                        b = .TextMatrix(u, i)
                                        b = Format(b, "###,##0.00")
                                        If Len(b) >= 5 Then
                                                  v = b
                                                  dato = (dato & v)
                                        Else

                                                  Lcalc = 5 - Len(b)
                                                  v = Space(Lcalc) + b
                                                  dato = (dato + v)
                                        End If
                              ElseIf i = 3 Then

                                        b = .TextMatrix(u, i)
                                        b = Format(b, "###,##0.00")
                                        If Len(b) >= 5 Then
                                                  v = b
                                                  dato = (dato & v & " ")
                                        Else

                                                  Lcalc = 5 - Len(b)
                                                  v = Space(Lcalc) + b
                                                  dato = (dato + v & " ")
                                        End If

                              Else

                                        'dato = Space(M) + dato
                                        If .ColWidth(i) = 0 Then
                                        Else
                                                  If i = 3 Or i = 4 Then
                                                  Else
                                                            dato = (dato & r + " ")
                                                  End If
                                        End If
                              End If

                    Next i
'==============================================

Print #canal, dato; Chr$(&HA);

'==============================================
          Next u

End With
'==============================================
Print #canal, "------------------------------"; Chr$(&HA);

Print #canal, Chr$(&H1B); "a"; Chr$(1); 'CENTRADO
Print #canal, Chr$(&H1B); "!"; Chr$(17); 'TAMAÑO DE LETRA
dato = frmCaja.Label11.Caption
dato = Replace(dato, "€", " ")
Print #canal, "TOTAL: "; dato; Chr$(&HA);
'=======================================
txtefectivo = Format(txtefectivo, "###,##0.00")
'=======================================
Print #canal, Chr$(&H1B); "!"; Chr$(1); 'TAMAÑO DE LETRA
Print #canal, Chr$(&HA);
Print #canal, "Efectivo: "; txtefectivo.Text; Chr$(&HA);
Print #canal, "--------------------------"; Chr$(&HA);
dato = Label4.Caption
dato = Replace(dato, "€", " ")
Print #canal, "Cambio: "; dato; Chr$(&HA);
Print #canal, Chr$(&H1B); "d"; Chr$(2); '3 SALTOS DE LINEA(ESC d)

Print #canal, Chr$(&H1B); "a"; Chr$(0); 'CENTRADO
Print #canal, "------------------------------"; Chr$(&HA);

'==================================
tbIva (idTicketPrint)

Dim d As String * Ca
Dim q As String
          dato = ""
          d = "Can"
          dato = d & " "
          d = "IVA"
          dato = dato & " " & d & " "
          v = "B.IMP"
          dato = dato & v & " "
          v = "Cuota"
          dato = dato & v & " "

'==================================

Print #canal, dato; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);

          dato = ""
'==================================
With rsIVA

          .MoveFirst
          For i = 1 To .RecordCount

                    d = !Cant
                    If Len(!Cant) >= 3 Then
                              dato = dato + d + " "
                    Else
                              Lcalc = 3 - Len(!Cant)
                              d = Space(Lcalc) & !Cant
                              dato = (dato + d + " ")
                    End If
                    d = !ivaProd
                    If Len(!ivaProd) >= 3 Then
                              dato = dato + d + " "
                    Else
                              Lcalc = 3 - Len(!ivaProd)
                              d = Space(Lcalc) & !ivaProd
                              dato = (dato + d + "%")
                    End If
                    q = !sNeto
                    q = Format(q, "###,##0.00")
                    v = q

                    If Len(q) >= 5 Then
                              dato = dato + " " + v
                    Else
                              Lcalc = 5 - Len(q)
                              v = Space(Lcalc) & q
                              dato = (dato + " " + v)
                    End If
                    q = !r
                    q = Format(q, "###,##0.00")
                    v = q

                    If Len(q) >= 5 Then
                              dato = dato + " " + v
                    Else
                              Lcalc = 5 - Len(q)
                              v = Space(Lcalc) & q
                              dato = (dato + " " + v)
                    End If
'=======================================

Print #canal, dato; Chr$(&HA);

'=======================================

                    dato = ""
                    If .EOF = False Then .MoveNext
          Next i
          Dim Riva, RivaA, RivaB, RivaC, RBiva, BivaA, BivaB, BivaC As Currency

          .MoveFirst
          If .EOF = False Then
                     BivaA = !sNeto
                     RivaA = !r
                    .MoveNext
          Else
                    BivaA = 0
                    RivaA = 0
          End If
          If .EOF = False Then
                     BivaB = !sNeto
                     RivaB = !r
                    .MoveNext
          Else
                    BivaB = 0
                    RivaB = 0
          End If
          If .EOF = False Then
                    BivaC = !sNeto
                    RivaC = !r
          Else
                    BivaC = 0
                    RivaC = 0
          End If

          RBiva = BivaA + BivaB + BivaC
          Riva = RivaA + RivaB + RivaC
          RBiva = Format(RBiva, "###,##0.00")
          If Len(RBiva) >= 5 Then
                    dato = RBiva
          Else
                    Lcalc = 5 - Len(RBiva)
                    v = Space(Lcalc) & RBiva
                    dato = v
          End If
          Riva = Format(Riva, "###,##0.00")
          If Len(Riva) >= 5 Then
                    dato = Riva
          Else
                    Lcalc = 5 - Len(Riva)
                    v = Space(Lcalc) & Riva
                    datoB = v

                    dato = dato & " " & datoB
          End If
'======================================
Print #canal, Chr$(&H1B); "t"; Chr$(19); 'Caracteres en español
Print #canal, Chr(238); Chr$(&HA);

Print #canal, "TOTAL" & Space(3) & dato

End With

Print #canal, Chr$(&H1D); "V"; Chr$(66); Chr$(0); 'Feeds paper & cut

Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
Print #canal, Chr$(&H1B); "d"; Chr$(15); '3 SALTOS DE LINEA(ESC d)
Close #canal


End Sub

Private Sub cmdSalir_Click()

Dim boton As String

frmCaja.Enabled = True
frmCaja.fldCaja.Clear
 
 'Quitar el titulo del SSTAB
 frmCaja.SSTab2.Caption = ""
 frmCaja.Label10.Caption = ""
 frmCaja.Label11.Caption = ""
 
 'Dibujar el GRID
 Filas = 1
 With frmCaja.fldCaja
          .Rows = Filas
          .Rows = 1
          .Cols = 9
          
          'Ocultar la primer columna del grid
          .ColWidth(0) = 0
          
          'Poner los ENCABEZADOS del GRID
          
          .ColWidth(1) = 703
          .Row = 0
          .Col = 1
          .Text = "Can"
          
          .ColWidth(2) = 3703
          .Row = 0
          .Col = 2
          .ColAlignment(2) = 4
          .Text = "Detalle"
          
          .ColWidth(3) = 1033
          .Row = 0
          .Col = 3
          .Text = "Pre.U"
          
          .ColWidth(4) = 803
          .Row = 0
          .Col = 4
          .Text = "Dto."
          
          .ColWidth(5) = 1403
          .Row = 0
          .Col = 5
          .Text = "Pre.T"
          
          .ColWidth(6) = 0
          .Row = 0
          .Col = 6
          .Text = "Cod"
          
          .ColWidth(7) = 0
          .Row = 0
          .Col = 7
          .Text = "CodM"
          
          .ColWidth(8) = 0
          .Row = 0
          .Col = 8
          .Text = "IVA"
     
 End With

 Fila = 1
 Unload Me

End Sub

Private Sub Form_Load()
cmdSalir.Enabled = True
frmCaja.Enabled = False
If Not frmCaja.Label11 = "" Then
    Label1.Caption = CCur(frmCaja.Label11)
    Label1.Caption = Format(Label1.Caption, "###,##0.00") & "€"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmCaja.Enabled = True
End Sub


Private Sub txtefectivo_KeyPress(KeyAscii As Integer)
On Error GoTo error
If KeyAscii = 13 Then
cmdCal_Click
chkCobrar.SetFocus
End If
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    End If
error:
End Sub

Function DetalleTicket(MSFlexGrid As Object) As String
               
With MSFlexGrid
          Dim dato As String
          Dim i, u As Long
          u = 1
          For i = 1 To .Rows - 1
                    dato = dato + .TextMatrix(i, u)
                    u = u + 1
          Next i
          Printer.Print dato
          Printer.EndDoc
End With

End Function
