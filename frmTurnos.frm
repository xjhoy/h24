VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTurnos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Turnos"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3600
      Picture         =   "frmTurnos.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3600
      Picture         =   "frmTurnos.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   240
   End
   Begin VB.CommandButton cmdInformeZ 
      Caption         =   "Informe Z"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton cmdTurnoDetalle 
      Caption         =   "Ver turno en detalle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16200
      Picture         =   "frmTurnos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   4080
      TabIndex        =   6
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTurnos.frx":27DE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmTurnos.frx":27FA
         Height          =   6975
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12303
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "N. Turno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "TFechaIni"
            Caption         =   "Fecha de inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "iniTurno"
            Caption         =   "Inicio de turno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "TFechaFin"
            Caption         =   "Fecha de cierre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "finTurno"
            Caption         =   "Fin de turno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "TVendido"
            Caption         =   "Total vendido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   2145,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1874,835
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1785,26
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1769,953
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   6840
      Picture         =   "frmTurnos.frx":280F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TURNOS"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      TabIndex        =   0
      Top             =   600
      Width           =   7095
   End
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pre As Byte = 8

Private Sub cmdCerrar_Click()
    bControlAcceso = False
    MDIfrmMadre.Picture1.Visible = True
    Unload Me
End Sub

Private Sub cmdInformeZ_Click()
Dim idT  As String
Dim canal%
Dim Impresora

'On Error GoTo error

Impresora = "\\127.0.0.1\" & Printer.DeviceName
canal = FreeFile
Open Impresora For Output As #canal

Print #canal, Chr$(&H1B); "@"; 'INICIO (ESC @)
Print #canal, Chr$(&H1B); "a"; Chr$(1); 'CENTRADO
Print #canal, Chr$(&H1B); "t"; Chr$(19); 'Caracteres en español
Print #canal, Chr$(&H1B); "!"; Chr$(17); 'TAMAÑO LETRA (ESC !)

Print #canal, "Tienda STITCH"; Chr$(&HA);
Print #canal, Chr$(&H1B); "!"; Chr$(1); 'TAMAÑO LETRA (ESC !)
Print #canal, "Pan y Caprichos"; Chr$(&HA);
Print #canal, Chr$(&H1B); "d"; Chr$(2); '2 SALTOS DE LINEA(ESC d)

Print #canal, "C/ Cristobal Sanz Nº 86"; Chr$(&HA);
Print #canal, "CP. 03201"; Chr$(&HA);
Print #canal, "Elche - Alicante"; Chr$(&HA);
Print #canal, "Tlf 966109608"; Chr$(&HA);
Print #canal, "CIF 74246484T"; Chr$(&HA);
Print #canal, Chr$(&H1B); "d"; Chr$(2); '2 SALTOS DE LINEA(ESC d)

tbTurnos
With rsTurnos
          .Requery
          
          'Decimos que turno se va a imprimir, deacuerdo con el seleccionado
          .Find "id = '" & idTurno & "'"
                    idT = !id
Print #canal, "Turno: " & idT; Chr$(&HA);
Print #canal, Chr$(&H1B); "a"; Chr$(0); 'LEFT
Print #canal, "Apertura "; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);
Print #canal, "Fecha: " & !TFechaIni; Chr$(&HA);
Print #canal, "Hora: " & !iniTurno; Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, "Cierre"; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);
Print #canal, "Fecha: " & !TFechaFin
Print #canal, "Hora: " & !finTurno
Print #canal, Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, "VENTAS"
Print #canal, "------------------------------"; Chr$(&HA);

tbTurnoDetalle (idTurno)
With rsTurnoDetalle
          Dim Cuenta As Long
          Dim TV As String
          Cuenta = .RecordCount
          .Close
End With
Print #canal, "Numero de facturas: " & Cuenta; Chr$(&HA);
Print #canal, Chr$(&HA);
TV = !Tvendido
TV = Format(TV, "###,##0.00")
Print #canal, "Total vendido: " & TV; "e"; Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, "FORMAS DE PAGO"; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);
Print #canal, "Contado: " & TV; "e"; Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, Chr$(&HA);
Print #canal, "DESGLOSE DE IMPUESTOS"; Chr$(&HA);
Print #canal, "------------------------------"; Chr$(&HA);
Print #canal, Chr$(&HA);
End With
tbIvaZ (idTurno)
With rsivaZ
          Dim zIva(10), zNeto(10), zPrecio(10), zCuota(10) As String
          Dim zC As Long
          Dim dato As String
          Dim zR As String * Pre
          zR = "Venta"
          dato = dato & zR & " "
          zR = "IMP%"
          dato = dato & zR & " "
          zR = "B.IMP"
          dato = dato & zR & " "
          zR = "Cuota"
          dato = dato & zR & " "
Print #canal, dato
          dato = ""
          
          .Requery
          For zC = 1 To .RecordCount
                    If .EOF Or .BOF Then
                              zIva(zC) = 0
                              zNeto(zC) = 0
                              zPrecio(zC) = 0
                              zCuota(zC) = 0
                              .MoveNext
                    Else
                              zIva(zC) = !ivaProd
                              zNeto(zC) = !zNeto
                              zPrecio(zC) = !zPrecio
                              zCuota(zC) = !zPrecio - !zNeto
                              .MoveNext
                    End If
                    zPrecio(zC) = Format(zPrecio(zC), "###,##0.00")
                    zNeto(zC) = Format(zNeto(zC), "###,##0.00")
                    zCuota(zC) = Format(zCuota(zC), "###,##0.00")
                    zR = zPrecio(zC) & "e"
                    dato = dato & zR & " "
                    zR = zIva(zC) & "%"
                    dato = dato & zR & " "
                    zR = zNeto(zC) & "e"
                    dato = dato & zR & " "
                    zR = zCuota(zC) & "e"
                    dato = dato & zR & " "
                    
                    
Print #canal, dato; Chr$(&HA);
                    dato = ""
          Next zC
Print #canal, Chr$(&H1B); "d"; Chr$(15); '3 SALTOS DE LINEA(ESC d)
Close #canal

End With

'error:

End Sub

Private Sub cmdTurnoDetalle_Click()
On Error GoTo error
If Not idTurno = "" Then
With rsTurnos
          If .State = 1 Then
                    .Close
          End If
End With
          frmTurnoDetalle.Show
          Unload Me
Else
error:
          MsgBox "Antes debe seleccionar un turno", vbExclamation, Me.Caption
End If

End Sub

'-- Guardar el idTurno al dar click en la tabla --
'=================================================

Private Sub DataGrid1_Click()
    On Error GoTo error
    With rsTurnos
        If .State = 1 Then
            If .BOF Or .EOF Then
                      Exit Sub
            End If
            
            .Find "id= '" & (DataGrid1.Columns(0).Text) & "'"
            idTurno = DataGrid1.Columns(0).Text
            
            If KeyAscii = 13 Then
                cmdTurnoDetalle_Click
            Else
                Exit Sub
            End If
            
            Exit Sub
        End If
        
    End With
    
error:

End Sub

Private Sub Form_Load()
SSTab1.Caption = ""
tbTurnos
Set DataGrid1.DataSource = rsTurnos
End Sub

Private Sub Form_Unload(Cancel As Integer)
With rsTurnos
    If .State = 1 Then .Close
End With
End Sub
