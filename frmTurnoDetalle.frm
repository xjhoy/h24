VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTurnoDetalle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Detalle de turno"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
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
      Left            =   15000
      Picture         =   "frmTurnoDetalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdTurnoDetalle 
      Caption         =   "Ver Factura simple en detalle"
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
      Left            =   840
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3840
      Picture         =   "frmTurnoDetalle.frx":1CCA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   240
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTurnoDetalle.frx":2254
      Height          =   10095
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   17806
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "idtbTicket"
         Caption         =   "Nº Fact. Simple"
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
         DataField       =   "Tfecha"
         Caption         =   "Fecha"
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
         DataField       =   "Thora"
         Caption         =   "Hora"
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
         DataField       =   "VentaT"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00 €"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "idTurno"
         Caption         =   "Cód. Turno"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1890,142
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1695,118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   2204,788
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTurnoDetalle.frx":2269
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE FACTURAS SIMPLES"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   5400
      Picture         =   "frmTurnoDetalle.frx":2285
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "frmTurnoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Des As Byte = 10
Const Ca As Byte = 3
Const Pre As Byte = 5
Const TPre As Byte = 9

Private Sub cmdCerrar_Click()
frmTurnos.Show
Unload Me
End Sub

Private Sub cmdTurnoDetalle_Click()
On Error GoTo error
If Not idTurnoD = "" Then
With rsTurnos
          If .State = 1 Then
                    .Close
          End If
End With
          frmTurnoDetalleT.Show
          Unload Me
Else
error:
          MsgBox "Antes debe seleccionar un turno", vbExclamation, Me.Caption
End If
End Sub

Private Sub DataGrid1_Click()
With rsTurnoDetalle
          If .State = 1 Then
                    If .BOF Or .EOF Then
                              Exit Sub
                    End If
                    
                    .Find "idtbTicket= '" & (DataGrid1.Columns(0).Text) & "'"
                    idTurnoD = DataGrid1.Columns(0).Text
                    
                    If KeyAscii = 13 Then
                              cmdTurnoDetalle_Click
                    Else
                              Exit Sub
                    End If
                    
                    Exit Sub
          End If
End With
End Sub

Private Sub Form_Load()
SSTab1.Caption = ""
tbTurnoDetalle (idTurno)
Set DataGrid1.DataSource = rsTurnoDetalle

End Sub

Private Sub Form_Unload(Cancel As Integer)
With rsTurnoDetalle
    If .State = 1 Then .Close
End With
End Sub

