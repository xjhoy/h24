VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProvDetalle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Proveedor"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12600
      Picture         =   "frmProvDetalle.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   6240
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12600
      Picture         =   "frmProvDetalle.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   7440
      Width           =   240
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12600
      Picture         =   "frmProvDetalle.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   8640
      Width           =   240
   End
   Begin VB.CommandButton cmdProdAdd 
      Caption         =   " Agregar producto"
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
      Left            =   12960
      TabIndex        =   23
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdModProd 
      Caption         =   " Modificar producto"
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
      Left            =   12960
      TabIndex        =   22
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdBorrarProd 
      Caption         =   "Eliminar producto"
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
      Left            =   12960
      TabIndex        =   21
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox txtnomProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtProvinciaProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox txtcifProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtlocProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txtcpProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtdirProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox txttlfProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtcontProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   4
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox txtemailProv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   3
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   15960
      Picture         =   "frmProvDetalle.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10800
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dtgProv 
      Height          =   6735
      Left            =   1800
      TabIndex        =   2
      Top             =   5880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11880
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
         DataField       =   "idProdProv"
         Caption         =   "ID"
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
      BeginProperty Column01 
         DataField       =   "nomProd"
         Caption         =   "Producto"
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
         DataField       =   "precioProd"
         Caption         =   "Precio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 €"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "notaProd"
         Caption         =   "Nota"
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
         DataField       =   "cifProv"
         Caption         =   "CIF"
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
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   3674,835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4605,166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   1680
      TabIndex        =   27
      Top             =   5760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmProvDetalle.frx":2D68
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Razón social"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CIF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Código postal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Persona de contacto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   13
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   12
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "proveedor"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   3840
      Picture         =   "frmProvDetalle.frx":2D84
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "frmProvDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBorrarProd_Click()
Dim intboton As Integer
Dim intboton2 As Integer
On Error GoTo error

If Not SelectProd = "" Then
          intboton = MsgBox("¿Desea borrar el producto?" & vbNewLine & "Nombre del producto: " & dtgProv.Columns(1).Text & vbNewLine & "Precio: " & dtgProv.Columns(2).Text & "€", vbQuestion Or vbYesNo, Me.Caption)
          If intboton = vbYes Then

                    If (rsProd.EOF Or rsProd.BOF) Then
                              Exit Sub
                    Else
                              tbProdProv
                              With rsUsuarios
                                        .Requery
                                        .Find "idProdProv = '" & (SelectProd) & "'"
                                                  idProdProv = dtgProv.Columns(0).Text
                                                  !nomProd = dtgProv.Columns(1).Text
                                                  !precioProd = dtgProv.Columns(2).Text
                                                  !notaProd = dtgProv.Columns(3).Text
                                        .Delete
                                        .MoveFirst
                                        .Close
                              End With
                              rsProd.Requery
                              SelectProd = ""
                    End If
          End If
Else
error:
intboton2 = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
End If

End Sub

Private Sub cmdCerrar_Click()
bControlAcceso = False
frmListProv.Show
Unload Me
End Sub

Private Sub cmdModProd_Click()
sOperacion = "B"

If Not SelectProd = "" Then
          frmAddProdProv.Show
Else
          MsgBox "No ha seleccionado ningun producto para modificar", vbExclamation, Me.Caption
End If


End Sub

Private Sub cmdProdAdd_Click()
sOperacion = "A"
frmAddProdProv.Show
End Sub

Private Sub dtgProv_Click()
With rsProd
          If .BOF Or .EOF Then Exit Sub
          .Find "idProdProv = '" & (dtgProv.Columns(0).Text) & "'"
          SelectProd = dtgProv.Columns(0).Text
End With
End Sub

Private Sub Form_Activate()

Label1.Caption = txtnomProv.Text
With rsProd
          If .State = 1 Then
                    .Close
          End If
          .Open "SELECT * FROM tbProdProv WHERE cifProv ='" & Trim(txtcifProv.Text) & "'", base, adOpenStatic, adLockReadOnly
End With

Set dtgProv.DataSource = rsProd
dtgProv.Columns(2).Caption = "Precio/€"
End Sub

Private Sub Form_Load()
bControlAcceso = True
SSTab1.Caption = ""
With rsUsuarios
          .Requery
          .Find "cifProv = '" & (SelectDetalleProv) & "'"
                    txtcifProv.Text = !cifProv
                    txtnomProv.Text = !nomProv
                    txtProvinciaProv.Text = !provinciaProv
                    txtlocProv.Text = !locProv
                    txtdirProv.Text = !dirprov
                    txtcpProv.Text = !cpProv
                    txttlfProv.Text = !tlfProv
                    txtemailProv.Text = !emailProv
                    txtcontProv.Text = !contProv
                    txttlfProv.Text = !tlfProv
End With
End Sub

'Ordenar dtg
'=============
Private Sub dtgProv_HeadClick(ByVal ColIndex As Integer)
campo = ColIndex
If campo = "1" Then
          If sIdS = "idProd ASC" Then
                    campo = "nomProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "nomProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "2" Then
          If sIdS = "idProd ASC" Then
                    campo = "precioProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "precioProd DESC"
                    sIdS = "idProd ASC"
          End If
Else
          Exit Sub
End If

rsProd.Sort = campo
Set dtgProv.DataSource = rsProd
If rsProd.BOF = False Then rsProd.MoveFirst
          SelectProd = ""
End Sub

