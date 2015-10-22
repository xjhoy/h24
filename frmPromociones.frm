VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPromociones 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Listado de Promociones"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   16200
      Picture         =   "frmPromociones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15600
      Picture         =   "frmPromociones.frx":1CCA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   240
   End
   Begin VB.CommandButton cmdAddPromo 
      Caption         =   "Crear Promoción"
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
      Left            =   15960
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdModPromo 
      Caption         =   "Modificar Promoción"
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
      Left            =   15960
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdBorrarPromo 
      Caption         =   "Borrar promoción"
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
      Left            =   15960
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15600
      Picture         =   "frmPromociones.frx":2254
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   240
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15600
      Picture         =   "frmPromociones.frx":27DE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   240
   End
   Begin MSDataGridLib.DataGrid dtgAlmacen 
      Height          =   8775
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   15478
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "idProd"
         Caption         =   "Código"
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
         Caption         =   "                Descripción"
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
         DataField       =   "uniProd"
         Caption         =   "Unidades"
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
         DataField       =   "netoProd"
         Caption         =   "Neto"
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
      BeginProperty Column04 
         DataField       =   "ivaProd"
         Caption         =   "IVA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 ""%"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "pvpProd"
         Caption         =   "PVP"
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
      BeginProperty Column06 
         DataField       =   "idProdManual"
         Caption         =   "Cód.M"
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
      BeginProperty Column07 
         DataField       =   "favProd"
         Caption         =   "Fav."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   ""
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "promoProd"
         Caption         =   "Promo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   ""
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2039,811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   810,142
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPromociones.frx":2D68
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA DE PROMOCIONES "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   5160
      Picture         =   "frmPromociones.frx":2D84
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "frmPromociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddProd_Click()
sOperacion = "A"
frmNuevoProducto.Show
End Sub

Private Sub cmdAddPromo_Click()
sOperacion = "A"
frmAddPromo.Show
End Sub

Private Sub cmdBorrarPromo_Click()
Dim intboton As Integer
Dim intboton2 As Integer
'On Error GoTo error

'If Not selectId = "" Then
If rsPromo.BOF Or rsPromo.EOF Then Exit Sub
          intboton = MsgBox("¿Desea borrar el Producto?" & vbNewLine & "Código de barras: " & dtgAlmacen.Columns(0).Text & vbNewLine & "Descripción: " & dtgAlmacen.Columns(1).Text, vbQuestion Or vbYesNo, Me.Caption)
          If intboton = vbYes Then
                    If (rsPromo.EOF Or rsPromo.BOF) Then
                              Exit Sub
                    Else
                    'rsPromo.Requery
                              'base.Execute ("DELETE * FROM tbpromo Where idProd ='" & selectId & "'")
                              With rsSortPromo
                                        If .State = 1 Then
                                                  .Delete
                                                  .MoveNext
                                                  '.Requery
                                        End If
                              End With
                              With rsPromo
                                        '.Requery
                                        .Delete
                                        If .BOF = False Then .MoveNext
                                        selectId = ""
                              End With
                    End If
          End If
'Else
'error:
'  intboton2 = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
'End If

End Sub

Private Sub cmdModPromo_Click()
Dim intboton As String
sOperacion = "M"
With rsPromo
          If .State = 1 And !promoProd = True Then
                    If Not selectId = "" Then

                                        tbAddProd
                                        With rsPromo
                                                  .Requery
                                                  .Find "idProd = '" & selectId & "'"
                                                  If !promoProd = True Then
                                                            frmAddPromo.Show
                                                  Else
                                                            intboton = MsgBox("Seleccione una promoción", vbExclamation, Me.Caption)
                                                  End If
                                        End With
                    End If
          Else
                    If rsSortPromo.State = 1 Then
                              With rsSortPromo
                                        If !promoProd = True Then
                                                  If Not selectId = "" Then
                                                            sOperacion = "M"
                                                            tbAddProd
                                                            With rsPromo
                                                                      .Requery
                                                                      .Find "idProd = '" & selectId & "'"
                                                                      If !promoProd = True Then
                                                                                frmAddPromo.Show
                                                                      Else
                                                                                intboton = MsgBox("Seleccione una promoción", vbExclamation, Me.Caption)
                                                                      End If
                                                            End With
                                                  End If
                                        Else
                                                  intboton = MsgBox("Seleccione una promoción", vbExclamation, Me.Caption)
                                        End If
                              End With
                    'Else
                     '         MsgBox "Debe seleccionar un producto en oferta", vbExclamation, Me.Caption
                    End If
          End If
End With
End Sub

Private Sub dtgAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then cmdBorrar_Click ' Else Exit Sub
End Sub

Private Sub Form_Activate()
 tbPromo
 With rsPromo
          .Filter = "promoProd = 'True'"
 End With
 
Set dtgAlmacen.DataSource = rsPromo

End Sub

Private Sub Form_Load()
SSTab1.Caption = ""
sIdS = "idProd ASC"
tbPromo
End Sub

Private Sub cmdAddUnid_Click()
Dim inpMod As String
On Error GoTo error
If Not selectId = "" Then
          inpMod = InputBox("Introducir unidades que va a añadir al producto" & vbNewLine & "Unidades actuales: " & dtgAlmacen.Columns(2).Text, Me.Caption)
End If
If inpMod = "" Then
          Exit Sub
End If
If IsNumeric(inpMod) = False Then
          MsgBox "Debe introducir solo números", vbCritical, Me.Caption
          Exit Sub
End If
With rsPromo
          .Requery
          .Find "idProd = '" & selectId & "'"
                    !uniProd = !uniProd + Val(inpMod)
          .UpdateBatch
          .Requery
End With
error:
Exit Sub

End Sub

Private Sub cmdBorrar_Click()
Dim intboton As Integer
Dim intboton2 As Integer
'On Error GoTo error

'If Not selectId = "" Then
If rsPromo.BOF Or rsPromo.EOF Then Exit Sub
          intboton = MsgBox("¿Desea borrar el Producto?" & vbNewLine & "Código de barras: " & dtgAlmacen.Columns(0).Text & vbNewLine & "Descripción: " & dtgAlmacen.Columns(1).Text, vbQuestion Or vbYesNo, Me.Caption)
          If intboton = vbYes Then
                    If (rsPromo.EOF Or rsPromo.BOF) Then
                              Exit Sub
                    Else
                    'rsPromo.Requery
                              'base.Execute ("DELETE * FROM tbpromo Where idProd ='" & selectId & "'")
                              With rsSortPromo
                                        If .State = 1 Then
                                                  .Delete
                                                  '.MoveNext
                                                  selectId = ""
                                                  Exit Sub
                                                  '
                                                  '.Requery
                                        End If
                              End With
                              With rsPromo
                                        '.Requery
                                        .Delete
                                        'If .BOF = False Then .MoveNext
                                        selectId = ""
                              End With
                    End If
          End If
'Else
'error:
'  intboton2 = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
'End If


End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdModProducto_Click()
Dim intboton As String

If Not selectId = "" Then
          sOperacion = "M"
          frmNuevoProducto.Show
Else
          intboton = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
End If
End Sub



Private Sub dtgAlmacen_Click()
Dim x As String
With rsSortPromo
          If .State = 1 Then
                    If .BOF Or .EOF Then
                              Exit Sub
                    End If
                    '.Filter = "promoProd = 'True'"
                    '.Requery
                    .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
                    selectId = dtgAlmacen.Columns(0).Text
                    Exit Sub
          End If
End With
With rsPromo
          If .BOF Or .EOF Then Exit Sub
          '.Filter = "promoProd = 'True'"
          '.Requery
          .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
          x = dtgAlmacen.Row
          selectId = dtgAlmacen.Columns(0).Text
End With

End Sub

'Organizar dtg
'=========
Private Sub dtgAlmacen_HeadClick(ByVal ColIndex As Integer)
campo = ColIndex

If campo = "0" Then
          If sIdS = "idProd ASC" Then
                    campo = "idProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "idProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "1" Then
          If sIdS = "idProd ASC" Then
                    campo = "nomProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "nomProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "2" Then
          If sIdS = "idProd ASC" Then
                    campo = "uniProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "uniProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "3" Then
          If sIdS = "idProd ASC" Then
                    campo = "netoProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "netoProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "4" Then
          If sIdS = "idProd ASC" Then
                    campo = "ivaProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "ivaProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "5" Then
          If sIdS = "idProd ASC" Then
                    campo = "pvpProd ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "pvpProd DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "6" Then
          If sIdS = "idProd ASC" Then
                    campo = "idProdManual ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "idProdManual DESC"
                    sIdS = "idProd ASC"
          End If
Else
          Exit Sub
End If
tbSortPromo
With rsSortPromo
          .Filter = "promoProd = 'True'"
End With
rsSortPromo.Sort = campo
Set dtgAlmacen.DataSource = rsSortPromo
If rsSortPromo.BOF = False Then
          rsSortPromo.MoveFirst
End If

'With rsSortPromo
'    If .BOF Or .EOF Then Exit Sub
'    .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
'    selectId = dtgAlmacen.Columns(0).Text
'End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
With rsSortPromo
    If .State = 1 Then .Close
    'bControlAcceso = False
End With
With rsPromo
    If .State = 1 Then .Close
    bControlAcceso = False
End With
frmListaProductos.Show
End Sub

