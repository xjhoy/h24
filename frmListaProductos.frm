VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListaProductos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Productos Almacen"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14070
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBuscar 
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
      Left            =   7320
      TabIndex        =   14
      Top             =   1680
      Width           =   5295
   End
   Begin VB.CommandButton cmdPromo 
      Caption         =   "Ver promociónes"
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
      Left            =   14520
      TabIndex        =   13
      Top             =   2520
      Width           =   2535
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   17160
      Picture         =   "frmListaProductos.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   240
   End
   Begin MSDataGridLib.DataGrid dtgAlmacen 
      Height          =   7455
      Left            =   3120
      TabIndex        =   0
      Top             =   3720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13150
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
      ColumnCount     =   10
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
         Caption         =   "                     Descripción"
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
      BeginProperty Column08 
         DataField       =   ""
         Caption         =   "sadasd"
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
      BeginProperty Column09 
         DataField       =   "favProd"
         Caption         =   "Fav."
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
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   3270,047
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1019,906
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7800
      Picture         =   "frmListaProductos.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   240
   End
   Begin VB.CommandButton cmdAddProd 
      Caption         =   "Nuevo producto"
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
      Left            =   5160
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddUnid 
      Caption         =   "Agregar unidades"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdModProducto 
      Caption         =   "Modificar producto"
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
      Left            =   8280
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar producto"
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
      Left            =   11400
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   14040
      Picture         =   "frmListaProductos.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10920
      Picture         =   "frmListaProductos.frx":109E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4680
      Picture         =   "frmListaProductos.frx":1628
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   240
   End
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
      Left            =   16320
      Picture         =   "frmListaProductos.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   2880
      TabIndex        =   3
      Top             =   3600
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmListaProductos.frx":387C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   4800
      Picture         =   "frmListaProductos.frx":3898
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA DE PRODUCTOS EN EL ALMACEN"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   600
      Width           =   7575
   End
End
Attribute VB_Name = "frmListaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-- Lo que se carga cuando el formulario esta activo --
'======================================================
Private Sub Form_Activate()
    tbAlmacen
    Set dtgAlmacen.DataSource = rsUsuarios
End Sub

'-- Lo que se carga al abrir el formulario --
'============================================
Private Sub Form_Load()
    SSTab1.Caption = ""
    sIdS = "idProd ASC"
    tbAlmacen
End Sub

'-- Boton crear promoción --
'============================
Private Sub cmdPromo_Click()
    frmPromociones.Show
End Sub

'-- Boton agregar productos --
'================================
Private Sub cmdAddProd_Click()
    sOperacion = "A"
    frmNuevoProducto.Show
End Sub

'-- Boton agregar unidad --
'==========================
Private Sub cmdAddUnid_Click()
    
    Dim inpMod As String
    On Error GoTo error
    If Not selectId = "" Then
              inpMod = InputBox(dtgAlmacen.Columns(1).Text & vbNewLine & "Introducir unidades que va a añadir al producto" & vbNewLine & "Unidades actuales: " & dtgAlmacen.Columns(2).Text, Me.Caption)
    End If
    If inpMod = "" Then
              Exit Sub
    End If
    If IsNumeric(inpMod) = False Then
              MsgBox "Debe introducir solo números", vbCritical, Me.Caption
              Exit Sub
    End If
    With rsUsuarios
              .Requery
              .Find "idProd = '" & selectId & "'"
                        !uniProd = !uniProd + Val(inpMod)
              .UpdateBatch
              .Requery
    End With
error:
    Exit Sub

End Sub

'-- Boton modificar producto --
'===============================
Private Sub cmdModProducto_Click()
    Dim intboton As String
    If Not selectId = "" Then
        
        With rsUsuarios
            If .State = 1 Then
                If !promoProd = True Then
                    MsgBox "" & !nomProd & ", Esto es una promocion, para modificar este producto ve al listado de promociones", vbExclamation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
        
        With rsSort
            If .State = 1 Then
                If !promoProd = True Then
                    MsgBox "" & !nomProd & ", Esto es una promocion, para modificar este producto ve al listado de promociones", vbExclamation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    
        sOperacion = "M"
        frmNuevoProducto.Show
    Else
        intboton = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
    End If
End Sub

'-- Boton borrar producto --
'===========================
Private Sub cmdBorrar_Click()
    Dim intboton As Integer
    Dim intboton2 As Integer
    
    If rsUsuarios.BOF Or rsUsuarios.EOF Then Exit Sub
        intboton = MsgBox("¿Desea borrar el Producto?" & vbNewLine & "Código de barras: " & dtgAlmacen.Columns(0).Text & vbNewLine & "Descripción: " & dtgAlmacen.Columns(1).Text, vbQuestion Or vbYesNo, Me.Caption)
        If intboton = vbYes Then
            If (rsUsuarios.EOF Or rsUsuarios.BOF) Then
                Exit Sub
            Else
                With rsSort
                    If .State = 1 Then
                        .Delete
                        selectId = ""
                        Exit Sub
                    End If
                End With
                With rsBuscarProd
                    If .State = 1 Then
                        .Delete
                        
                        selectId = ""
                        Exit Sub
                    End If
                End With
                With rsUsuarios
                    .Delete
                    selectId = ""
                End With
            End If
        End If
    
    
End Sub

'-- Borrar con el boton suprimir en la tabla --
'==============================================
Private Sub dtgAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then cmdBorrar_Click
End Sub

'-- Organizar datagrid --
'=========================
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
    
    tbSortAlmacen
    
    rsSort.Sort = campo
    
    Set dtgAlmacen.DataSource = rsSort
    
    If rsSort.BOF = False Then
        rsSort.MoveFirst
    End If

End Sub

'-- Click en dataGrid --
'========================
Private Sub dtgAlmacen_Click()
    Dim x As String
    With rsBuscarProd
        If .State = 1 Then
            If .BOF Or .EOF Then
                Exit Sub
            End If
            .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
            selectId = dtgAlmacen.Columns(0).Text
            Exit Sub
        End If
    End With
    With rsSort
        If .State = 1 Then
            If .BOF Or .EOF Then
                Exit Sub
            End If
            .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
            selectId = dtgAlmacen.Columns(0).Text
            Exit Sub
            End If
    End With
    With rsUsuarios
        If .BOF Or .EOF Then Exit Sub
        .Find "idProd = '" & (dtgAlmacen.Columns(0).Text) & "'"
        x = dtgAlmacen.Row
        selectId = dtgAlmacen.Columns(0).Text
    End With
    
End Sub

'-- Texto buscar --
'===================
Private Sub txtBuscar_Change()
    Dim bProd As String
    
    bProd = txtBuscar.Text
    tbBuscarProd (bProd)
    
    Set dtgAlmacen.DataSource = rsBuscarProd
    
    With rsSort
        If .State = 1 Then
            .Close
        End If
    End With
    
End Sub

'-- Boton cerrar formulario --
'=============================
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

'-- formulario cerrado --
'=========================
Private Sub Form_Unload(Cancel As Integer)
    With rsSort
        If .State = 1 Then .Close
    End With
    With rsUsuarios
        If .State = 1 Then .Close
        bControlAcceso = False
    End With
    MDIfrmMadre.Picture1.Visible = True
End Sub

