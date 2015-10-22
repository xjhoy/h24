VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaProductos 
   Caption         =   "Productos Almacen"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8070
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
   ScaleHeight     =   4170
   ScaleWidth      =   8070
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar producto"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdModProducto 
      Caption         =   "Modificar prodicto"
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddUnid 
      Caption         =   "Agregar unidades"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   17520
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dtgAlmacen 
      Height          =   6495
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   11456
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblidProducto 
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Height          =   1335
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmListaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAddUnid_Click()
Dim inpMod As String
inpMod = InputBox("Introducir unidades que va a añadir al producto", Me.Caption)
    With rsUsuarios
        .Requery
        .Find "idProducto = '" & (lblidProducto.Caption) & "'"
          !Unidades = !Unidades + Val(inpMod)
        .UpdateBatch
        .Requery
    End With


End Sub

Private Sub cmdBorrar_Click()
Dim intboton As Integer
Dim intboton2 As Integer

If Not lblidProducto = "" Then
intboton = MsgBox("¿Desea borar el producto?" & vbNewLine & "Código: " & dtgAlmacen.Columns(0).Text & vbNewLine & "Nombre del producto: " & dtgAlmacen.Columns(1).Text, vbQuestion Or vbYesNo, Me.Caption)
If intboton = vbYes Then

    If (rsUsuarios.EOF Or rsUsuarios.BOF) Then
    Exit Sub
    Else
    With rsUsuarios
    .Requery
    .Find "idProducto = '" & (lblidProducto.Caption) & "'"
            idProd = dtgAlmacen.Columns(0).Text
            !nomProducto = dtgAlmacen.Columns(1).Text
            !stockmin = dtgAlmacen.Columns(2).Text
            !Unidades = dtgAlmacen.Columns(3).Text
            !precioProducto = dtgAlmacen.Columns(4).Text
            !descripcionProducto = dtgAlmacen.Columns(5).Text
    .Delete
    .MoveFirst
    '.Requery

    End With
    End If
End If
Else
  intboton2 = MsgBox("Seleccione un cliente", vbExclamation, "#Racer")
End If

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdModProducto_Click()
Dim intboton As String

If Not lblidProducto.Caption = "" Then
frmModProducto.Show
Else
   intboton = MsgBox("Seleccione un producto", vbExclamation, Me.Caption)
End If
End Sub

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub dtgAlmacen_Click()
With rsUsuarios
    If .BOF Or .EOF Then Exit Sub
    .Find "idProducto = '" & (dtgAlmacen.Columns(0).Text) & "'"
    lblidProducto = idProd
    
    
End With

End Sub

Private Sub Form_Load()
tbAlmacen
Set dtgAlmacen.DataSource = rsUsuarios
End Sub

Private Sub Form_Unload(Cancel As Integer)
With rsUsuarios
    If .State = 1 Then .Close
    bControlAcceso = False
End With
End Sub

