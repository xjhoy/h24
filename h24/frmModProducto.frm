VERSION 5.00
Begin VB.Form frmModProducto 
   Caption         =   "Modificar - Producto"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
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
   ScaleHeight     =   6540
   ScaleWidth      =   10185
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   1800
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtdescripcionProducto 
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox txtPrecioProducto 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtUnidades 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtstockmin 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtnomProducto 
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label lblidProducto2 
      Caption         =   "Codigo del producto"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblidProducto 
      Caption         =   "Codigo del producto"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lbldescripcionProducto 
      Caption         =   "Descripción / Observaciones"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblprecioProducto 
      Caption         =   "Precio"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblUnidades 
      Caption         =   "Unidades"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblstockmin 
      Caption         =   "Stock minimo"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblnomProducto 
      Caption         =   "Nombre del producto"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "frmModProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

    With rsUsuarios
        .Requery
        .Find "idProducto = '" & (lblidProducto2.Caption) & "'"
            !nomProducto = txtnomProducto.Text
            !stockmin = txtstockmin.Text
            !Unidades = txtUnidades.Text
            !precioProducto = txtPrecioProducto.Text
            !descripcionProducto = txtdescripcionProducto.Text
        .UpdateBatch
        .Requery
         Unload Me
    End With

End Sub

Private Sub Form_Load()
With rsUsuarios
.Requery
.Find "idProducto = '" & (frmListaProductos.lblidProducto.Caption) & "'"
            lblidProducto2.Caption = idProd
            txtnomProducto.Text = !nomProducto
            txtstockmin.Text = !stockmin
            txtUnidades.Text = !Unidades
            txtPrecioProducto.Text = !precioProducto
            txtdescripcionProducto.Text = !descripcionProducto
End With
End Sub
