VERSION 5.00
Begin VB.Form frmNuevoProducto 
   Caption         =   "Nuevo Producto"
   ClientHeight    =   5040
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
   ScaleHeight     =   5040
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   8040
      TabIndex        =   14
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   10800
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox txtnomProducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtstockmin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox txtUnidades 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox txtPrecioProducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox txtdescripcionProducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox txtidProducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblnomProducto 
      Caption         =   "Nombre del producto"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblstockmin 
      Caption         =   "Stock minimo"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblUnidades 
      Caption         =   "Unidades"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblprecioProducto 
      Caption         =   "Precio"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lbldescripcionProducto 
      Caption         =   "Descripción / Observaciones"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblidProducto 
      Caption         =   "Codigo del producto"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo producto"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmNuevoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
With rsUsuarios
    If .State = 1 Then .Close
    bControlAcceso = False
End With
End Sub

'====================================================
'Personaliza el color de los textbox y de los label  cuando tiene el focus y lo deja
'====================================================

Private Sub txtidProducto_GotFocus()
    lblidProducto.FontBold = True
    lblidProducto.ForeColor = &HFF
    
End Sub
Private Sub txtidProducto_LostFocus()
    lblidProducto.FontBold = False
    lblidProducto.ForeColor = &H80000012
    
End Sub
Private Sub txtnomProducto_GotFocus()
    lblnomProducto.FontBold = True
    lblnomProducto.ForeColor = &HFF
  
End Sub
Private Sub txtnomProducto_LostFocus()
    lblnomProducto.FontBold = False
    lblnomProducto.ForeColor = &H80000012
 
End Sub
Private Sub txtstockmin_GotFocus()
    lblstockmin.FontBold = True
    lblstockmin.ForeColor = &HFF

End Sub
Private Sub txtstockmin_LostFocus()
    lblstockmin.FontBold = False
    lblstockmin.ForeColor = &H80000012

End Sub
Private Sub txtUnidades_GotFocus()
    lblUnidades.FontBold = True
    lblUnidades.ForeColor = &HFF

End Sub
Private Sub txtUnidades_LostFocus()
    lblUnidades.FontBold = False
    lblUnidades.ForeColor = &H80000012

End Sub
Private Sub txtPrecioProducto_GotFocus()
    lblprecioProducto.FontBold = True
    lblprecioProducto.ForeColor = &HFF
 
End Sub
Private Sub txtPrecioProducto_LostFocus()
    lblprecioProducto.FontBold = False
    lblprecioProducto.ForeColor = &H80000012

End Sub
Private Sub txtdescripcionProducto_GotFocus()
    lbldescripcionProducto.FontBold = True
    lbldescripcionProducto.ForeColor = &HFF

End Sub
Private Sub txtdescripcionProducto_LostFocus()
    lbldescripcionProducto.FontBold = False
    lbldescripcionProducto.ForeColor = &H80000012

End Sub
'====================================
'Agrega usuarios al momento de dar Aceptar
'====================================

Private Sub cmdAceptar_Click()
 Dim intboton As Integer

'If ValidarCampos Then

    With rsUsuarios

    .MoveFirst
    '.Find "dniCliente= '" + txtdniCliente + "'"
    .Find "idProd= '" + txtidProducto + "'"
    If .EOF = False Then

    intboton = MsgBox("Este producto existe en la base de datos", vbCritical Or vbOKOnly, Me.Caption)
    txtidProducto.SetFocus
    

    Else

        .Requery
        .AddNew
            idProd = txtidProducto.Text
            !nomProducto = txtnomProducto.Text
            !stockmin = txtstockmin.Text
            !Unidades = txtUnidades.Text
            !precioProducto = txtPrecioProducto.Text
            !descripcionProducto = txtdescripcionProducto.Text
        .Update
        .Requery
        intboton = MsgBox("Producto agregado al almacen", vbInformation, Me.Caption)
        LimpiarCampos
    End If
    End With

'End If


End Sub



'==============================================
'txtnomCliente_Keypress solo dejara ingresar números y retroceso
' en el Else no volvemos el resto de teclas en null
'Autor: Jhoy
'==============================================

Sub txtyyMoto_Keypress(KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    
    End If

End Sub


'==========================================================
'ValidardniCliente
'Usa bCSA_ValidarCampodniCliente para comprobar si hay 9 caracteres en el txtdniCliente
'==========================================================
Public Function ValidardniCliente()
Dim intboton As Integer

    ValidardniCliente = True
  If ValidarCampodniCliente(txtdniCliente.Text) = False Then
      intboton = MsgBox("El Dni debe tener 9 caracteres", vbCritical, "#Racer - Error")
        ValidardniCliente = False
  
 End If
    
End Function



'=================================================================
'ValidarprovCliente
'Usa bCSA_ValidarCampoLetras para comprobar si contiene letras y son mas de 2 en el txtprovCliente
'=================================================================
Public Function ValidarprovCliente()
Dim intboton As Integer
Dim intboton2 As Integer
    ValidarprovCliente = True

    If ValidarCampoLetras(txtprovCliente.Text) = False Then
       intboton = MsgBox("Debe contener al menos una letra", vbCritical, "#Racer - Error")
        
        ValidarprovCliente = False
    Else
      If ValidarCamponomCliente(txtprovCliente.Text) = False Then
       intboton2 = MsgBox("Debe de introducir al menos 2 caracteres", vbCritical, "#Racer - Error")
        ValidarprovCliente = False
       End If
    End If
    
End Function
'=================================================================
'ValidarprovCliente
'Usa bCSA_ValidarCampoLetras para comprobar si contiene letras y son mas de 2 en el txtprovCliente
'=================================================================
Public Function ValidarlocCliente()
Dim intboton As Integer
Dim intboton2 As Integer
    ValidarlocCliente = True

    If ValidarCampoLetras(txtlocCliente.Text) = False Then
      intboton = MsgBox("Debe contener al menos una letra", vbCritical, "#Racer - Error")
        
        ValidarlocCliente = False
    Else
      If ValidarCamponomCliente(txtlocCliente.Text) = False Then
        intboton2 = MsgBox("Debe de introducir al menos 2 caracteres", vbCritical, "#Racer - Error")
        ValidarlocCliente = False
       End If
    End If
    
End Function

'=========================================================
'ValidartlfCliente
'Usa bCSA_ValidarCampotlfCliente para comprobar si solo hay números en el txttlfCliente
'=========================================================

Public Function ValidarcodMoto()
Dim intboton2 As Integer
Dim intboton3 As Integer

    ValidartlfCliente = True
 If ValidarCampoNumerico(txtcodMoto.Text) = False Then
        intboton2 = MsgBox("Debe introducir un dato númerico en ese campo", vbCritical, "#Racer - Error")
        ValidarcodMoto = False
 Else
    If txtcodMoto.Text < 0 Then
          intboton3 = MsgBox("Debe introducir un valor mayor que cero en este campo", vbCritical, "#Racer - Error")
            ValidarcodMoto = False
    End If
 End If
    
End Function

'==================
'Valida campos
'==================

Function ValidarCampos()

    ValidarCampos = True
   If ValidardniCliente = False Then
        txtdniCliente.SetFocus
        ValidarCampos = False
  End If
      
      
End Function


Private Sub Form_Load()
tbAlmacen
End Sub

Public Sub LimpiarCampos()
      txtidProducto.Text = ""
      txtnomProducto.Text = ""
      txtstockmin.Text = ""
      txtUnidades.Text = ""
      txtPrecioProducto.Text = ""
      txtdescripcionProducto.Text = ""
End Sub
