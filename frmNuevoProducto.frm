VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNuevoProducto 
   BorderStyle     =   0  'None
   Caption         =   "Nuevo Producto"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
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
   ScaleHeight     =   7230
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFav 
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5400
      TabIndex        =   21
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkFav 
      Caption         =   "Bebidas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   4080
      TabIndex        =   20
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkFav 
      Caption         =   "Promociones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkFav 
      Caption         =   "Bolleria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmNuevoProducto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblcodBarras"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblivaProd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblcodM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbluniProd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblnomProd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblpvpProd"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFavoritos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtnomProd"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtuniProd"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtpvpProd"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtivaProd"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtidM"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtidProd"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAceptar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdCerrar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkFav(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.CheckBox chkFav 
         Caption         =   "Pan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   3960
         TabIndex        =   7
         Top             =   6000
         Width           =   2175
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   735
         Left            =   1440
         TabIndex        =   6
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txtidProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   0
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtidM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox txtivaProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtpvpProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtuniProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtnomProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   7335
      End
      Begin VB.Label lblFavoritos 
         Caption         =   "Favoritos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   5
         X1              =   1320
         X2              =   6000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblpvpProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio/€"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblnomProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lbluniProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblcodM 
         BackStyle       =   0  'Transparent
         Caption         =   "Código manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblivaProd 
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblcodBarras 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Barras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmNuevoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCerrar_Click()
'rsUsuarios.Requery
Unload Me
End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIfrmMadre.Enabled = True
'selectId = ""
End Sub

Private Sub Label2_Click()

End Sub

'====================================================
'Personaliza el color de los textbox y de los label  cuando tiene el focus y lo deja
'====================================================

Private Sub txtidProd_GotFocus()
    lblcodBarras.FontBold = True
    lblcodBarras.ForeColor = &HFF
    
End Sub
Private Sub txtidProd_LostFocus()
    lblcodBarras.FontBold = False
    lblcodBarras.ForeColor = &H80000012
    
End Sub
Private Sub txtnomProd_GotFocus()
    lblnomProd.FontBold = True
    lblnomProd.ForeColor = &HFF
  
End Sub
Private Sub txtnomProd_Lostfocus()
    lblnomProd.FontBold = False
    lblnomProd.ForeColor = &H80000012
 
End Sub
Private Sub txtuniProd_GotFocus()
    lbluniProd.FontBold = True
    lbluniProd.ForeColor = &HFF

End Sub
Private Sub txtuniProd_LostFocus()
    lbluniProd.FontBold = False
    lbluniProd.ForeColor = &H80000012

End Sub
Private Sub txtpvpProd_GotFocus()
    lblpvpProd.FontBold = True
    lblpvpProd.ForeColor = &HFF

End Sub
Private Sub txtpvpProd_LostFocus()
    lblpvpProd.FontBold = False
    lblpvpProd.ForeColor = &H80000012

End Sub
Private Sub txtivaProd_GotFocus()
    lblivaProd.FontBold = True
    lblivaProd.ForeColor = &HFF
 
End Sub
Private Sub txtivaProd_LostFocus()
    lblivaProd.FontBold = False
    lblivaProd.ForeColor = &H80000012

End Sub
Private Sub txtidM_GotFocus()
    lblcodM.FontBold = True
    lblcodM.ForeColor = &HFF

End Sub
Private Sub txtidM_LostFocus()
    lblcodM.FontBold = False
    lblcodM.ForeColor = &H80000012

End Sub

'==========================
'Aceptar con ENTER en cualquier TextBox
'==========================

Private Sub txtidProd_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And txtnomProd.Text <> "" Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtnomProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtuniProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 46
    If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    
    End If
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtpvpProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then cmdAceptar_Click ' Else Exit Sub
    If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    
    End If

End Sub
Private Sub txtivaProd_KeyPress(KeyAscii As Integer)
'If KeyAscii = 44 Then KeyAscii = 46
If KeyAscii = 13 Then cmdAceptar_Click 'Else Exit Sub
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    
    End If

End Sub
Private Sub txtidM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

'====================================
'Agrega usuarios al momento de dar Aceptar
'====================================

Private Sub cmdAceptar_Click()

 Dim intboton As Integer
 If txtidProd.Text = "" And txtidM.Text = "" Then
 MsgBox "El campo código de barras es obligatorio, si es un producto que no posee codigo de barras debe ingresar un código manual", vbCritical, Me.Caption
 txtidProd.SetFocus
 Exit Sub
 Else
    If txtidProd.Text = "" Then
    txtidProd.Text = txtidM.Text
    End If
End If
Select Case sOperacion
Case "A"
If validarIdProd = False Then Exit Sub
If validaridProdM = False Then Exit Sub
        
With rsAlmacen
        .Requery
        .AddNew
            
            !idProd = txtidProd.Text
            !nomProd = txtnomProd.Text
            
            If txtuniProd.Text = "" Then
                    txtuniProd.Text = "0"
            End If
            
            !uniProd = txtuniProd.Text
            
            If txtivaProd.Text = "" Then
                    txtivaProd.Text = "0"
            End If
            
            !ivaProd = CDec(txtivaProd.Text)
            
            If txtpvpProd.Text = "" Then
                    txtpvpProd.Text = "0"
            End If
            !pvpProd = CDec(txtpvpProd.Text)
            !idProdManual = txtidM.Text
            Dim Rneto As Double
            Rneto = CDec(txtpvpProd.Text) - ((CDec(txtpvpProd.Text) * CDec(txtivaProd.Text)) / 100)
            Rneto = Format(Rneto, "###,##0.00")
            !netoProd = Rneto
            !favA = chkFav(0)
            !favB = chkFav(1)
            !favC = chkFav(2)
            !favD = chkFav(3)
            !favE = chkFav(4)
            If (!favA = True) Or (!favB = True) Or (!favC = True) Or (!favD = True) Or (!favE = True) Then
                !favProd = True
            End If
        .Update
        .Requery
 
 End With
 With rsSort
        If .State = 1 Then
                     .Requery
        End If
 End With

rsUsuarios.Requery
MsgBox "Producto creado", vbInformation, Me.Caption

intboton = MsgBox("¿Desea crear otro producto?", vbQuestion Or vbYesNo, Me.Caption)

If intboton = vbYes Then
          LimpiarCampos
          txtidProd.SetFocus
Else
          selectId = ""
          Unload Me
End If

Case "M"
If sModidProd <> txtidProd.Text Then
    If validarIdProd = False Then Exit Sub
End If
If sModidProdM <> txtidM.Text Then
    If validaridProdM = False Then Exit Sub
End If

    With rsAlmacen
        .Requery
        .Find "idProd = '" & (selectId) & "'"
            !idProd = txtidProd.Text
             !idProdManual = txtidM.Text
            !nomProd = txtnomProd.Text
            !uniProd = txtuniProd.Text
            !ivaProd = CDec(txtivaProd.Text)
            !pvpProd = CDec(txtpvpProd.Text)
            
            Dim RnetoM As Double
            RnetoM = CDec(txtpvpProd.Text) - ((CDec(txtpvpProd.Text) * CDec(txtivaProd.Text)) / 100)
            RnetoM = Format(RnetoM, "###,##0.00")
            !netoProd = RnetoM
            !favA = chkFav(0)
            !favB = chkFav(1)
            !favC = chkFav(2)
            !favD = chkFav(3)
            !favE = chkFav(4)
            If (!favA = True) Or (!favB = True) Or (!favC = True) Or (!favD = True) Or (!favE = True) Then
                !favProd = True
            End If
        .UpdateBatch
        'base.Execute ("UPDATE tbAlmacen Set netoProd = pvpProd-((pvpProd*ivaProd)/100) WHERE idprod ='" & txtidProd & "'")
        With rsSort
        If .State = 1 Then
        .Requery
        End If
        End With
        rsUsuarios.Requery
        Unload Me
    End With
End Select
End Sub


Public Function validarIdProd()
validarIdProd = True
If rsAlmacen.BOF = False Then
    rsAlmacen.MoveFirst
End If
rsAlmacen.Find "idProd= '" + txtidProd + "'"
If rsAlmacen.EOF = False Then
    intboton = MsgBox("Este código de barras ya existe en la base de datos." & vbNewLine & vbNewLine & "Recuerda que sí esta utilizando solamente un código manual, este no puede ser igual que un código de barras ya existente.", vbCritical, Me.Caption)
    txtidProd.SetFocus
    validarIdProd = False
End If
End Function
Public Function validaridProdM()
validaridProdM = True
If rsAlmacen.BOF = False Then
    rsAlmacen.MoveFirst
End If
If txtidM.Text <> "" Then
    rsAlmacen.Find "idProdManual = '" + txtidM + "'"
        If rsAlmacen.EOF = False Then
        intboton = MsgBox("Este código manual ya existe en la base de datos", vbCritical, Me.Caption)
        txtidM.SetFocus
        validaridProdM = False
        End If
End If
End Function




Private Sub Form_Load()

MDIfrmMadre.Enabled = False
SSTab1.Caption = ""
tbAddProd

Select Case sOperacion

Case "A"
LimpiarCampos

Case "M"
Dim VfavA, VfavB, VfavC, VfavD, VfavE As String

Label1.Caption = "Modificar producto"
With rsAlmacen
.Requery
.Find "idProd = '" & (selectId) & "'"
            txtidProd.Text = !idProd
            txtnomProd.Text = !nomProd
            txtuniProd.Text = !uniProd
            txtivaProd.Text = !ivaProd
            txtpvpProd.Text = !pvpProd
            txtidM.Text = !idProdManual
            
            If !favA = True Then
                VfavA = 1
            Else
                VfavA = 0
            End If
            If !favB = True Then
                VfavB = 1
            Else
                VfavB = 0
            End If
            If !favC = True Then
                VfavC = 1
            Else
                VfavC = 0
            End If
            If !favD = True Then
                VfavD = 1
            Else
                VfavD = 0
            End If
            If !favE = True Then
                VfavE = 1
            Else
                VfavE = 0
            End If
            chkFav(0) = VfavA
            chkFav(1) = VfavB
            chkFav(2) = VfavC
            chkFav(3) = VfavD
            chkFav(4) = VfavE
End With
sModidProd = txtidProd.Text
sModidProdM = txtidM.Text
End Select
End Sub


Public Sub LimpiarCampos()

txtidProd.Text = ""
txtnomProd.Text = ""
txtuniProd.Text = ""
txtivaProd.Text = ""
txtpvpProd.Text = ""
txtidM.Text = ""
chkFav(0) = 0
chkFav(1) = 0
chkFav(2) = 0
chkFav(3) = 0
chkFav(4) = 0

tbCountFav2
With rsCountFav2

    If !favA = -24 Then
        chkFav(0).Enabled = False
    End If
    If !favB = -24 Then
        chkFav(1).Enabled = False
    End If
    If !favC = -24 Then
        chkFav(2).Enabled = False
    End If
    If !favD = -24 Then
        chkFav(3).Enabled = False
    End If
    If !favE = -24 Then
        chkFav(4).Enabled = False
    End If
    
End With

End Sub
