VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddPromo 
   BorderStyle     =   0  'None
   Caption         =   "Promocion"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAddPromo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFavoritos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblpvpProd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblnomProd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbluniProd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblcodM"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblivaProd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblcodBarras"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPromo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLleva"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPaga"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCerrar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAceptar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtidProd"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtidM"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtivaProd"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtpvpProd"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtuniProd"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtnomProd"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkFav(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtLleva"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPaga"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      Begin VB.TextBox txtPaga 
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
         Left            =   4200
         TabIndex        =   7
         Top             =   4920
         Width           =   495
      End
      Begin VB.TextBox txtLleva 
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
         Left            =   3360
         TabIndex        =   6
         Top             =   4920
         Width           =   495
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
         Left            =   2760
         TabIndex        =   8
         Top             =   5520
         Width           =   1695
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
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   735
         Left            =   1440
         TabIndex        =   9
         Top             =   6360
         Width           =   2175
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   3960
         TabIndex        =   10
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label lblPaga 
         BackStyle       =   0  'Transparent
         Caption         =   "Paga"
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
         Left            =   4800
         TabIndex        =   23
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblLleva 
         BackStyle       =   0  'Transparent
         Caption         =   "Lleva"
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
         Left            =   2760
         TabIndex        =   22
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   3960
         TabIndex        =   21
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblPromo 
         BackStyle       =   0  'Transparent
         Caption         =   "Promoción:"
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
         TabIndex        =   20
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva promoción"
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
         TabIndex        =   13
         Top             =   0
         Width           =   7335
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
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
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
         TabIndex        =   18
         Top             =   3720
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
         TabIndex        =   17
         Top             =   4320
         Width           =   1815
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
         TabIndex        =   16
         Top             =   2520
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
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
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
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   5
         X1              =   1320
         X2              =   6000
         Y1              =   720
         Y2              =   720
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
         TabIndex        =   12
         Top             =   5640
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAddPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkFav1_Click(Index As Integer)

End Sub

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
            
            If txtLleva.Text = "" Then
            txtLleva.Text = "0"
            End If
            !llevaProd = txtLleva.Text
            
            If txtPaga.Text = "" Then
            txtPaga.Text = "0"
            End If
            !pagaProd = txtPaga.Text
            
            !promoProd = True
            !favC = chkFav(2)
        .Update
        .Requery
        End With
        With rsSortPromo
        If .State = 1 Then
                    .Requery
        Else
                    rsPromo.Requery
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
          
          If txtuniProd.Text = "" Then
                    txtuniProd.Text = 0
          End If
          !uniProd = txtuniProd.Text
          
          If txtivaProd.Text = "" Then
                    txtivaProd = 0
          End If
          !ivaProd = CDec(txtivaProd.Text)
          If txtpvpProd.Text = "" Then
                    txtpvpProd.Text = 0
          End If
          !pvpProd = CDec(txtpvpProd.Text)
          !favC = chkFav(2)
          Dim RnetoM As Double
          RnetoM = CDec(txtpvpProd.Text) - ((CDec(txtpvpProd.Text) * CDec(txtivaProd.Text)) / 100)
          RnetoM = Format(RnetoM, "###,##0.00")
          !netoProd = RnetoM
          
          If txtLleva.Text = "" Then
                    txtLleva.Text = 0
          End If
          !llevaProd = txtLleva.Text
          
          If txtPaga.Text = "" Then
                    txtPaga.Text = 0
          End If
          !pagaProd = txtPaga.Text
          
    .UpdateBatch
          With rsSortPromo
                    If .State = 1 Then
                                .Requery
                    ElseIf rsAlmacen.State = 1 Then
                              rsAlmacen.Requery
                    ElseIf rsUsuarios.State = 1 Then
                              rsUsuarios.Requery
                    Else
                                rsPromo.Requery
                    End If
          End With

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
tbAddProd
MDIfrmMadre.Enabled = False
SSTab1.Caption = ""

Select Case sOperacion
Case "A"
LimpiarCampos
Case "M"
Dim Vfav As String
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
            txtLleva.Text = !llevaProd
            txtPaga.Text = !pagaProd
            If !favC = True Then
                Vfav = 1
            Else
                Vfav = 0
            End If
            
            chkFav(2) = Vfav
End With
sModidProd = txtidProd.Text
sModidProdM = txtidM.Text
'Dim pvpString As String
'pvpString = Replace(txtpvpProd, ",", ".")
'txtpvpProd.Text = pvpString
'Dim ivaString As String
'ivaString = Replace(txtivaProd, ",", ".")
'txtivaProd.Text = ivaString
End Select
End Sub


Public Sub LimpiarCampos()
            txtidProd.Text = ""
            txtnomProd.Text = ""
            txtuniProd.Text = ""
            txtivaProd.Text = ""
            txtpvpProd.Text = ""
            txtidM.Text = ""
            chkFav(2).Value = 0
            txtLleva.Text = ""
            txtPaga.Text = ""
End Sub

