VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddProdProv 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Nuevo producto"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtprecioProd 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtnotaProd 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox txtnomProd 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAddProdProv.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stitch"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre producto"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lbl3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lbl2 
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
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAddProdProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       _________________________________
'       "Iluminación" de los labels




Private Sub Form_Load()
Label4.Caption = "stitch"
End Sub

Public Sub txtnomProd_GotFocus()
lbl1.FontBold = True
lbl1.ForeColor = &HC000C0
End Sub
Public Sub txtnomProd_Lostfocus()
lbl1.FontBold = False
lbl1.ForeColor = &H80000012
End Sub
Public Sub txtprecioProd_GotFocus()
lbl2.FontBold = True
lbl2.ForeColor = &HC000C0
End Sub
Public Sub txtprecioProd_Lostfocus()
lbl2.FontBold = False
lbl2.ForeColor = &H80000012
End Sub
Public Sub txtnotaProd_GotFocus()
lbl3.FontBold = True
lbl3.ForeColor = &HC000C0
End Sub
Public Sub txtnotaProd_Lostfocus()
lbl3.FontBold = False
lbl3.ForeColor = &H80000012
End Sub

'______________________________________
'Usar enter para aceptar
Private Sub txtnomProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtprecioProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 46
If KeyAscii = 13 Then cmdAceptar_Click 'Else Exit Sub
 If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
    
    Else
        KeyAscii = 0
    
    End If

'Dim comma, punto As String


End Sub
Private Sub txtnotaProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub cmdAceptar_Click()
If ValidartlfCliente = True Then
          Select Case sOperacion
                    Case "A"
                              If txtprecioProd.Text = "" Then
                                        txtprecioProd.Text = "0"
                                        'Exit Sub
                              End If
                              Dim codigoProd As Integer
                              
                              With rsUsuarios
                                        If .BOF = False Then
                                                  If .EOF = False Then
                                                            .MoveLast
                                                  End If
                                                  codigoProd = !idProdProv + 1
                                        Else
                                                  codigoProd = 1
                                        End If
                                        .Requery
                                        .AddNew
                                                  !idProdProv = codigoProd
                                                  !nomProd = txtnomProd.Text
                                                  !precioProd = txtprecioProd.Text
                                                  !notaProd = txtnotaProd.Text
                                                  !cifProv = SelectDetalleProv
                                        .Update
                                        .Requery
                                        Unload Me
                              End With
                              rsProd.Requery
          
                    Case "B"
                              With rsUsuarios
                              .Requery
                              .Find "idProdProv = '" & (SelectProd) & "'"
                                        !nomProd = txtnomProd.Text
                                        !precioProd = txtprecioProd.Text
                                        !notaProd = txtnotaProd.Text
                              .UpdateBatch
                              .Requery
                              End With
                              rsProd.Requery
                              Unload Me
                    End Select
Else
                    txtprecioProd.SetFocus
End If

End Sub

Public Sub LimpiarCampos()
txtnomProd.Text = ""
txtprecioProd.Text = ""
txtnotaProd.Text = ""

End Sub

Private Sub Form_Activate()
SSTab1.Caption = ""
MDIfrmMadre.Enabled = False
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
tbProdProv
Select Case sOperacion
          Case "A"
                    Label4.Caption = "Nuevo producto"
          Case "B"
                    Label4.Caption = "Actualizar producto"
                    tbProdProv
                    With rsUsuarios
                              .Requery
                              .Find "idProdProv = '" & (SelectProd) & "'"
                                        txtnomProd.Text = !nomProd
                                        txtprecioProd.Text = !precioProd
                                        txtnotaProd.Text = !notaProd
                    End With
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
SelectProd = ""
MDIfrmMadre.Enabled = True
End Sub

'=========================================================
'ValidartlfCliente
'Usa bCSA_ValidarCampotlfCliente para comprobar si solo hay números en el txttlfCliente
'=========================================================

Public Function ValidartlfCliente()
Dim intboton As Integer
Dim intboton2 As Integer
Dim intboton3 As Integer

    ValidartlfCliente = True
    
 If ValidarCampoNumerico(txtprecioProd.Text) = False Then
        MsgBox "Debe introducir solo números", vbCritical, "Stitch"
        ValidartlfCliente = False
 Else
    If txtprecioProd.Text < 0 Then
          MsgBox "Debe introducir un valor mayor que cero en este campo", vbCritical, "#Racer - Error"
            ValidartlfCliente = False
    End If
 End If
    
End Function
