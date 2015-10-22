VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProvAdd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nuevo proveedor"
   ClientHeight    =   4155
   ClientLeft      =   -8475
   ClientTop       =   1035
   ClientWidth     =   6000
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
   ScaleHeight     =   4155
   ScaleWidth      =   6000
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
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
      Left            =   10560
      TabIndex        =   10
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   7680
      TabIndex        =   9
      Top             =   8760
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   840
      TabIndex        =   12
      Top             =   2880
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProvAdd.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtdirProv"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtcpProv"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtlocProv"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtcifProv"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProvinciaProv"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtnomProv"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   5415
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
         Left            =   2040
         TabIndex        =   5
         Top             =   3840
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
         Left            =   2040
         TabIndex        =   1
         Top             =   1320
         Width           =   2295
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
         Left            =   5520
         TabIndex        =   4
         Top             =   3000
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
         Left            =   2040
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
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
         Left            =   2040
         TabIndex        =   2
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label lbl3 
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
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lbl4 
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
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lbl2 
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
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbl6 
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
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lbl1 
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl5 
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
         Left            =   4080
         TabIndex        =   13
         Top             =   3120
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3135
      Left            =   10320
      TabIndex        =   19
      Top             =   2880
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProvAdd.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtemailProv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtcontProv"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txttlfProv"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
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
         Left            =   2880
         TabIndex        =   6
         Top             =   480
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
         Left            =   2880
         TabIndex        =   7
         Top             =   1320
         Width           =   3735
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
         Left            =   2880
         TabIndex        =   8
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label lbl7 
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lbl8 
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
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lbl9 
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
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   6000
      Picture         =   "frmProvAdd.frx":0038
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo proveedor"
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
      Left            =   8400
      TabIndex        =   11
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmProvAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'               _____________________________________
'               "Iluminación" de los labels del formularios

Private Sub txtnomProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtcifProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtprovinciaProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtlocProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtdirProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtcpProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click ' Else Exit Sub
If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
          Exit Sub
Else
          KeyAscii = 0
End If
End Sub
Private Sub txtemailProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtcontProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txttlfProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click ' Else Exit Sub
If KeyAscii >= 44 And KeyAscii <= 57 Or KeyAscii = 8 Then
       Exit Sub
Else
        KeyAscii = 0
End If
End Sub

Public Sub txtnomProv_GotFocus()
lbl1.FontBold = True
lbl1.ForeColor = &HC000C0
End Sub
Public Sub txtcifProv_gotfocus()
lbl2.FontBold = True
lbl2.ForeColor = &HC000C0

End Sub
Public Sub txtdirProv_GotFocus()
lbl3.FontBold = True
lbl3.ForeColor = &HC000C0
End Sub
Public Sub txtcpProv_gotfocus()
lbl4.FontBold = True
lbl4.ForeColor = &HC000C0

End Sub
Public Sub txtlocProv_GotFocus()
lbl5.FontBold = True
lbl5.ForeColor = &HC000C0
End Sub
Public Sub txtprovinciaProv_gotfocus()
lbl6.FontBold = True
lbl6.ForeColor = &HC000C0

End Sub
Public Sub txttlfProv_GotFocus()
lbl7.FontBold = True
lbl7.ForeColor = &HC000C0
End Sub
Public Sub txtcontProv_gotfocus()
lbl8.FontBold = True
lbl8.ForeColor = &HC000C0

End Sub
Public Sub txtemailProv_gotfocus()
lbl9.FontBold = True
lbl9.ForeColor = &HC000C0

End Sub
Public Sub txtnomProv_Lostfocus()
lbl1.FontBold = True
lbl1.ForeColor = &H80000012

End Sub
Public Sub txtcifProv_Lostfocus()
lbl2.FontBold = True
lbl2.ForeColor = &H80000012

End Sub
Public Sub txtdirProv_Lostfocus()
lbl3.FontBold = True
lbl3.ForeColor = &H80000012

End Sub
Public Sub txtcpProv_Lostfocus()
lbl4.FontBold = True
lbl4.ForeColor = &H80000012

End Sub
Public Sub txtlocProv_Lostfocus()
lbl5.FontBold = True
lbl5.ForeColor = &H80000012

End Sub
Public Sub txtprovinciaProv_Lostfocus()
lbl6.FontBold = True
lbl6.ForeColor = &H80000012

End Sub
Public Sub txttlfProv_Lostfocus()
lbl7.FontBold = True
lbl7.ForeColor = &H80000012

End Sub
Public Sub txtcontProv_Lostfocus()
lbl8.FontBold = True
lbl8.ForeColor = &H80000012

End Sub
Public Sub txtemailProv_Lostfocus()
lbl9.FontBold = True
lbl9.ForeColor = &H80000012

End Sub

'                   Fin de "Iiluminación"
'                  ____________________


Private Sub Cancelar_Click()
Set frmListProv.dtgProv.DataSource = rsUsuarios
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAceptar_Click()
Select Case sOperacion
Case "A"

If txtnomProv.Text = "" Then
          MsgBox "Debe escribir un nombre de proveedor", vbExclamation, Me.Caption
          txtnomProv.SetFocus
          Exit Sub
End If

If txtcifProv.Text = "" Then
          MsgBox "Debe escribir el cif de proveedor", vbExclamation, Me.Caption
          txtcifProv.SetFocus
          Exit Sub
End If

Dim intboton As Integer
intboton = MsgBox("¿Desea guardar este proveedor?", vbQuestion Or vbOKCancel, Me.Caption)

If intboton = vbOK Then
          With rsUsuarios
                    
                    If .BOF = False Then
                              .MoveFirst
                    End If
                    
                    .Find "cifProv= '" + txtcifProv + "'"
                    
                    If .EOF = False Then
                              intboton = MsgBox("Este proveedor existe en la base de datos", vbQuestion Or vbOKOnly, Me.Caption)
                              txtcifProv.SetFocus
    
                     Else
                              .Requery
                              .AddNew
                                        !nomProv = txtnomProv.Text
                                        !cifProv = txtcifProv.Text
                                        !provinciaProv = txtProvinciaProv.Text
                                        !locProv = txtlocProv.Text
                                        !dirprov = txtdirProv.Text
                                        !cpProv = txtcpProv.Text
                                        !tlfProv = txttlfProv.Text
                                        !emailProv = txtemailProv.Text
                                        !contProv = txtcontProv.Text
                              .Update
                              .Requery
                              
                              MsgBox "Proveedor creado", vbInformation, Me.Caption
                              intboton = MsgBox("¿Desea crear otro proveedor?", vbQuestion Or vbYesNo, Me.Caption)
                              If intboton = vbYes Then
                                        LimpiarCampos
                                        txtnomProv.SetFocus
                              Else
                                        Unload Me
                              End If
                    End If
          End With
End If


Case "B"

With rsUsuarios
          .Requery
          .Find "cifProv = '" & (selectId) & "'"
                    !nomProv = txtnomProv.Text
                    !cifProv = txtcifProv.Text
                    !provinciaProv = txtProvinciaProv.Text
                    !locProv = txtlocProv.Text
                    !dirprov = txtdirProv.Text
                    !cpProv = txtcpProv.Text
                    !tlfProv = txttlfProv.Text
                    !emailProv = txtemailProv.Text
                    !contProv = txtcontProv.Text
          .UpdateBatch
          .Requery
          Unload Me
End With

End Select


End Sub

Private Sub Form_Load()

SSTab1.Caption = ""
SSTab2.Caption = ""
tbProv

Select Case sOperacion

Case "A"
LimpiarCampos

Case "B"
Label1.Caption = "Modificar proveedor"

With rsUsuarios
.Requery
.Find "cifProv = '" & (selectId) & "'"
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

End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
With rsUsuarios
          If .State = 1 Then .Close
End With
End Sub

Public Sub LimpiarCampos()
            txtnomProv.Text = ""
            txtcifProv.Text = ""
            txtProvinciaProv.Text = ""
            txtlocProv.Text = ""
            txtdirProv.Text = ""
            txtcpProv.Text = ""
            txttlfProv.Text = ""
            txtemailProv.Text = ""
            txtcontProv.Text = ""
End Sub

Public Sub validarCampos()

If txtnomProv.Text = "" Then
          MsgBox "Debe escribir un nombre de proveedor", vbExclamation, Me.Caption
          txtnomProv.SetFocus
End If

End Sub



