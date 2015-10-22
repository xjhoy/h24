VERSION 5.00
Begin VB.Form frmIniAdmin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stitch - Iniciar sesión"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIniAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtPassUser 
      Alignment       =   2  'Center
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
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "×"
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblModPass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¿Desea cambiar su contraseña?"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmIniAdmin.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmIniAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
MDIfrmMadre.Enabled = True
End Sub




Private Sub lblModPass_Click()
frmModPass.Show
Unload Me
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub
Private Sub txtPassUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

Private Sub cmdAceptar_Click()
If txtUser.Text = "" Then
          MsgBox "Debe introducir un usuario", vbInformation, Me.Caption
          txtUser.SetFocus
          Exit Sub
End If
If txtPassUser.Text = "" Then
          MsgBox "Debe introducir una contraseña", vbInformation, Me.Caption
          txtPassUser.SetFocus
          Exit Sub
End If
With rsUsuarios
          .Requery
          .Find "iduser = '" & Trim(txtUser.Text) & "'"
          If .EOF Then
                    MsgBox "Usuario o contraseña incorrecta", vbInformation, "Aviso"
                    txtUser.Text = ""
                    txtPassUser.Text = ""
                    txtUser.SetFocus
                    Exit Sub
          Else
                    If !passUser = Trim(txtPassUser.Text) Then
                              MDIfrmMadre.Toolbar1.Buttons(7).Enabled = True
                              MDIfrmMadre.Toolbar1.Buttons(9).Enabled = True
                              chkadmin = True
                              MDIfrmMadre.Toolbar1.Buttons(1).Caption = "Cerrar sesión"
                              Unload Me
                    Else
                              MsgBox "Usuario o contraseña incorrecta", vbInformation, "Aviso"
                              txtPassUser.Text = ""
                              txtUser.Text = ""
                              txtUser.SetFocus
                              Exit Sub
                    End If
          End If

End With


End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIfrmMadre.Enabled = False
tbUsers
End Sub

Private Sub txtUser_GotFocus()
    Label1.FontBold = True
    Label1.ForeColor = &HC000C0
End Sub
Private Sub txtUser_LostFocus()
    Label1.FontBold = False
    Label1.ForeColor = &H80000012
End Sub
Private Sub txtPassUser_GotFocus()
    Label2.FontBold = True
    Label2.ForeColor = &HC000C0
End Sub
Private Sub txtPassUser_LostFocus()
    Label2.FontBold = False
    Label2.ForeColor = &H80000012
End Sub
