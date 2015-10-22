VERSION 5.00
Begin VB.Form frmIniAdmin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Iniciar sesión"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
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
   ScaleHeight     =   4170
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtPassUser 
      Alignment       =   2  'Center
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar sesión"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmIniAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
If txtUser.Text = "" Then MsgBox "Debe introducir un usuario", vbInformation, Me.Caption: txtUser.SetFocus
If txtPassUser.Text = "" Then MsgBox "Debe introducir una contraseña", vbInformation, Me.Caption: txtPassUser.SetFocus
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
    MDIfrmMadre.Show
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
frmInicio.Show
Unload Me
End Sub

Private Sub Form_Load()
tbUsers
End Sub

