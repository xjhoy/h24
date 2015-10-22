VERSION 5.00
Begin VB.Form frmModPass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stitch - Cambio de contraseña"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmModPass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewPassUser 
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
      Left            =   720
      PasswordChar    =   "×"
      TabIndex        =   2
      Top             =   3720
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
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
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
      Left            =   720
      PasswordChar    =   "×"
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva contraseña"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Antigua contraseña"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar Contraseña"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmModPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
frmIniAdmin.Show
End Sub


Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

Private Sub txtPassUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

Private Sub txtNewPassUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAceptar_Click Else Exit Sub
End Sub

Private Sub cmdAceptar_Click()
If txtUser.Text = "" Then
          MsgBox "Debe introducir el usuario", vbInformation, Me.Caption
          txtUser.SetFocus
          Exit Sub
End If
If txtPassUser.Text = "" Then
          MsgBox "Debe introducir la  anterior contraseña", vbInformation, Me.Caption
          txtPassUser.SetFocus
          Exit Sub
End If
If txtNewPassUser.Text = "" Then
          MsgBox "Debe introducir una nueva contraseña", vbInformation, Me.Caption
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
                    If !passUser = txtPassUser Then
                              !passUser = txtNewPassUser.Text
                              .UpdateBatch
                              MsgBox "Contraseña cambiada!", vbDefaultButton1, Me.Caption
                              txtPassUser.Text = ""
                              txtUser.Text = ""
                              txtNewPassUser.Text = ""
                              cmdCancelar.SetFocus

                    Else
                              MsgBox "Usuario o contraseña incorrecta", vbInformation, "Aviso"
                              txtPassUser.Text = ""
                              txtUser.Text = ""
                              txtNewPassUser.Text = ""
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
Private Sub txtnewPassUser_GotFocus()
    Label4.FontBold = True
    Label4.ForeColor = &HC000C0
End Sub
Private Sub txtnewPassUser_LostFocus()
    Label4.FontBold = False
    Label4.ForeColor = &H80000012
End Sub
