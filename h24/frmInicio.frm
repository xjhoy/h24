VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stich"
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
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cmdIniCaja 
      BackColor       =   &H00C0C000&
      Caption         =   "Caja"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CheckBox cmdIniAdmin 
      BackColor       =   &H00FF8080&
      Caption         =   "Administrador"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   1320
      Picture         =   "frmInicio.frx":0442
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2400
   End
   Begin VB.Label lblIni 
      BackStyle       =   0  'Transparent
      Caption         =   "BIENVENIDO"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIniAdmin_Click()
frmIniAdmin.Show
Unload Me
End Sub

Private Sub cmdIniCaja_Click()
frmCaja.Show
End Sub
