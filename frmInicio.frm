VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stich"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5640
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cmdIniCaja 
      BackColor       =   &H00FF8080&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CheckBox cmdIniAdmin 
      BackColor       =   &H00C0C000&
      Caption         =   "Panel de control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stitch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pan y Caprichos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   1680
      Picture         =   "frmInicio.frx":058A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2160
   End
   Begin VB.Label lblIni 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hola!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Inicio del programa despues de haber pasado por mod1/sub main
'==========================================

'Private Sub cmdIniAdmin_Click()
'cmdIniAdmin.Value = 0
'MDIfrmMadre.Show
'Unload Me
'End Sub
'
'Private Sub cmdIniCaja_Click()
'cmdIniCaja.Value = 0
'frmCaja.Show
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'If chkCaja = False Then
'          cmdIniCaja.Enabled = False
'Else
'          cmdIniCaja.Enabled = True
'End If
'Me.Left = (Screen.Width - Me.Width) / 2
'Me.Top = (Screen.Height - Me.Height) / 2
'
'End Sub

