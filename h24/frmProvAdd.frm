VERSION 5.00
Begin VB.Form frmProvAdd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nuevo proveedor"
   ClientHeight    =   7785
   ClientLeft      =   -8475
   ClientTop       =   1035
   ClientWidth     =   10905
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
   ScaleHeight     =   7785
   ScaleWidth      =   10905
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   10560
      TabIndex        =   20
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   7680
      TabIndex        =   19
      Top             =   8760
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   14160
      TabIndex        =   17
      Top             =   4920
      Width           =   3735
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   15240
      TabIndex        =   15
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   14280
      TabIndex        =   13
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   8760
      TabIndex        =   11
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   8760
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   6480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   18
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Persona de contacto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   16
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   14
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Código postal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CIF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Razón social"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   6000
      Picture         =   "frmProvAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   7560
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "frmProvAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
