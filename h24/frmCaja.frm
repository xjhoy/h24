VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCaja 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Caja"
   ClientHeight    =   9915
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10650
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
   ScaleHeight     =   9915
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2520
      Width           =   3615
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   10440
      TabIndex        =   14
      Top             =   1920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6800
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      Separators      =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Index           =   11
      Left            =   10680
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   735
      Index           =   9
      Left            =   16200
      TabIndex        =   12
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   735
      Index           =   8
      Left            =   15000
      TabIndex        =   11
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   735
      Index           =   7
      Left            =   16200
      TabIndex        =   10
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   735
      Index           =   6
      Left            =   13800
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   735
      Index           =   5
      Left            =   15000
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   735
      Index           =   4
      Left            =   16200
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   735
      Index           =   3
      Left            =   15000
      TabIndex        =   6
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   735
      Index           =   2
      Left            =   13800
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   735
      Index           =   1
      Left            =   16200
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   735
      Index           =   0
      Left            =   15000
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   2040
      TabIndex        =   1
      Top             =   4800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   735
      Index           =   10
      Left            =   13800
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   7560
      Picture         =   "frmCaja.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   17
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Favoritos"
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
      Left            =   14400
      TabIndex        =   16
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   6975
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click(Index As Integer)
Unload Me
End Sub

