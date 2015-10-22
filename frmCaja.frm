VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCaja 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Caja"
   ClientHeight    =   8445
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
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
   Icon            =   "frmCaja.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8445
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCant 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      MaxLength       =   4
      TabIndex        =   161
      Text            =   "1"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPprecio 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11040
      TabIndex        =   159
      Text            =   "0"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkPrecio 
      BackColor       =   &H0080FFFF&
      Caption         =   "P. Precio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtBuscar 
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
      Left            =   1560
      TabIndex        =   155
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton btnMenos 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   16920
      TabIndex        =   31
      Top             =   12360
      Width           =   1335
   End
   Begin VB.CommandButton btnRetroceso 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13560
      Picture         =   "frmCaja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   12360
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   15240
      TabIndex        =   29
      Top             =   12360
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   16920
      TabIndex        =   28
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   15240
      TabIndex        =   27
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   13560
      TabIndex        =   26
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   16920
      TabIndex        =   25
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   15240
      TabIndex        =   24
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   13560
      TabIndex        =   23
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   16920
      TabIndex        =   22
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   15240
      TabIndex        =   21
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   13560
      TabIndex        =   20
      Top             =   8400
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid fldCaja 
      Height          =   6615
      Left            =   360
      TabIndex        =   18
      Top             =   5880
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11668
      _Version        =   393216
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   16777215
      GridColorFixed  =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkDto 
      Caption         =   "Dto."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox txtDto 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11040
      TabIndex        =   2
      Text            =   "0"
      Top             =   9480
      Visible         =   0   'False
      Width           =   975
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
      Left            =   9480
      TabIndex        =   3
      Top             =   10320
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Cobrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9480
      Picture         =   "frmCaja.frx":1254
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   11520
      Width           =   2895
   End
   Begin VB.TextBox txtcod 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11040
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   11
      Left            =   16200
      Picture         =   "frmCaja.frx":17DE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1850
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6855
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCaja.frx":34A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   9480
      TabIndex        =   32
      Top             =   3120
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   8388736
      TabCaption(0)   =   "Pan"
      TabPicture(0)   =   "frmCaja.frx":34C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdFavA(23)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFavA(22)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFavA(21)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdFavA(20)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdFavA(19)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFavA(18)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdFavA(17)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFavA(16)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFavA(15)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdFavA(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdFavA(13)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdFavA(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdFavA(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdFavA(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdFavA(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdFavA(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdFavA(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdFavA(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdFavA(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdFavA(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdFavA(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdFavA(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdFavA(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdFavA(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Bolleria"
      TabPicture(1)   =   "frmCaja.frx":3A5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape5"
      Tab(1).Control(1)=   "cmdFavB(23)"
      Tab(1).Control(2)=   "cmdFavB(22)"
      Tab(1).Control(3)=   "cmdFavB(21)"
      Tab(1).Control(4)=   "cmdFavB(20)"
      Tab(1).Control(5)=   "cmdFavB(19)"
      Tab(1).Control(6)=   "cmdFavB(18)"
      Tab(1).Control(7)=   "cmdFavB(17)"
      Tab(1).Control(8)=   "cmdFavB(16)"
      Tab(1).Control(9)=   "cmdFavB(15)"
      Tab(1).Control(10)=   "cmdFavB(14)"
      Tab(1).Control(11)=   "cmdFavB(13)"
      Tab(1).Control(12)=   "cmdFavB(12)"
      Tab(1).Control(13)=   "cmdFavB(11)"
      Tab(1).Control(14)=   "cmdFavB(10)"
      Tab(1).Control(15)=   "cmdFavB(9)"
      Tab(1).Control(16)=   "cmdFavB(8)"
      Tab(1).Control(17)=   "cmdFavB(7)"
      Tab(1).Control(18)=   "cmdFavB(6)"
      Tab(1).Control(19)=   "cmdFavB(5)"
      Tab(1).Control(20)=   "cmdFavB(4)"
      Tab(1).Control(21)=   "cmdFavB(3)"
      Tab(1).Control(22)=   "cmdFavB(2)"
      Tab(1).Control(23)=   "cmdFavB(1)"
      Tab(1).Control(24)=   "cmdFavB(0)"
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "Promociones"
      TabPicture(2)   =   "frmCaja.frx":3FF8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape4"
      Tab(2).Control(1)=   "cmdfavC(21)"
      Tab(2).Control(2)=   "cmdfavC(22)"
      Tab(2).Control(3)=   "cmdfavC(23)"
      Tab(2).Control(4)=   "cmdfavC(17)"
      Tab(2).Control(5)=   "cmdfavC(16)"
      Tab(2).Control(6)=   "cmdfavC(20)"
      Tab(2).Control(7)=   "cmdfavC(18)"
      Tab(2).Control(8)=   "cmdfavC(19)"
      Tab(2).Control(9)=   "cmdfavC(15)"
      Tab(2).Control(10)=   "cmdfavC(14)"
      Tab(2).Control(11)=   "cmdfavC(13)"
      Tab(2).Control(12)=   "cmdfavC(12)"
      Tab(2).Control(13)=   "cmdfavC(10)"
      Tab(2).Control(14)=   "cmdfavC(11)"
      Tab(2).Control(15)=   "cmdfavC(5)"
      Tab(2).Control(16)=   "cmdfavC(4)"
      Tab(2).Control(17)=   "cmdfavC(8)"
      Tab(2).Control(18)=   "cmdfavC(9)"
      Tab(2).Control(19)=   "cmdfavC(6)"
      Tab(2).Control(20)=   "cmdfavC(7)"
      Tab(2).Control(21)=   "cmdfavC(3)"
      Tab(2).Control(22)=   "cmdfavC(2)"
      Tab(2).Control(23)=   "cmdfavC(1)"
      Tab(2).Control(24)=   "cmdfavC(0)"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Bebidas"
      TabPicture(3)   =   "frmCaja.frx":4592
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape3"
      Tab(3).Control(1)=   "cmdfavD(21)"
      Tab(3).Control(2)=   "cmdfavD(7)"
      Tab(3).Control(3)=   "cmdfavD(6)"
      Tab(3).Control(4)=   "cmdfavD(9)"
      Tab(3).Control(5)=   "cmdfavD(8)"
      Tab(3).Control(6)=   "cmdfavD(11)"
      Tab(3).Control(7)=   "cmdfavD(10)"
      Tab(3).Control(8)=   "cmdfavD(12)"
      Tab(3).Control(9)=   "cmdfavD(13)"
      Tab(3).Control(10)=   "cmdfavD(14)"
      Tab(3).Control(11)=   "cmdfavD(15)"
      Tab(3).Control(12)=   "cmdfavD(19)"
      Tab(3).Control(13)=   "cmdfavD(18)"
      Tab(3).Control(14)=   "cmdfavD(20)"
      Tab(3).Control(15)=   "cmdfavD(16)"
      Tab(3).Control(16)=   "cmdfavD(17)"
      Tab(3).Control(17)=   "cmdfavD(23)"
      Tab(3).Control(18)=   "cmdfavD(22)"
      Tab(3).Control(19)=   "cmdfavD(0)"
      Tab(3).Control(20)=   "cmdfavD(1)"
      Tab(3).Control(21)=   "cmdfavD(2)"
      Tab(3).Control(22)=   "cmdfavD(3)"
      Tab(3).Control(23)=   "cmdfavD(4)"
      Tab(3).Control(24)=   "cmdfavD(5)"
      Tab(3).ControlCount=   25
      TabCaption(4)   =   "Varios"
      TabPicture(4)   =   "frmCaja.frx":4B2C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Shape2"
      Tab(4).Control(1)=   "cmdfavE(21)"
      Tab(4).Control(2)=   "cmdfavE(7)"
      Tab(4).Control(3)=   "cmdfavE(6)"
      Tab(4).Control(4)=   "cmdfavE(9)"
      Tab(4).Control(5)=   "cmdfavE(8)"
      Tab(4).Control(6)=   "cmdfavE(11)"
      Tab(4).Control(7)=   "cmdfavE(10)"
      Tab(4).Control(8)=   "cmdfavE(12)"
      Tab(4).Control(9)=   "cmdfavE(13)"
      Tab(4).Control(10)=   "cmdfavE(14)"
      Tab(4).Control(11)=   "cmdfavE(15)"
      Tab(4).Control(12)=   "cmdfavE(19)"
      Tab(4).Control(13)=   "cmdfavE(18)"
      Tab(4).Control(14)=   "cmdfavE(20)"
      Tab(4).Control(15)=   "cmdfavE(16)"
      Tab(4).Control(16)=   "cmdfavE(17)"
      Tab(4).Control(17)=   "cmdfavE(23)"
      Tab(4).Control(18)=   "cmdfavE(22)"
      Tab(4).Control(19)=   "cmdfavE(0)"
      Tab(4).Control(20)=   "cmdfavE(1)"
      Tab(4).Control(21)=   "cmdfavE(2)"
      Tab(4).Control(22)=   "cmdfavE(3)"
      Tab(4).Control(23)=   "cmdfavE(4)"
      Tab(4).Control(24)=   "cmdfavE(5)"
      Tab(4).ControlCount=   25
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   5
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   4
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   3
         Left            =   -70200
         TabIndex        =   151
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   2
         Left            =   -71640
         TabIndex        =   150
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   1
         Left            =   -73080
         TabIndex        =   149
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   0
         Left            =   -74520
         TabIndex        =   148
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   5
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   4
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   3
         Left            =   -70200
         TabIndex        =   145
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   2
         Left            =   -71640
         TabIndex        =   144
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   1
         Left            =   -73080
         TabIndex        =   143
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   0
         Left            =   -74520
         TabIndex        =   142
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   0
         Left            =   -74520
         TabIndex        =   141
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   1
         Left            =   -73080
         TabIndex        =   140
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   2
         Left            =   -71640
         TabIndex        =   139
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   3
         Left            =   -70200
         TabIndex        =   138
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   7
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   6
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   9
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   8
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   4
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   5
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   11
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   10
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   12
         Left            =   -74520
         TabIndex        =   129
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   13
         Left            =   -73080
         TabIndex        =   128
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   14
         Left            =   -71640
         TabIndex        =   127
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   15
         Left            =   -70200
         TabIndex        =   126
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   19
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   18
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   20
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   16
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   17
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   23
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   22
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavC 
         Caption         =   "Command2"
         Height          =   855
         Index           =   21
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   0
         Left            =   -74520
         TabIndex        =   117
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   1
         Left            =   -73080
         TabIndex        =   116
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   2
         Left            =   -71640
         TabIndex        =   115
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   3
         Left            =   -70200
         TabIndex        =   114
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   4
         Left            =   -68760
         TabIndex        =   113
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   5
         Left            =   -67320
         TabIndex        =   112
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   6
         Left            =   -74520
         TabIndex        =   111
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   7
         Left            =   -73080
         TabIndex        =   110
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   8
         Left            =   -71640
         TabIndex        =   109
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   9
         Left            =   -70200
         TabIndex        =   108
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   10
         Left            =   -68760
         TabIndex        =   107
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   11
         Left            =   -67320
         TabIndex        =   106
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   12
         Left            =   -74520
         TabIndex        =   105
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   13
         Left            =   -73080
         TabIndex        =   104
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   14
         Left            =   -71640
         TabIndex        =   103
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   15
         Left            =   -70200
         TabIndex        =   102
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   16
         Left            =   -68760
         TabIndex        =   101
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   17
         Left            =   -67320
         TabIndex        =   100
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   18
         Left            =   -74520
         TabIndex        =   99
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   19
         Left            =   -73080
         TabIndex        =   98
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   20
         Left            =   -71640
         TabIndex        =   97
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   21
         Left            =   -70200
         TabIndex        =   96
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   22
         Left            =   -68760
         TabIndex        =   95
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavB 
         Caption         =   "Command1"
         Height          =   855
         Index           =   23
         Left            =   -67320
         TabIndex        =   94
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Command2"
         Height          =   855
         Left            =   360
         TabIndex        =   93
         Top             =   -1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   22
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   23
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   17
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   16
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   20
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   18
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   19
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   15
         Left            =   -70200
         TabIndex        =   85
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   14
         Left            =   -71640
         TabIndex        =   84
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   13
         Left            =   -73080
         TabIndex        =   83
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   12
         Left            =   -74520
         TabIndex        =   82
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   10
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   11
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   8
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   9
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   6
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   7
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavD 
         Caption         =   "Command2"
         Height          =   855
         Index           =   21
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   22
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   23
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   17
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   16
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   20
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   18
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   19
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   15
         Left            =   -70200
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   14
         Left            =   -71640
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   13
         Left            =   -73080
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   12
         Left            =   -74520
         TabIndex        =   64
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   10
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   11
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   8
         Left            =   -71640
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   9
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   6
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   7
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdfavE 
         Caption         =   "Command2"
         Height          =   855
         Index           =   21
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   2
         Left            =   3360
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   3
         Left            =   4800
         TabIndex        =   53
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   4
         Left            =   6240
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   5
         Left            =   7680
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   6
         Left            =   480
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   7
         Left            =   1920
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   8
         Left            =   3360
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   9
         Left            =   4800
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   10
         Left            =   6240
         TabIndex        =   46
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   11
         Left            =   7680
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   12
         Left            =   480
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   13
         Left            =   1920
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   14
         Left            =   3360
         TabIndex        =   42
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   15
         Left            =   4800
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   16
         Left            =   6240
         TabIndex        =   40
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   17
         Left            =   7680
         TabIndex        =   39
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   18
         Left            =   480
         TabIndex        =   38
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   19
         Left            =   1920
         TabIndex        =   37
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   20
         Left            =   3360
         TabIndex        =   36
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   21
         Left            =   4800
         TabIndex        =   35
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   22
         Left            =   6240
         TabIndex        =   34
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFavA 
         Caption         =   "Command1"
         Height          =   855
         Index           =   23
         Left            =   7680
         TabIndex        =   33
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080FFFF&
         Height          =   4575
         Left            =   120
         Top             =   360
         Width           =   8895
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         Height          =   4575
         Left            =   -74880
         Top             =   360
         Width           =   8895
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   4575
         Left            =   -74880
         Top             =   360
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         Height          =   4575
         Left            =   -74880
         Top             =   360
         Width           =   8895
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         Height          =   4575
         Left            =   -74880
         Top             =   360
         Width           =   8895
      End
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   3015
      Left            =   240
      TabIndex        =   154
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCaja.frx":50C6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgBuscar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtgBuscar 
         CausesValidation=   0   'False
         Height          =   2775
         Left            =   120
         TabIndex        =   157
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "idProd"
            Caption         =   "Código"
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
            DataField       =   "nomProd"
            Caption         =   "Descripción"
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
         BeginProperty Column02 
            DataField       =   "uniProd"
            Caption         =   "Unidades"
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
         BeginProperty Column03 
            DataField       =   "pvpProd"
            Caption         =   "Precio Unidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00 €"
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
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3404,977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1275,024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   160
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblBusq 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar"
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
      Left            =   600
      TabIndex        =   156
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento activado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   855
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   17
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   12720
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   12720
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cód."
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
      Left            =   10320
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
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
      Left            =   6720
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stitch"
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
      Left            =   9840
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pan - Caprichos"
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
      Left            =   9360
      TabIndex        =   11
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Favoritos"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total unidades: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   12840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total venta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   12840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   7560
      Picture         =   "frmCaja.frx":50E2
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
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- Boton Retroceso, borrar numero --
'====================================
Private Sub btnRetroceso_Click()
    If txtCant.Text <> "" Then
        txtCant.Text = Mid(txtCant.Text, 1, Len(txtCant.Text) - 1)
    End If
End Sub

'-- Boton cobrar --
'===================
Private Sub Check1_Click()

    If Check1.Value = 1 Then
        If Not txtcod = "" Then
            MsgBox "Hay un producto sin registrar"
            Check1.Value = 0
            Exit Sub
        Else
            With fldCaja
                If .Row = 0 Then
                    Check1.Value = 0
                    Exit Sub
                End If
            End With
            Check1.Value = 0
            frmCobrar.Show
        End If
    End If


End Sub

'=======================================
'CHECKBOX de descuento, para habilitar el TEXTBOX
'=======================================
Private Sub chkDto_Click()
    
    'Cambia el estado del CHECKBOX cada vez que se da click
    
    If chkDto.Value = 1 Then
        txtDto.Enabled = True
        chkDto.ForeColor = vbWhite
        chkDto.BackColor = &HFF&
        txtDto.Visible = True
        Label12.Visible = True
        txtDto.Text = ""
        txtDto.SetFocus
        Shape1.Visible = True
        Label13.Visible = True
        Label13.Caption = "Descuento activado del " & txtDto.Text & " %"
        chkPrecio.Enabled = False
    Else
        txtDto.Enabled = False
        txtDto.Text = 0
        chkDto.ForeColor = vbBlack
        chkDto.BackColor = &H8000000F
        Shape1.Visible = False
        Label13.Visible = False
        txtDto.Visible = False
        Label12.Visible = False
        chkPrecio.Enabled = True
    End If

End Sub

Private Sub chkPrecio_Click()

    'Cambia el estado del CHECKBOX cada vez que se da click
    If chkPrecio.Value = 1 Then
        txtPprecio.Visible = True
        txtPprecio.Text = ""
        txtPprecio.SetFocus
        chkDto.Enabled = False
        Shape1.Visible = True
        Label13.Visible = True
        Label13.Caption = "Personalizar precio a " & txtPprecio.Text & " €"
        
    Else
        txtPprecio.Text = 0
        txtPprecio.Visible = False
        chkDto.Enabled = True
        Shape1.Visible = False
        Label13.Visible = False
        
    End If

End Sub

'-- Dar click en un favorito A Pan--
'=====================================
Private Sub cmdfavA_Click(Index As Integer)
    txtcod.Text = X0(Index)
    cmdAceptar_Click
End Sub

'-- Dar click en un favorito B Bolleria--
'==========================================
Private Sub cmdfavB_Click(Index As Integer)
    txtcod.Text = X1(Index)
    cmdAceptar_Click
End Sub

'-- Dar click en un favorito C Promociones--
'===========================================
Private Sub cmdfavC_Click(Index As Integer)
    txtcod.Text = X2(Index)
    cmdAceptar_Click
End Sub

'-- Dar click en un favorito D Bebidas--
'========================================
Private Sub cmdfavD_Click(Index As Integer)
    txtcod.Text = X3(Index)
    cmdAceptar_Click
End Sub

'-- Dar click en un favorito E Varios--
'=======================================
Private Sub cmdfavE_Click(Index As Integer)
    txtcod.Text = X4(Index)
    cmdAceptar_Click
End Sub

'-- Botones numeros --
'=======================
Private Sub btnNum_Click(Index As Integer)
    txtCant.Text = txtCant.Text & CInt(Index)
End Sub

'-- Boton signo menos --
'========================
Private Sub btnMenos_Click()
    txtCant.MaxLength = 3
    txtCant.Text = "-" & txtCant.Text
    txtCant.MaxLength = 2
End Sub

'-- Doble click en el dataGrid Buscar
'=====================================
Private Sub dtgBuscar_DblClick()
    On Error GoTo error
    
    txtcod.Text = dtgBuscar.Columns(0).Text
    cmdAceptar_Click
    
error:
    
End Sub

'-- Escribir en el textBox Buscar --
'====================================
Private Sub txtBuscar_Change()
    Dim bProd As String
    bProd = txtBuscar.Text
    tbBuscarProd (bProd)
    Set dtgBuscar.DataSource = rsBuscarProd
End Sub

'-- Abrir cajón desde txtBuscar --
'=================================
Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        Dim canal%
        Dim Impresora
        Impresora = "\\127.0.0.1\" & Printer.DeviceName
        canal = FreeFile
        Open Impresora For Output As #canal
        
        'Drawer Kick (ESC p)
        Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
        Close #canal
    End If
    If KeyCode = vbKeyF5 Then
        Check1.Value = 1
    End If
End Sub

'-- Abrir cajón desde txtCod --
'===============================
Private Sub txtcod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        Dim canal%
        Dim Impresora
        Impresora = "\\127.0.0.1\" & Printer.DeviceName
        canal = FreeFile
        Open Impresora For Output As #canal
        
        'Drawer Kick (ESC p)
        Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
        Close #canal
    End If
    If KeyCode = vbKeyF5 Then
        Check1.Value = 1
    End If
End Sub

Private Sub Form_Activate()
    
    txtBuscar.SetFocus
    
    '-- Mostrar favoritos --
    '=========================
    Dim f As String
    f = "favA"
    
    tbFavoritosCount (f)
        
    With rsAlmacen
        countFav = !Expre
    End With
    
    tbFav (f)
    
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                cmdFavA(c).Visible = True
                cmdFavA(c).Caption = !nomProd
                X0(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favB"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdFavB(c).Visible = True
                cmdFavB(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X1(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favC"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavC(c).Visible = True
                cmdfavC(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X2(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favD"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavD(c).Visible = True
                cmdfavD(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X3(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favE"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavE(c).Visible = True
                cmdfavE(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X4(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    'Cargar tabla de buscar con todos los productos

End Sub

Private Sub Form_Load()

    Dim c As Integer
    
    txtDto.MaxLength = 2
    txtCant.Text = "1"
    '==========================
    'FAVORITOS
    '==========================
    
    Dim f As String
    f = "favA"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdFavA(c).Visible = True
                cmdFavA(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X0(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favB"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdFavB(c).Visible = True
                cmdFavB(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X1(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favC"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavC(c).Visible = True
                cmdfavC(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X2(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favD"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavD(c).Visible = True
                cmdfavD(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X3(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    
    f = "favE"
    tbFavoritosCount (f)
    
    With rsAlmacen
        countFav = !Expre
    End With
    tbFav (f)
    
    'Abrimos recorset para tener los favoritos
    With rsAlmacen
        'Recorremos la tabla desde 0 hasta el numero de favoritos que nos da la consulta anterior
        For c = 0 To countFav
            'Comprobamos que estamos en el fin
            If .EOF = False Then
                'Damos que sea visible el boton y le asignamos la propiedad caption
                cmdfavE(c).Visible = True
                cmdfavE(c).Caption = !nomProd
                'X1(36) Variable publica  que guarda el codigo de barras de cada boton
                X4(c) = !idProd
                'Mueve al siguiente registro
                .MoveNext
            End If
        Next c
    End With
    '===========================
    '===========================
    
    
    'Quitar el titulo del SSTAB
    SSTab2.Caption = ""
    SSTab3.Caption = ""
    
    'Dibujar el GRID
    Filas = 1
    fldCaja.Rows = Filas
    fldCaja.Rows = 1
    fldCaja.Cols = 9
    
    'Ocultar la primer columna del grid
    fldCaja.ColWidth(0) = 0
    
    'Poner los ENCABEZADOS del GRID
    With fldCaja
        .ColWidth(1) = 703
        .Row = 0
        .Col = 1
        .Text = "Can"
        
        .ColWidth(2) = 3703
        .Row = 0
        .Col = 2
        .ColAlignment(2) = 4
        .Text = "Detalle"
        
        .ColWidth(3) = 1033
        .Row = 0
        .Col = 3
        .Text = "Pre.U"
        
        .ColWidth(4) = 803
        .Row = 0
        .Col = 4
        .Text = "Dto."
        
        .ColWidth(5) = 1403
        .Row = 0
        .Col = 5
        .Text = "Pre.T"
        
        .ColWidth(6) = 0
        .Row = 0
        .Col = 6
        .Text = "Cod"
        
        .ColWidth(7) = 0
        .Row = 0
        .Col = 7
        .Text = "CodM"
        
        .ColWidth(8) = 0
        .Row = 0
        .Col = 8
        .Text = "IVA"
    End With
    
    'Establece el valor de la variable FILA
    Fila = 1
    
    '=================================
    '   Tabla de buscar productos
    '=================================
    'Cargar tabla de buscar con todos los productos
    tbCaja
    Set dtgBuscar.DataSource = rsCaja
    


End Sub
Public Function dibujarTabla()
'Dibujar el GRID
Filas = 1
fldCaja.Rows = Filas
fldCaja.Rows = 1
fldCaja.Cols = 9

'Ocultar la primer columna del grid
fldCaja.ColWidth(0) = 0

'Poner los ENCABEZADOS del GRID
With fldCaja
    .ColWidth(1) = 703
    .Row = 0
    .Col = 1
    .Text = "Can"
    
    .ColWidth(2) = 3703
    .Row = 0
    .Col = 2
    .ColAlignment(2) = 4
    .Text = "Detalle"
    
    .ColWidth(3) = 1033
    .Row = 0
    .Col = 3
    .Text = "Pre.U"
    
    .ColWidth(4) = 803
    .Row = 0
    .Col = 4
    .Text = "Dto."
    
    .ColWidth(5) = 1403
    .Row = 0
    .Col = 5
    .Text = "Pre.T"
    
    .ColWidth(6) = 0
    .Row = 0
    .Col = 6
    .Text = "Cod"
    
    .ColWidth(7) = 0
    .Row = 0
    .Col = 7
    .Text = "CodM"
    
    .ColWidth(8) = 0
    .Row = 0
    .Col = 8
    .Text = "IVA"
End With
End Function

'Al tener el Focus TXTCANT borra el texto
Private Sub txtCant_GotFocus()
txtCant.Text = ""
End Sub

'================================
'LLAMAR ENTER cuando estes en un  TEXTBOX
'================================
Private Sub txtCant_KeyPress(keyascii As Integer)

If keyascii = 13 Then cmdAceptar_Click
    
    If keyascii >= 44 And keyascii <= 57 Or keyascii = 8 Then
       Exit Sub
    
    Else
        keyascii = 0
    
    End If
End Sub

Private Sub txtcod_KeyPress(keyascii As Integer)
If keyascii = 99 Then
    keyascii = 0
    Check1_Click
    Exit Sub
End If
If keyascii = 13 Then cmdAceptar_Click 'Else Exit Sub

If keyascii = 44 Then keyascii = 46
    
    If keyascii >= 48 And keyascii <= 57 Or keyascii = 8 Then
       Exit Sub
    Else
        keyascii = 0
    End If

If txtcod = "" Then txtcod.SetFocus: Exit Sub
End Sub
'================================

Private Sub cmdAceptar_Click()

'===================
'   Variables
'===================

Dim nv As Boolean
Dim i As Integer
Dim descrip As String
Dim preProd, iva As String
Dim codigo As String
Dim codigoM As String
Dim Yf As Integer
Dim Cc As Integer
Dim z As Integer
Dim y As Integer
Dim dtotal As Currency
                        
If txtDto.Text = "" Then
          txtDto.Text = 0
End If

If txtCant = "" Then txtCant.SetFocus: Exit Sub

If Not txtcod.Text = "" Then
            Cc = 6
            tbCaja
            With rsCaja
                    .Requery
                    .Find "idProd ='" & txtcod & "'"
                    If Not .EOF = False Then
                    .Requery
                    .Find "idProdManual ='" & txtcod & "'"
                        If Not .EOF = False Then
                            MsgBox "Este producto no esta registrado en el almacen", vbExclamation, Me.Caption
                            txtcod.Text = ""
                            txtCant.Text = "1"
                            chkDto.Value = 0
                            txtBuscar.SetFocus
                        Exit Sub
                        Else
                            Cc = 7
                        End If
                    End If
                    Dim blnChk As Boolean
                    z = 0
                    For i = 0 To fldCaja.Rows - 1
                        blnChk = False
                        fldCaja.Row = i
                        fldCaja.Col = 6
                            If !idProd = fldCaja.Text Then
                                fldCaja.Col = 1
                                z = z + CInt(fldCaja.Text)
                                
                                'Eliminar fila de la tabla
                                If chkDto.Value = 1 Then
                                    blnChk = True
                                End If
                                
                                If (CInt(fldCaja.Text) + CInt(txtCant.Text)) < 1 Then
                                    
                                    If fldCaja.Rows = 2 Then
                                        
                                        If blnChk = True Then
                                            
                                            fldCaja.Col = 4
                                            
                                            If fldCaja.Text = txtDto.Text Then
                                                fldCaja.Clear
                                                dibujarTabla
                                                Fila = 1
                                                txtCant.Text = 1
                                                txtcod.Text = ""
                                                
                                                dtotal = Sumar(fldCaja, 1)
                                                ' formatear el resultado y mostrarlo
                                                Label10 = dtotal
                                                
                                                dtotal = Sumar(fldCaja, 5)
                                                dtotal = Format(dtotal, "###,##0.00")
                                                ' formatear el resultado y mostrarlo
                                                Label11 = dtotal
                                                Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
                                                
                                                Exit Sub
                                            End If
                                            
                                            fldCaja.Col = 1
                                        Else
                                                fldCaja.Clear
                                                dibujarTabla
                                                Fila = 1
                                                txtCant.Text = 1
                                                txtcod.Text = ""
                                                
                                                dtotal = Sumar(fldCaja, 1)
                                                ' formatear el resultado y mostrarlo
                                                Label10 = dtotal
                                                
                                                dtotal = Sumar(fldCaja, 5)
                                                dtotal = Format(dtotal, "###,##0.00")
                                                ' formatear el resultado y mostrarlo
                                                Label11 = dtotal
                                                Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
                                                     
                                                Exit Sub
                                        End If
                                        
                                        dibujarTabla
                                        Fila = 1
                                        txtCant.Text = 1
                                        txtcod.Text = ""
                                        
                                        dtotal = Sumar(fldCaja, 1)
                                        ' formatear el resultado y mostrarlo
                                        Label10 = dtotal
                                        
                                        dtotal = Sumar(fldCaja, 5)
                                        dtotal = Format(dtotal, "###,##0.00")
                                        ' formatear el resultado y mostrarlo
                                        Label11 = dtotal
                                        Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
                                                                                
                                        Exit Sub
                                    
                                    End If
                                    If blnChk = True Then
                                                                            
                                        fldCaja.Col = 4
                                        
                                        If fldCaja.Text = txtDto.Text Then
                                        
                                            fldCaja.RemoveItem (i)
                                            txtCant.Text = 1
                                            txtcod.Text = ""
                                            Fila = fldCaja.Rows

                                            dtotal = Sumar(fldCaja, 1)
                                            ' formatear el resultado y mostrarlo
                                            Label10 = dtotal
                                            
                                            dtotal = Sumar(fldCaja, 5)
                                            dtotal = Format(dtotal, "###,##0.00")
                                            ' formatear el resultado y mostrarlo
                                            Label11 = dtotal
                                            Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
                                                                                        
                                            Exit Sub
                                        Else
                                         
                                            
                                        End If
                                        
                                        fldCaja.Col = 1
                                    Else
                                        fldCaja.RemoveItem (i)
                                        txtCant.Text = 1
                                        txtcod.Text = ""
                                        Fila = fldCaja.Rows
    
                                        dtotal = Sumar(fldCaja, 1)
                                        ' formatear el resultado y mostrarlo
                                        Label10 = dtotal
                                        
                                        dtotal = Sumar(fldCaja, 5)
                                        dtotal = Format(dtotal, "###,##0.00")
                                        ' formatear el resultado y mostrarlo
                                        Label11 = dtotal
                                        Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
                                                                                
                                        Exit Sub
                                    End If

                                        
                                End If
                                
                            End If
                    
                    Next i
                    If !uniProd < txtCant.Text + z Then
                            MsgBox "No hay suficientes existencias de este producto", vbCritical, Me.Caption
                            Exit Sub
                    Else
                    y = !uniProd - (txtCant.Text + z)
                        If y < 1 Then
                                MsgBox "¡Stock minimo! quedan " & CStr(y) & " unidades disponibles" & " de " & !nomProd, vbExclamation, Me.Caption
                        End If
                    End If
                        codigo = !idProd
                        descrip = !nomProd
                        If chkPrecio.Value = 1 Then
                            preProd = Format(txtPprecio.Text, "###,##0.00")
                        Else
                            preProd = Format(!pvpProd, "###,##0.00")
                        End If
                        codigoM = !idProdManual
                        iva = !ivaProd
             End With
                        With fldCaja
                Dim cmpPrecio As String
                
                For i = 0 To .Rows - 1
                        If txtcod.Text = fldCaja.TextMatrix(i, Cc) And txtDto.Text = fldCaja.TextMatrix(i, 4) And preProd = fldCaja.TextMatrix(i, 3) Then
                            .Row = i
                            .Col = 1
                            Yf = CInt(.Text) + CInt(txtCant)
                            .Text = Yf
                            .Col = 5
                            xx = ((CDec(preProd) - ((CDec(preProd) * CDec(txtDto.Text)) / 100)) * CDec(txtCant.Text)) + CDec(.Text)
                            .Text = Format(xx, "###,##0.00")
                            txtcod.Text = ""
                            txtCant.Text = "1"
                            
                            dtotal = Sumar(fldCaja, 1)
                            ' formatear el resultado y mostrarlo
                            Label10 = dtotal
                            
                            dtotal = Sumar(fldCaja, 5)
                            dtotal = Format(dtotal, "###,##0.00")
                            ' formatear el resultado y mostrarlo
                            Label11 = dtotal
                            Label11.Caption = Format(Label11.Caption, "###,##0.00")
                            'If chkDto.Value = 1 Then chkDto.Value = 0
                            txtBuscar.SetFocus
                            Exit Sub
                        End If
                Next i
            End With
            
            Filas = Filas + 1
            'fldCaja.Rows = Filas
            fldCaja.Rows = fldCaja.Rows + 1
            
            fldCaja.Col = 1
            fldCaja.Row = Fila
            y = txtCant.Text
            fldCaja.Text = y
            
            fldCaja.Col = 2
            fldCaja.Row = Fila
            fldCaja.Text = descrip
            
            fldCaja.Col = 3
            fldCaja.Row = Fila
            fldCaja.Text = preProd
            
            fldCaja.Col = 4
            fldCaja.Row = Fila
            fldCaja.Text = txtDto.Text
            
            fldCaja.Col = 5
            fldCaja.Row = Fila
            x = (preProd - ((preProd * txtDto.Text) / 100)) * txtCant.Text
            fldCaja.Text = Format(x, "###,##0.00")
            
            fldCaja.Col = 6
            fldCaja.Row = Fila
            fldCaja.Text = codigo
            
            fldCaja.Col = 7
            fldCaja.Row = Fila
            fldCaja.Text = codigoM
            
            fldCaja.Col = 8
            fldCaja.Row = Fila
            fldCaja.Text = iva
            
            dtotal = Sumar(fldCaja, 5)
            ' formatear el resultado y mostrarlo
            
            Label11.Caption = dtotal
            Label11.Caption = Format(Label11.Caption, "###,##0.00")
            dtotal = Sumar(fldCaja, 1)
            ' formatear el resultado y mostrarlo
            Label10 = dtotal
                        
            
            txtcod.Text = ""
            txtBuscar.SetFocus
            txtCant.Text = "1"
            Fila = Fila + 1
            
        Else
            Exit Sub
End If
txtBuscar.SetFocus
End Sub
Private Sub fldCaja_MouseUp _
(Button As Integer, Shift As Integer, x As Single, _
y As Single)
    If fldCaja.Col = 3 Or fldCaja.Col = 5 Or fldCaja.Col = 1 Then
        fldCaja.Text = ""
    End If
End Sub

Private Sub fldCaja_keypress(keyascii As Integer)
    
    'EDITAR GRID
    '=================
    
    If keyascii = 46 Then keyascii = 44
    
    'Editar PrecioU
    If (keyascii >= 44 And keyascii <= 57 And fldCaja.Col = 3) Or (keyascii = 8 And fldCaja.Col = 3) Then
        Dim valCant As Integer
        Dim valPrecio As Double
        'tecla borrar
        If keyascii = 8 Then
            If Len(fldCaja.Text) > 0 Then
                fldCaja.Text = Left(fldCaja.Text, Len(fldCaja.Text) - 1)
            End If
        Else
            fldCaja.Text = fldCaja.Text & Chr(keyascii)
        End If
        
        'Campo vacio o con ,
        If fldCaja.Text = "" Then
            valPrecio = 0
        ElseIf fldCaja.Text = "," Then
            fldCaja.Text = "0,"
            valPrecio = fldCaja.Text
        Else
            valPrecio = fldCaja.Text
        End If
        
        
        fldCaja.Col = 1
        valCant = fldCaja.Text
        
        fldCaja.Col = 5
        fldCaja.Text = valCant * valPrecio
        fldCaja.Text = Format(fldCaja, "###,##0.00")
        fldCaja.Col = 3
        
    End If
    
    
    'Editar PrecioT
    If (keyascii >= 44 And keyascii <= 57 And fldCaja.Col = 5) Or (keyascii = 8 And fldCaja.Col = 5) Then
        
        'tecla borrar
        If keyascii = 8 Then
            If Len(fldCaja.Text) >= 1 Then
                fldCaja.Text = Left(fldCaja.Text, Len(fldCaja.Text) - 1)
            End If
        Else
            fldCaja.Text = fldCaja.Text & Chr(keyascii)
        End If
        
        If fldCaja.Text = "" Then
            valPrecio = 0
        ElseIf fldCaja.Text = "," Then
            fldCaja.Text = "0,"
            valPrecio = fldCaja.Text
        Else
            valPrecio = fldCaja.Text
        End If
    End If
    
    'borrar fila
    If (keyascii >= 44 And keyascii <= 57 And fldCaja.Col = 1) Or (keyascii = 27 And fldCaja.Col = 1) Then
        
        
        If keyascii = 27 Then
            If fldCaja.Rows = 2 Then
                fldCaja.Clear
                dibujarTabla
                Fila = 1
            Else
                fldCaja.RemoveItem (fldCaja.RowSel)
                Fila = fldCaja.Rows
            End If
        Else
            fldCaja.Text = fldCaja.Text & Chr(keyascii)
            
            If fldCaja.Text = "" Then
                valPrecio = 0
            ElseIf fldCaja.Text = "," Then
                fldCaja.Text = "0,"
                valPrecio = fldCaja.Text
            Else
                valPrecio = fldCaja.Text
            End If
            valCant = fldCaja.Text
            fldCaja.Col = 3
            valPrecio = fldCaja.Text
            fldCaja.Col = 5
            fldCaja.Text = valPrecio * valCant
            fldCaja.Text = Format(fldCaja, "###,##0.00")
            
        End If
    End If
    dtotal = Sumar(fldCaja, 5)
    ' formatear el resultado y mostrarlo
    
    Label11.Caption = dtotal
    Label11.Caption = Format(Label11.Caption, "###,##0.00") & "€"
    
    dtotal = Sumar(fldCaja, 1)
    ' formatear el resultado y mostrarlo
    
    Label10.Caption = dtotal
    
    
    
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
Dim intboton As String

intboton = MsgBox("¿Desea salir de  caja?", vbQuestion Or vbYesNo, Me.Caption)
If intboton = vbYes Then
                    MDIfrmMadre.Show
                    Unload Me
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    bControlAcceso = False
    MDIfrmMadre.Picture1.Visible = True
End Sub



Function Sumar(MSFlexGrid As Object, _
               Columna As Integer) As Currency
               
On Error GoTo error_function

    With MSFlexGrid
        Dim Total As Currency
        Dim i As Long
        
        If Columna > .Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
        
        ' recorrer  las filas de la grilla
        For i = 1 To .Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(.TextMatrix(i, Columna)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                Total = Total + .TextMatrix(i, Columna)
            End If
        Next
        
        ' retornar el total de la suma a la función
        Sumar = Total
    End With
    
Exit Function

error_function:
MsgBox Err.Description, vbCritical, "error al sumar"

End Function


Private Sub txtDto_Change()
    Label13.Caption = "Descuento activado del " & txtDto.Text & "%"
End Sub

Private Sub txtDto_KeyPress(keyascii As Integer)
    If keyascii >= 48 And keyascii <= 57 Or keyascii = 8 Then
       Exit Sub
    Else
        keyascii = 0
    End If
End Sub

Private Sub txtPprecio_Change()
    Label13.Caption = "Personalizar precio a " & txtPprecio.Text & " €"
End Sub

Private Sub txtPprecio_KeyPress(keyascii As Integer)
If keyascii = 46 Then keyascii = 44
    If keyascii >= 44 And keyascii <= 57 Or keyascii = 8 Then
       Exit Sub
    
    Else
        keyascii = 0
    
    End If
End Sub
