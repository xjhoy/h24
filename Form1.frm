VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   9360
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim x, Y, z, a, b, c As String
x = Text1.Text
Y = Text2.Text
z = "0,"
a = z + Y
b = CDbl(x) * CDbl(a)

c = CDbl(x) - CDbl(b)

Label1 = x
Label2 = Y
Label3 = c
 


End Sub

Private Sub Form_Load()


Set DataGrid1.DataSource = dtEntorno.rscmdTicket
Dim Palabras, Caracteres, MiCadena
For Palabras = 10 To 1 Step -1    ' Establece 10 repeticiones.
   For Caracteres = 0 To 9   ' Establece 10 repeticiones.
      MiCadena = MiCadena & Caracteres   ' Agrega un número a la cadena.
   Next Caracteres   ' Incrementa el contador
   MiCadena = MiCadena & " "   ' Agrega un espacio.
Next Palabras
'Label1.Caption = MiCadena


End Sub

Public Sub botones()

    Dim boton As Object
    Set boton = Controls.Add("VB.commandbutton", "Boton(contador)")
    boton.Visible = True
    boton.Caption = "Drinky94 ^^"
    boton.Width = 1250
    boton.Height = 450
    boton.Left = 150
    boton.Top = 900
End Sub

Public Sub calNeto()
Dim dat1 As Integer
Dim dat2 As Integer

Dim cal As Integer

dat1 = Text1.Text

dat2 = 0# & Text2.Text

cal = dat1 + dat2

End Sub
