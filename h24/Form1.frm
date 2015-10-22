VERSION 5.00
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   5055
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Palabras, Caracteres, MiCadena
For Palabras = 10 To 1 Step -1    ' Establece 10 repeticiones.
   For Caracteres = 0 To 9   ' Establece 10 repeticiones.
      MiCadena = MiCadena & Caracteres   ' Agrega un número a la cadena.
   Next Caracteres   ' Incrementa el contador
   MiCadena = MiCadena & " "   ' Agrega un espacio.
Next Palabras
Label1.Caption = MiCadena


End Sub

Public Sub botones()

    Dim boton As Object
    Set boton = Controls.Add("VB.commandbutton", "Boton")
    boton.Visible = True
    boton.Caption = "Drinky94 ^^"
    boton.Width = 1250
    boton.Height = 450
    boton.Left = 150
    boton.Top = 900
End Sub

