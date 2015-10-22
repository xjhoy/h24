VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIfrmMadre 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stitch"
   ClientHeight    =   5790
   ClientLeft      =   225
   ClientTop       =   255
   ClientWidth     =   11280
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   5295
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4815
            Text            =   "Tienda Stitch - Pan & Caprichos"
            TextSave        =   "Tienda Stitch - Pan & Caprichos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4815
            TextSave        =   "20/07/2015"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4815
            TextSave        =   "22:16"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4815
            Text            =   "Inicio turno"
            TextSave        =   "Inicio turno"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   3069
      ButtonWidth     =   4286
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Administrar"
            Key             =   "ADMIN"
            Description     =   "D"
            ImageIndex      =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-------------------------------"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Caja"
            Key             =   "CAJA"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Iniciar turno"
            Key             =   "INITURNO"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Proveedores"
            Key             =   "Prov"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Almacen"
            Key             =   "ALMACEN"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Registros de turnos"
            Key             =   "REGTURNOS"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "SALIR"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1658
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":218C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2726
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   11250
      TabIndex        =   1
      Top             =   1740
      Width           =   11280
   End
End
Attribute VB_Name = "MDIfrmMadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'codigo para poner una imagen dentro de un MDI, aunque el fallo es que
'esta imagen queda por encima de todos los MDIhijos e impide que se puedan ver
'_____________________________________________________________________________________
'Variable de tipo Ipicturedisp para cargar la imagen
Dim imagen As IPictureDisp

Private Sub MDIForm_Load()
tbTurnos
With rsTurnos
          .Requery
          If .BOF Or .EOF Then
                    Else
                    .MoveFirst
                    If IsNull(!TFechaFin) Then
                              chkturno = True
                              Toolbar1.Buttons.Item(5).Caption = "Cerrar turno"
                              timerInicioTurno = !iniTurno
                              fechaInicioTurno = !TFechaIni
                              idTurnos = !id
                    Else
                              chkturno = False
                    End If
          End If

If chkturno = True Then
          Toolbar1.Buttons(3).Enabled = True
Else
          Toolbar1.Buttons(3).Enabled = False
End If

Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(9).Enabled = False


If chkadmin = False Then
          Toolbar1.Buttons(1).Caption = "Administrar"
          
          'Inhabilitar los botones de almacen y proveedor hasta que se inicie sesion
          Toolbar1.Buttons(7).Enabled = False
          Toolbar1.Buttons(9).Enabled = False
Else
          Toolbar1.Buttons(1).Caption = "Cerrar sesion"
          
          'Inhabilitar los botones de almacen y proveedor hasta que se inicie sesion
          Toolbar1.Buttons(7).Enabled = True
          Toolbar1.Buttons(9).Enabled = True

End If

If chkturno = True Then
    Toolbar1.Buttons.Item(5).Caption = "Cerrar turno"
    StatusBar1.Panels.Item(4).Text = "Inicio de turno: " & !iniTurno & "   " & !TFechaIni
Else
    Toolbar1.Buttons.Item(5).Caption = "Iniciar turno"
    StatusBar1.Panels.Item(4).Text = "Iniciar turno"
End If

'Cargamos la imagen en la variable con LoadPicture pasandole la ruta
Set imagen = LoadPicture(App.Path & "\img\img.jpg")

'Para que al repintar la ventana _
 se mantenga el gráfico
Picture1.AutoRedraw = True
Picture1.BorderStyle = 0
End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Elimina de la memoria la imagen
If Not imagen Is Nothing Then
   Set imagen = Nothing
End If

End Sub

'evento Resize del picturebox
Private Sub Picture1_Resize()

Dim Ancho As Single
Dim Alto As Single
    
    With Picture1

        Picture1.Cls
        
        'Mazimizamos el Picture1 dentro del Mdi para que ocupe todo el área
        .Move 0, 600, Me.Width, Me.Height

        'Pasamos el ancho y el alto de la imágen de Himetric a la escala _
         que tenga el picture

        Ancho = Picture1.ScaleX(imagen.Width, vbHimetric, .ScaleMode)
        Alto = Picture1.ScaleY(imagen.Height, vbHimetric, .ScaleMode)

        ' Dibujamos la imágen y la centramos
        Picture1.PaintPicture imagen, (Picture1.ScaleWidth - Ancho) / 2, _
                                  (Picture1.ScaleHeight - Alto) / 2
    
    End With
End Sub
    

Private Sub MDIForm_Resize()
'Cuando se redimensiona el Form, _
 ejecutamos el REsize del Picture

Picture1_Resize

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim boton As String
Dim iniTurnos As String


Select Case Button.Key
Case "ADMIN"

If bControlAcceso = False Then
        'bControlAcceso = True
        'Picture1.Visible = False

    If chkadmin = False Then
 
        frmIniAdmin.Show

    Else
 
        boton = MsgBox("Desea cerrar sesión", vbQuestion Or vbYesNo, Me.Caption)
        If boton = vbYes Then
            chkadmin = False
            Toolbar1.Buttons(1).Caption = "Administrar"
 
             'Inhabilitar los botones de almacen y proveedor hasta que se inicie sesion
            Toolbar1.Buttons(7).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
        End If
    End If
End If

Case "INITURNO"

Dim msgTurnos As String

actHora = Format(Now, "HH:MM:SS")

If chkturno = False Then
          chkturno = True
          Toolbar1.Buttons(3).Enabled = True
          Toolbar1.Buttons.Item(5).Caption = "Cerrar turno"

          If timerInicioTurno = "" Then
                    timerInicioTurno = Format(Now, "HH:MM:SS")
          End If
          
          If fechaInicioTurno = "" Then
                fechaInicioTurno = Date
          End If
          
          tbTurnos
          With rsTurnos
          .Requery
                    
                    If .BOF = False Then
                              
                              If .EOF = False Then
                                        .MoveFirst
                              End If
                              idTurnos = !id + 1
                    Else
                              idTurnos = 1
                    End If
                    
                    .AddNew
                              !id = idTurnos
                    'Fecha inicio turno
                    fechaInicioTurno = Date
                              !TFechaIni = CStr(fechaInicioTurno)
                    iniTurnos = CStr(timerInicioTurno)
                              !iniTurno = iniTurnos
                    .Update
                    .Requery
          End With
          
          chkCaja = True
          StatusBar1.Panels.Item(4).Text = "Inicio de turno: " & timerInicioTurno & "   " & fechaInicioTurno

Else 'chkturno = True
          msgTurnos = MsgBox("Desea Terminar turno", vbQuestion Or vbYesNo, Me.Caption)
          If msgTurnos = vbYes Then
                    Unload frmCaja
                    tbTurnos
                    With rsTurnos
                              .Requery
                              .MoveFirst
                    fechaInicioTurno = Date
                              !TFechaFin = CStr(fechaInicioTurno)
                              !finTurno = actHora
                              .UpdateBatch
                    End With
                    
                    tbVentaT
                    With rsVentaT
                              If IsNull(!Expr2) Then
                                        Tventa = 0
                              Else
                                        Tventa = !Expr2
                              End If
                    End With
                    
                    tbTurnos
                    With rsTurnos
                              .Requery
                              .MoveFirst
                                     !Tvendido = CDec(Tventa)
                              .UpdateBatch
                              .Requery
                    End With
                    frmTurnos.DataGrid1.Refresh
                    
                    Toolbar1.Buttons.Item(5).Caption = "Iniciar turno"
                    StatusBar1.Panels.Item(4).Text = "Iniciar turno"
                    timerInicioTurno = ""
                    chkturno = False
                    chkCaja = False
                    Toolbar1.Buttons(3).Enabled = False
          End If
          
End If

Case "ALMACEN"

    If bControlAcceso = False Then
        bControlAcceso = True
        Picture1.Visible = False
        frmListaProductos.Show
    End If

Case "Prov"
    
    If bControlAcceso = False Then
        bControlAcceso = True
        Picture1.Visible = False
        frmListProv.Show
    End If

Case "SALIR"

Dim boton2 As String

If bControlAcceso = False Then
    If chkadmin = True Then
        boton2 = MsgBox("Desea cerrar sesión", vbQuestion Or vbYesNo, Me.Caption)
        If boton2 = vbYes Then
            chkadmin = False
            Unload Me
        Else
            Exit Sub
        End If
        Unload Me
    End If
Unload Me
End If

Case "REGTURNOS"

    If bControlAcceso = False Then
        bControlAcceso = True
        Picture1.Visible = False
        frmTurnos.Show
    End If

Case "CAJA"

If chkturno = True Then
                    frmCaja.Show

End If

End Select

End Sub


