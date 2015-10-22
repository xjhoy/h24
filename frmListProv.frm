VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListProv 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Lista de proveedores"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7800
      Picture         =   "frmListProv.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   10560
      Width           =   240
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   13920
      Picture         =   "frmListProv.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10800
      Picture         =   "frmListProv.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7560
      Picture         =   "frmListProv.frx":109E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   240
   End
   Begin VB.CommandButton cmdDetalleProv 
      Caption         =   " Ver proveedor"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdCerrar 
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
      Height          =   1215
      Left            =   16080
      Picture         =   "frmListProv.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar proveedor"
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
      Left            =   11280
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdModProv 
      Caption         =   "Modificar proveedor"
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
      Left            =   8160
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdProvAdd 
      Caption         =   "Nuevo Proveedor"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   -2147483633
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmListProv.frx":32F2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgProv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtgProv 
         Height          =   6495
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   11456
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "nomProv"
            Caption         =   "Razón social"
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
            DataField       =   "cifProv"
            Caption         =   "cif"
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
            DataField       =   "provinciaProv"
            Caption         =   "prov"
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
            DataField       =   "locProv"
            Caption         =   "loc"
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
         BeginProperty Column04 
            DataField       =   "dirProv"
            Caption         =   "dir"
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
         BeginProperty Column05 
            DataField       =   "cpProv"
            Caption         =   "cp"
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
         BeginProperty Column06 
            DataField       =   "tlfProv"
            Caption         =   "Teléfono"
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
         BeginProperty Column07 
            DataField       =   "emailProv"
            Caption         =   "Email"
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
         BeginProperty Column08 
            DataField       =   "contProv"
            Caption         =   "Persona de contacto"
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
               ColumnWidth     =   3839,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3014,929
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3764,977
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
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
      Left            =   8880
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   9360
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   5640
      Picture         =   "frmListProv.frx":330E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "frmListProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBorrar_Click()
Dim intboton As Integer
Dim intboton2 As Integer
'On Error GoTo error

'If Not selectId = "" Then
If rsUsuarios.BOF Or rsUsuarios.EOF Then Exit Sub
          intboton = MsgBox("¿Desea borrar el proveedor?" & vbNewLine & "Razón social: " & dtgProv.Columns(0).Text & vbNewLine & "Teléfono: " & dtgProv.Columns(6).Text & vbNewLine & "Persona de contacto: " & dtgProv.Columns(8).Text & vbNewLine & "¡Esto también borrará el listado de productos de este proveedor!", vbQuestion Or vbYesNo, Me.Caption)
          If intboton = vbYes Then
                    base.Execute ("DELETE * FROM tbProdProv WHERE cifProv ='" & selectId & "'")
                    If (rsUsuarios.EOF Or rsUsuarios.BOF) Then
                              Exit Sub
                    Else
                              With rsUsuarios
                                        .Delete
                                        If .BOF = False Then .MoveNext
                              'End With
                                        base.Execute ("DELETE * FROM tbProv Where cifProv ='" & selectId & "'")
                              'With rsUsuarios
                                        '.MoveFirst
                                        '.Requery
                                        selectId = ""
                              End With
                    End If
          End If
'Else
'error:
'         intboton2 = MsgBox("Seleccione un proveedor", vbExclamation, Me.Caption)
'End If

End Sub

Private Sub cmdCerrar_Click()
MDIfrmMadre.Picture1.Visible = True
bControlAcceso = False
Unload Me
End Sub

Private Sub cmdModProv_Click()
Dim intboton2 As String

sOperacion = "B"
If selectId = "" Then
          intboton2 = MsgBox("Seleccione un proveedor", vbExclamation, Me.Caption)
Else
          frmProvAdd.Show
End If

End Sub

Private Sub cmdProvAdd_Click()
        sOperacion = "A"
        frmProvAdd.Show
        'Unload Me
End Sub

Private Sub cmdDetalleProv_Click()
On Error GoTo error
With dtgConn
          If .State = 1 Then
                    .Close
                    tbProv
          End If
End With
If Not selectId = "" Then
          SelectDetalleProv = selectId
          bControlAcceso = False
          frmProvDetalle.Show
          Unload Me
Else
error:
          MsgBox "Antes debe seleccionar un proveedor", vbExclamation, Me.Caption
End If

End Sub

Private Sub dtgProv_Click()
With dtgConn
          If .State = 1 Then
                    If .BOF Or .EOF Then
                              Exit Sub
                    End If
                    
                    .Find "cifProv = '" & (dtgProv.Columns(1).Text) & "'"
                    selectId = dtgProv.Columns(1).Text
                    
                    If KeyAscii = 13 Then
                              cmdDetalleProv_Click
                    Else
                              Exit Sub
                    End If
                    
                    Exit Sub
          End If
          End With
          With rsUsuarios
          'On Error GoTo error
          If .BOF Or .EOF Then
                    Exit Sub
          End If
          .Find "cifProv = '" & (dtgProv.Columns(1).Text) & "'"
          selectId = dtgProv.Columns(1).Text
          
          If KeyAscii = 13 Then
                    cmdDetalleProv_Click
          Else
                    Exit Sub
          End If
          'error:
End With

End Sub

'Organizar dtg
'=========

Private Sub dtgProv_HeadClick(ByVal ColIndex As Integer)
campo = ColIndex
If campo = "0" Then
          If sIdS = "idProd ASC" Then
                    campo = "nomProv ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "nomProv DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "7" Then
          If sIdS = "idProd ASC" Then
                    campo = "emailProv ASC"
                    sIdS = "idProd DESC"
          Else
                    campo = "emailProv DESC"
                    sIdS = "idProd ASC"
          End If
ElseIf campo = "8" Then
          campo = "contProv"
Else
          Exit Sub
End If
'rsUsuarios.Close
dtgConnSort
dtgConn.Sort = campo
Set dtgProv.DataSource = dtgConn
If dtgConn.BOF = False Then
          dtgConn.MoveFirst
End If
'selectId = ""
'dtgProv.ReBind
End Sub

Private Sub dtgProv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then cmdBorrar_Click ' Else Exit Sub
End Sub

Private Sub dtgProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdDetalleProv_Click ' Else Exit Sub
End Sub

Private Sub Form_Activate()
SSTab1.Caption = ""
tbProv
Set dtgProv.DataSource = rsUsuarios
'dtgProv
End Sub

Private Sub Form_Load()
Dim sig As String

bControlAcceso = True
tbProv
End Sub

Private Sub Form_Unload(Cancel As Integer)
SelectDetalleProv = selectId
selectId = ""
With dtgConn
    If .State = 1 Then .Close
End With
With rsUsuarios
    If .State = 1 Then .Close
End With
End Sub







