VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTurnoDetalleT 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Detalle de Factura simple"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   15840
      Picture         =   "frmTurnoDetalleT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTurnoDetalleT.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8895
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   15690
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "idTProd"
            Caption         =   "id"
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
            DataField       =   "idProd"
            Caption         =   "Código de barras"
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
         BeginProperty Column03 
            DataField       =   "Tuni"
            Caption         =   "Cant."
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
            DataField       =   "pvpProd"
            Caption         =   "Precio Unit."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00 €"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ivaProd"
            Caption         =   "IVA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0 ""%"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "dtoProd"
            Caption         =   "Dto."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0 ""%"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "netoProd"
            Caption         =   "Neto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00 €"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "PrecioF"
            Caption         =   "Precio Final"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00 €"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "idtbTicket"
            Caption         =   "Nº Fact. S"
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
               ColumnWidth     =   1335,118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1709,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2684,977
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   4560
      Picture         =   "frmTurnoDetalleT.frx":1CE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE PRODUCTOS DE LA FACTURA SIMPLE"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
End
Attribute VB_Name = "frmTurnoDetalleT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Des As Byte = 10
Const Ca As Byte = 3
Const Pre As Byte = 5
Const TPre As Byte = 5


Private Sub cmdCerrar_Click()
frmTurnoDetalle.Show
Unload Me
End Sub

Private Sub cmdImprimir_Click()

'NOTAS:
'Hay que hacer nuevamente el codigo para imprimir, ya que
'el codigo actual es para hacerlo desde un MSFGrid Y no del DataGrid actual.


'***************************
'* Variables para calcular
'***************************
Dim idTicketFS As String
Dim r As String * Des
Dim s As String * Ca
Dim v As String * Pre
Dim TotalV As String
Dim b As String
Dim dato As String
Dim datoB As String
Dim i, u As Long
Dim a, e As Long
Dim Lcalc As Integer

'***************************
'* Variables de impresora
'***************************

Dim canal%
Dim Impresora

'Llamar a la impresora predeterminada y compartida en red
Impresora = "\\127.0.0.1\" & Printer.DeviceName
canal = FreeFile

'Abrimos el fichero donde vamos a escribir
Open Impresora For Output As #canal

Print #canal, Chr$(&H1B); "@"; 'INICIO (ESC @)
Print #canal, Chr$(&H1B); "a"; Chr$(1); 'CENTRADO
Print #canal, Chr$(&H1B); "t"; Chr$(19); 'Caracteres en español
Print #canal, Chr$(&H1B); "!"; Chr$(17); 'TAMAÑO LETRA (ESC !)

Print #canal, "Tienda STITCH"; Chr$(&HA);
Print #canal, Chr$(&H1B); "!"; Chr$(1); 'TAMAÑO LETRA (ESC !)

Print #canal, "Pan y Caprichos";
Print #canal, Chr$(&H1B); "d"; Chr$(2); '3 SALTOS DE LINEA(ESC d)

Print #canal, "C/ Cristobal Sanz Nº 86"; Chr$(&HA);
Print #canal, "CP. 03201"; Chr$(&HA);
Print #canal, "Elche - Alicante"; Chr$(&HA);
Print #canal, "Tlf 966109608"; Chr$(&HA);
Print #canal, "CIF 74246484T"; Chr$(&HA);
Print #canal, Chr$(&HA); 'SALTO DE LINA

'==============================
tbTicket
With rsUsuarios

    .Find "idtbTicket = '" & idTurnoD & "'"
        '==============================
        TotalV = !VentaT
        Print #canal, "Fecha: "; !TFecha; Chr$(&HA);
        Print #canal, "Hora: "; !Thora; Chr$(&HA);
        Print #canal, Chr$(&H1B); "a"; Chr$(0); 'IZQUIERDA
        Print #canal, Chr$(&HA); 'SALTO DE LINEA
        Print #canal, "Nº Factura simple:  "; !idtbTicket; Chr$(&HA);

End With

Print #canal, "------------------------------"; Chr$(&HA);

tbTurnoDetalleT (idTurnoD)

Dim sDato As String
Dim sDato2 As String
'Llamar a recortset

With rsTurnoDetalleT
    If DataGrid1.VisibleRows <> 0 Then
'==============================
'Imprimir Cabecera
        Dim hCan As String * Ca
        Dim hIVA As String * Ca
        Dim hDet As String * Des
        Dim hpvp As String * Pre
        Dim hPreF As String * TPre
        Dim hDato As String
        Dim sIVA, sPvp, sPreF As String
        Dim idT As String
        
        idT = !idtbTicket
        hCan = "Can"
        hDet = "Detalle"
        hpvp = "Pre.U"
        hPreF = "Pre.T"
        
        
        hpvp = alignIzq("Pre.U", hpvp, 5)
        hPreF = alignIzq("Pre.T", hPreF, 5)
        
        
        hDato = hCan + " " + hDet + " " + hpvp + " " + hPreF
        
        Print #canal, hDato; Chr$(&HA);
        
'================================
Print #canal, "------------------------------"; Chr$(&HA);

'=================================
                .MoveFirst
            For i = 1 To DataGrid1.VisibleRows
                hDet = !nomProd
                hCan = !Tuni
                hpvp = !pvpProd
                hPreF = !PrecioF
            
                sPvp = Format(!pvpProd, "###,##0.00")
                sPreF = Format(!PrecioF, "###,##0.00")
                
                hpvp = alignIzq(sPvp, hpvp, 5)
                hPreF = alignIzq(sPreF, hPreF, 5)
                
                sDato = hCan + " " + hDet + " " + hpvp + " " + hPreF
                    
                Print #canal, sDato; Chr$(&HA);
                
                If .EOF = False Then
                    .MoveNext
                End If
            
            Next i
        

'==============================================
Print #canal, "------------------------------"; Chr$(&HA);

Print #canal, Chr$(&H1B); "a"; Chr$(1); 'CENTRADO
Print #canal, Chr$(&H1B); "!"; Chr$(17); 'TAMAÑO DE LETRA
TotalV = Format(TotalV, "###,##0.00")
Print #canal, "TOTAL: "; TotalV; Chr$(&HA);
'=======================================
Print #canal, Chr$(&H1B); "d"; Chr$(2); '3 SALTOS DE LINEA(ESC d)

Print #canal, Chr$(&H1B); "a"; Chr$(0); 'Izquierda
Print #canal, Chr$(&H1B); "!"; Chr$(1); 'TAMAÑO DE LETRA
Print #canal, "------------------------------"; Chr$(&HA);

'==================================
tbIva (idTurnoD)
        With rsIVA
            'Cabecera de IVA desglosado
            hCan = "Can"
            hIVA = "IVA"
            hpvp = "B.IMP"
            hPreF = "Cuota"
            
            sDato = hCan & " " & hIVA & " " & hpvp & " " & hPreF
            
            Print #canal, sDato; Chr$(&HA);

            'Mostrar el IVA desglosado
            .MoveFirst
            For i = 1 To .RecordCount
                
                hCan = !Cant
                sIVA = !ivaProd
                
                sIVA = Format(sIVA, "###,##0.00") & "%"
                
                hIVA = alignIzq(sIVA, hIVA, 3)
                
                sPvp = !sNeto
                sPvp = Format(sPvp, "###,##0.00")
                
                hpvp = alignIzq(sPvp, hpvp, 5)
                
                sPreF = !r
                sPreF = Format(sPreF, "###,##0.00")
                
                hPreF = alignIzq(sPreF, hPreF, 5)
                
                sDato = hCan & " " & hIVA & " " & hpvp & " " & hPreF
                
                Print #canal, sDato; Chr$(&HA);
                
                If .EOF = False Then
                    .MoveNext
                End If
            Next i
            
            
          Dim Riva, RivaA, RivaB, RivaC, RBiva, BivaA, BivaB, BivaC As Currency

          .MoveFirst
          If .EOF = False Then
                     BivaA = !sNeto
                     RivaA = !r
                    .MoveNext
          Else
                    BivaA = 0
                    RivaA = 0
          End If
          If .EOF = False Then
                     BivaB = !sNeto
                     RivaB = !r
                    .MoveNext
          Else
                    BivaB = 0
                    RivaB = 0
          End If
          If .EOF = False Then
                    BivaC = !sNeto
                    RivaC = !r
          Else
                    BivaC = 0
                    RivaC = 0
          End If

          RBiva = BivaA + BivaB + BivaC
          Riva = RivaA + RivaB + RivaC
          
          RBiva = Format(RBiva, "###,##0.00")
          
          hpvp = alignIzq(RBiva, hpvp, 5)
          
          
          Riva = Format(Riva, "###,##0.00")
            
          hPreF = alignIzq(Riva, hPreF, 5)
          
          Print #canal, "TOTAL" & Space(3) & hpvp & " " & hPreF
                    
        End With
    End If
End With
Print #canal, Chr$(&H1D); "V"; Chr$(66); Chr$(0); 'Feeds paper & cut

Print #canal, Chr$(&H1B); Chr$(&H70); Chr$(&H0); Chr$(60); Chr$(120);
Print #canal, Chr$(&H1B); "d"; Chr$(15); '3 SALTOS DE LINEA(ESC d)

'Cerrar fichero
Close #canal


End Sub
Public Function alignIzq(ByVal s, s2 As String, ByVal i As Integer) As String

Dim esp As Integer

If Len(s) < i Then
   esp = i - Len(s)
   s2 = Space(esp) & s
   alignIzq = s2
Else
    s2 = s
    alignIzq = s2
End If

End Function
Private Sub Form_Load()
SSTab1.Caption = ""
tbTurnoDetalleT (idTurnoD)
Set DataGrid1.DataSource = rsTurnoDetalleT
End Sub
Private Sub Form_Unload(Cancel As Integer)
With rsTurnoDetalleT
    If .State = 1 Then .Close
End With
End Sub

