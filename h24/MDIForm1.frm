VERSION 5.00
Begin VB.MDIForm MDIfrmMadre 
   BackColor       =   &H8000000C&
   Caption         =   "Stich"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAlmacen 
      Caption         =   "Almacen"
      Begin VB.Menu mnuAlmacenNuevoProducto 
         Caption         =   "Nuevo producto"
      End
      Begin VB.Menu mnuAlmacenListaProductos 
         Caption         =   "Listado de productos"
      End
      Begin VB.Menu mnuAlmacenSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlmacenBuscar 
         Caption         =   "Buscar producto"
      End
   End
   Begin VB.Menu mnuProv 
      Caption         =   "Proveedores"
      Begin VB.Menu mnuProvAdd 
         Caption         =   "Nuevo proveedor"
      End
      Begin VB.Menu mnuProvList 
         Caption         =   "Lista de proveedores"
      End
      Begin VB.Menu mnuProvSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProvBuscar 
         Caption         =   "Buscar Proveedor"
      End
   End
End
Attribute VB_Name = "MDIfrmMadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAlmacenListaProductos_Click()
    If bControlAcceso = False Then
        bControlAcceso = True
    frmListaProductos.Show
    End If

End Sub

Private Sub mnuAlmacenNuevoProducto_Click()
    If bControlAcceso = False Then
        bControlAcceso = True
frmNuevoProducto.Show
    End If

End Sub

Private Sub mnuProvAdd_Click()

    If bControlAcceso = False Then
        bControlAcceso = True
        frmProvAdd.Show
    End If


End Sub

Private Sub mnuProvList_Click()
    If bControlAcceso = False Then
        bControlAcceso = True
        frmListProv.Show
    End If
End Sub
