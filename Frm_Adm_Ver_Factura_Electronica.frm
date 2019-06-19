VERSION 5.00
Begin VB.Form Frm_Adm_Ver_Factura_Electronica 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VISUALIZADOR DE ARCHIVOS"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   14220
   Begin VB.CommandButton Btn_Salir 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   555
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Archivos 
      BackColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
      Begin VB.OLE OLE_Archivos 
         BackColor       =   &H00FFFFFF&
         Height          =   8655
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   13695
      End
   End
End
Attribute VB_Name = "Frm_Adm_Ver_Factura_Electronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Mostrar_Archivo_Pantalla()
Dim Extension As String

On Error GoTo MuestraError
        'Revisa la extension del archivo para mostrarlo
        Extension = Right(Ruta_Archivo_Mostrar, 3)
        OLE_Archivos.Visible = False
        AcroPDF1.Visible = False
        OLE_Archivos.SourceDoc = Ruta_Archivo_Mostrar
        OLE_Archivos.SourceItem = Ruta_Archivo_Mostrar
        If UCase(Extension) = "DOC" Then
            Fra_Archivos.Visible = True
            OLE_Archivos.Visible = True
            OLE_Archivos.Class = "Word.Document.8"
        End If
        If UCase(Extension) = "XLS" Then
            OLE_Archivos.Visible = True
            OLE_Archivos.Class = "Excel.Chart.8"
        End If
        If UCase(Extension) = "PDF" Then
            Fra_Archivos.Visible = True
            AcroPDF1.Visible = True
            AcroPDF1.LoadFile (Ruta_Archivo_Mostrar)
            'Fra_Archivos.ZOrder
        End If
''        If UCase(Extension) = "JPG" Then
''            'Pdf_Archivos.Visible = False
''            Web_Archivos.Visible = False
''            OLE_Archivos.Visible = True
''            OLE_Archivos.Class = "MSPhotoEd.3"
''        End If
''        If UCase(Extension) = "PPT" Then
''            'Pdf_Archivos.Visible = False
''            Web_Archivos.Visible = False
''            OLE_Archivos.Visible = True
''            OLE_Archivos.Class = "PowerPoint.Show.8"
''        End If
        If OLE_Archivos.Visible = True Then
            OLE_Archivos.Action = 1
            OLE_Archivos.SizeMode = 0
        End If
    Exit Sub
MuestraError:
    MsgBox Err.Description
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 10260
    Me.Width = 14340
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
