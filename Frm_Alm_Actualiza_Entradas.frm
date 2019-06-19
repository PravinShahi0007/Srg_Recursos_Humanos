VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Alm_Actualiza_Entradas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACTUALIZA COSTO DE ENTRADAS"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   13200
   Begin VB.Frame Fra_Detalles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ubicación de Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6210
      Left            =   105
      TabIndex        =   10
      Top             =   1215
      Width           =   13005
      Begin VB.Frame Fra_Cambio_Ubicacion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambio de Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   2430
         TabIndex        =   12
         Top             =   1470
         Visible         =   0   'False
         Width           =   8700
         Begin VB.TextBox Txt_Costo_Pesos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1530
            Width           =   915
         End
         Begin VB.TextBox Txt_Costo_Sin_IVA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5490
            TabIndex        =   19
            Top             =   1530
            Width           =   1305
         End
         Begin VB.TextBox Txt_Tipo_Cambio 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3945
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1537
            Width           =   525
         End
         Begin VB.ComboBox Cmb_Moneda 
            Height          =   315
            ItemData        =   "Frm_Alm_Actualiza_Entradas.frx":0000
            Left            =   2700
            List            =   "Frm_Alm_Actualiza_Entradas.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1530
            Width           =   1245
         End
         Begin VB.CommandButton Btn_Regresar 
            Caption         =   "Regresar"
            Height          =   330
            Left            =   5775
            TabIndex        =   9
            Top             =   2145
            Width           =   1335
         End
         Begin VB.CommandButton Btn_Cambiar 
            Caption         =   "Cambiar"
            Height          =   330
            Left            =   1740
            TabIndex        =   8
            Top             =   2145
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Factura 
            Height          =   315
            Left            =   195
            TabIndex        =   23
            Top             =   1530
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy"
            Format          =   58392579
            CurrentDate     =   39136
            MaxDate         =   402133
            MinDate         =   2
         End
         Begin VB.Label Lbl_Unidad_Frame 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7380
            TabIndex        =   25
            Top             =   15
            Width           =   615
         End
         Begin VB.Label Lbl_Unidad 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6990
            TabIndex        =   24
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label Lbl_Conversion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "En Pesos"
            Height          =   195
            Left            =   6885
            TabIndex        =   22
            Top             =   1590
            Width           =   675
         End
         Begin VB.Label Lbl_Costo_Sin_IVA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Costo Unit."
            Height          =   195
            Left            =   4635
            TabIndex        =   20
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label Lbl_Moneda 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1875
            TabIndex        =   18
            Top             =   1590
            Width           =   690
         End
         Begin VB.Label Lbl_Ubicacion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DESCRIPCION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   667
            Width           =   7170
         End
         Begin VB.Label Lbl_Nuevo_Dato 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "COSTO NUEVO"
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   8700
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Label Lbl_Descripcion_Producto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PRODUCTO:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   285
            Width           =   6810
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Productos 
         Height          =   5850
         Left            =   75
         TabIndex        =   7
         Top             =   225
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   10319
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Fra_Busqueda_Entradas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Búsqueda de Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   105
      TabIndex        =   0
      Top             =   450
      Width           =   13005
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   11880
         Picture         =   "Frm_Alm_Actualiza_Entradas.frx":0031
         TabIndex        =   6
         Tag             =   "A"
         Top             =   285
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.ComboBox Cmb_Tipo_Producto 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frm_Alm_Actualiza_Entradas.frx":0133
         Left            =   1575
         List            =   "Frm_Alm_Actualiza_Entradas.frx":013D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   2310
      End
      Begin VB.CheckBox Chk_Tipo_Producto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Producto"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   1320
      End
      Begin VB.ComboBox Cmb_Producto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5250
         TabIndex        =   4
         Top             =   285
         Width           =   5490
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   10830
         Picture         =   "Frm_Alm_Actualiza_Entradas.frx":0164
         TabIndex        =   5
         Tag             =   "A"
         Top             =   285
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CheckBox Chk_Producto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Material"
         Height          =   285
         Left            =   4065
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Label Lbl_Titulo_Entradas 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUALIZA COSTO DE ENTRADAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Frm_Alm_Actualiza_Entradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Columna_Orden_Seleccionada As Integer
Dim Modo_Orden_Grid_Productos As Integer

Private Sub Btn_Buscar_Click()
    Consulta_Entradas
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Consulta_Entradas
'DESCRIPCIÓN: Consulta las entradas
'PARÁMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 03-Octubre-2008
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Entradas()
Dim Mi_SQL As String
Dim Rs_Consulta_Detalles_Entradas As rdoResultset
Dim Empresa As String

    'Valida que inventario va a consultar
    If Cmb_Tipo_Producto.ListIndex = 0 Then
        'Consulta el inventario de materia prima
        Mi_SQL = "SELECT Alm_Entradas_Materia_Prima_Detalles.*,Cat_Materias_Primas.Nombre AS Materia_Prima,Cat_Materias_Primas.Codigo AS Codigo_Producto,Cat_Unidades.Nombre AS Unidad"
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima_Detalles INNER JOIN Cat_Materias_Primas ON Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Unidades ON Cat_Materias_Primas.Unidad_ID=Cat_Unidades.Unidad_ID"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima_Detalles.Faltante>0"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Costo_Compra=0"
        If Chk_Producto.Value = 1 Then
            If Cmb_Producto.ListIndex > -1 Then             'Valida si selecciono un producto
                Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
            Else
                Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Nombre LIKE '%" & Cmb_Producto.Text & "%'"
            End If
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Materias_Primas.Nombre,Alm_Entradas_Materia_Prima_Detalles.Numero_Lote"
        Set Rs_Consulta_Detalles_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Detalles_Entradas
            Grid_Productos.Rows = 0
            Grid_Productos.Cols = 11
            Grid_Productos.FixedCols = 1
            Grid_Productos.AddItem "Cambiar" & Chr(9) & "No_Entrada" & Chr(9) & "Materia_Prima_ID" & Chr(9) & "No_Partida" & Chr(9) & "Codigo" & Chr(9) & "Decripcion" & Chr(9) & "No. Lote" & Chr(9) & "Caducidad" & Chr(9) & "Existencia" & Chr(9) & "Costo" & Chr(9) & "Unidad"
            While Not .EOF
                Grid_Productos.AddItem ">>" & Chr(9) & .rdoColumns("No_Entrada") & Chr(9) & .rdoColumns("Materia_Prima_ID") & Chr(9) & .rdoColumns("No_Partida") & Chr(9) & .rdoColumns("Codigo_Producto") & Chr(9) & .rdoColumns("Materia_Prima") & Chr(9) & .rdoColumns("Numero_Lote") & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") & Chr(9) & .rdoColumns("Faltante") & Chr(9) & Format(.rdoColumns("Costo_Compra"), "#0.0000") & Chr(9) & .rdoColumns("Unidad")
                Grid_Productos.FixedRows = 1
                .MoveNext
            Wend
        End With
        Rs_Consulta_Detalles_Entradas.Close
        Grid_Productos.ColWidth(0) = 700            'Cambiar
        Grid_Productos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Productos.ColWidth(1) = 0              'No_Entrada
        Grid_Productos.ColWidth(2) = 0              'Materia_Prima_ID
        Grid_Productos.ColWidth(3) = 0              'No_Partida
        Grid_Productos.ColWidth(4) = 1800           'Codigo
        Grid_Productos.ColAlignment(4) = flexAlignLeftCenter
        Grid_Productos.ColWidth(5) = 4300           'Producto
        Grid_Productos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Productos.ColWidth(6) = 1400           'Numero_Lote
        Grid_Productos.ColAlignment(6) = flexAlignLeftCenter
        Grid_Productos.ColWidth(7) = 1100           'Fecha_Caducidad
        Grid_Productos.ColAlignment(7) = flexAlignCenterCenter
        Grid_Productos.ColWidth(8) = 1100           'Existencia
        Grid_Productos.ColWidth(9) = 1100           'Costo
        Grid_Productos.ColWidth(10) = 1000          'Unidad
    ElseIf Cmb_Tipo_Producto.ListIndex = 1 Then
        'Consulta el inventario de producto terminado
        Mi_SQL = "SELECT Detalles_Entradas.*,Cat_Productos.Nombre AS Producto_Terminado,Cat_Productos.Codigo AS Codigo_Producto,Cat_Almacenes.Nombre AS Almacen,Cat_Racks.Nombre AS Rack"
        Mi_SQL = Mi_SQL & " FROM Detalles_Entradas INNER JOIN Cat_Productos ON Detalles_Entradas.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Almacenes ON Detalles_Entradas.Almacen_ID=Cat_Almacenes.Almacen_ID"
        Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Racks ON Detalles_Entradas.Rack_ID=Cat_Racks.Rack_ID"
        Mi_SQL = Mi_SQL & " WHERE Detalles_Entradas.Faltante>0"
        If Cmb_Producto.ListIndex > -1 Then             'Valida si selecciono un producto
            Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Productos.Nombre,Detalles_Entradas.Numero_Lote"
        Set Rs_Consulta_Detalles_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Detalles_Entradas
            Grid_Productos.Rows = 0
            Grid_Productos.Cols = 15
            Grid_Productos.FixedCols = 1
            Grid_Productos.AddItem "Mover" & Chr(9) & "No_Entrada" & Chr(9) & "Producto_ID" & Chr(9) & "Tipo" & Chr(9) & "Empresa" & Chr(9) & "Codigo" & Chr(9) & "Decripcion" & Chr(9) & "No. Lote" & Chr(9) & "Caducidad" & Chr(9) & "Existencia" & Chr(9) & "Almacen" & Chr(9) & "Rack" & Chr(9) & "Fondo" & Chr(9) & "Nivel" & Chr(9) & ""
            While Not .EOF
                Empresa = ""
                Select Case .rdoColumns("Empresa")
                    Case 1
                        Empresa = "NHE"
                    Case 2
                        Empresa = "PNC"
                    Case 3
                        Empresa = "MAQ"
                End Select
                Grid_Productos.AddItem ">>" & Chr(9) & .rdoColumns("No_Entrada") & Chr(9) & .rdoColumns("Producto_ID") & Chr(9) & "PT" & Chr(9) & Empresa & Chr(9) & .rdoColumns("Codigo_Producto") & Chr(9) & .rdoColumns("Producto_Terminado") & Chr(9) & .rdoColumns("Numero_Lote") & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") & Chr(9) & .rdoColumns("Faltante") & Chr(9) & .rdoColumns("Almacen") & Chr(9) & .rdoColumns("Rack") & Chr(9) & Format(.rdoColumns("Fila"), "00") & Chr(9) & .rdoColumns("Nivel") & Chr(9) & ""
                Grid_Productos.FixedRows = 1
                .MoveNext
            Wend
        End With
        Rs_Consulta_Detalles_Entradas.Close
        Grid_Productos.ColWidth(0) = 600            'Mover
        Grid_Productos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Productos.ColWidth(1) = 0              'No_Entrada
        Grid_Productos.ColWidth(2) = 0              'Producto_ID
        Grid_Productos.ColWidth(3) = 400            'Tipo Producto
        Grid_Productos.ColAlignment(3) = flexAlignCenterCenter
        Grid_Productos.ColWidth(4) = 700            'Empresa
        Grid_Productos.ColAlignment(4) = flexAlignCenterCenter
        Grid_Productos.ColWidth(5) = 1500           'Codigo
        Grid_Productos.ColAlignment(5) = flexAlignLeftCenter
        Grid_Productos.ColWidth(6) = 2950           'Producto
        Grid_Productos.ColAlignment(6) = flexAlignLeftCenter
        Grid_Productos.ColWidth(7) = 1400           'Numero_Lote
        Grid_Productos.ColAlignment(7) = flexAlignLeftCenter
        Grid_Productos.ColWidth(8) = 1100           'Fecha_Caducidad
        Grid_Productos.ColAlignment(8) = flexAlignCenterCenter
        Grid_Productos.ColWidth(9) = 950            'Existencia
        Grid_Productos.ColWidth(10) = 1350          'Almacen
        Grid_Productos.ColAlignment(10) = flexAlignLeftCenter
        Grid_Productos.ColWidth(11) = 500           'Rack
        Grid_Productos.ColAlignment(11) = flexAlignCenterCenter
        Grid_Productos.ColWidth(12) = 550           'Fondo
        Grid_Productos.ColAlignment(12) = flexAlignCenterCenter
        Grid_Productos.ColWidth(13) = 500           'Nivel
        Grid_Productos.ColAlignment(13) = flexAlignCenterCenter
        Grid_Productos.ColWidth(14) = 0             '
    Else
        MsgBox "Seleccione un tipo de inventario", vbExclamation
    End If
End Sub

Private Sub Btn_Cambiar_Click()
Dim Mi_SQL As String
Dim Rs_Actualiza_Entrada As rdoResultset
Dim Rs_Actualiza_Catalogo As rdoResultset
Dim Actualizo_Entradas As Boolean
Dim Materia_Prima As String
Dim Fila As Integer

On Error GoTo Handler
    If MsgBox("¿Está seguro de cambiar el costo del producto?", vbQuestion + vbYesNo, "Actualizacion de costos") = vbYes Then
        If Val(Txt_Costo_Sin_IVA.Text) > 0 Then
            Conexion_Base.BeginTrans
            'Valida que inventario fue el que consulto el usuario
            If Cmb_Tipo_Producto.ListIndex = 0 Then     'Materia Prima
                'Consulta la entrada para hacer el cambio del costo
                Mi_SQL = "SELECT * FROM Alm_Entradas_Materia_Prima_Detalles"
                Mi_SQL = Mi_SQL & " WHERE No_Partida=" & Val(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 3))
            Else                                        'Producto Terminado
                Mi_SQL = "SELECT * FROM Detalles_Entradas"
                Mi_SQL = Mi_SQL & " WHERE No_Entrada='" & Trim(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 1)) & "'"
                Mi_SQL = Mi_SQL & " AND Producto_ID='" & Trim(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 2)) & "'"
            End If
            Set Rs_Actualiza_Entrada = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Actualiza_Entrada.EOF Then
                With Rs_Actualiza_Entrada
                    .Edit
                        .rdoColumns("Costo_Compra") = Val(Txt_Costo_Sin_IVA.Text)
                        .rdoColumns("Moneda") = Cmb_Moneda.Text
                        .rdoColumns("Tipo_Cambio") = Val(Txt_Tipo_Cambio.Text)
                        .rdoColumns("Tipo_Cambio_2") = Val(Conectar_Ayudante.Consulta_Tipo_Cambio(Format(Dtp_Fecha_Factura.Value, "MM/dd/yyyy"), "DOLARES"))
                        .rdoColumns("Costo_Conversion") = Val(Txt_Costo_Pesos.Text)
                    .Update
                End With
            End If
            Rs_Actualiza_Entrada.Close
            'Actualiza el costo de todos los productos
            If MsgBox("Se ha actualizado el costo del producto para esta entrada" & Chr(13) & "¿Desea actualizar todos lo demás que tengan costo 0 del mismo producto?", vbQuestion + vbYesNo) = vbYes Then
                'Valida que inventario fue el que consulto el usuario
                If Cmb_Tipo_Producto.ListIndex = 0 Then     'Materia Prima
                    Mi_SQL = "UPDATE Alm_Entradas_Materia_Prima_Detalles"
                    Mi_SQL = Mi_SQL & " SET Costo_Compra=" & Val(Txt_Costo_Sin_IVA.Text)
                    Mi_SQL = Mi_SQL & " , Moneda='" & Cmb_Moneda.Text & "'"
                    Mi_SQL = Mi_SQL & " , Tipo_Cambio=" & Val(Txt_Tipo_Cambio.Text)
                    Mi_SQL = Mi_SQL & " , Tipo_Cambio_2=" & Val(Conectar_Ayudante.Consulta_Tipo_Cambio(Format(Dtp_Fecha_Factura.Value, "MM/dd/yyyy"), "DOLARES"))
                    Mi_SQL = Mi_SQL & " , Costo_Conversion=" & Val(Txt_Costo_Pesos.Text)
                    Mi_SQL = Mi_SQL & " WHERE Materia_Prima_ID='" & Trim(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 2)) & "'"
                    Mi_SQL = Mi_SQL & " AND Costo_Compra=0"
                    Mi_SQL = Mi_SQL & " AND Faltante>0"
                    Conexion_Base.Execute Mi_SQL
                End If
                Actualizo_Entradas = True
                Materia_Prima = Trim(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 2))
            End If
            'Actualiza el costo del catálogo
            If MsgBox("¿Desea actualizar el último costo en el catálogo?", vbQuestion + vbYesNo) = vbYes Then
                If Cmb_Tipo_Producto.ListIndex = 0 Then     'Materia Prima
                    Mi_SQL = "SELECT Materia_Prima_ID,Costo FROM Cat_Materias_Primas"
                    Mi_SQL = Mi_SQL & " WHERE Materia_Prima_ID='" & Trim(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 2)) & "'"
                    Set Rs_Actualiza_Catalogo = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Actualiza_Catalogo.EOF Then
                        Rs_Actualiza_Catalogo.Edit
                            Rs_Actualiza_Catalogo.rdoColumns("Costo") = Val(Txt_Costo_Pesos.Text)
                        Rs_Actualiza_Catalogo.Update
                    End If
                    Rs_Actualiza_Catalogo.Close
                End If
            End If
            Conexion_Base.CommitTrans
            Btn_Regresar_Click
            If Actualizo_Entradas = True Then
1:              For Fila = 1 To Grid_Productos.Rows - 1
                    If Materia_Prima = Trim(Grid_Productos.TextMatrix(Fila, 2)) Then
                        If Grid_Productos.Rows = 2 Then 'Si solo quedan la fila fija del encabezado y otra mas
                            Grid_Productos.FixedRows = 0
                            Grid_Productos.RemoveItem Grid_Productos.RowSel + 1
                        Else
                            Grid_Productos.RemoveItem Grid_Productos.RowSel
                        End If
                        GoTo 1
                    End If
                Next Fila
            Else
                If Grid_Productos.Rows = 2 Then 'Si solo quedan la fila fija del encabezado y otra mas
                    Grid_Productos.FixedRows = 0
                    Grid_Productos.RemoveItem Grid_Productos.RowSel + 1
                Else
                    Grid_Productos.RemoveItem Grid_Productos.RowSel
                End If
            End If
        Else
            MsgBox "Para hacer el cambio debe capturar los datos completos del costo", vbExclamation
        End If
    End If
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Regresar_Click()
    Fra_Cambio_Ubicacion.Visible = False
    Fra_Busqueda_Entradas.Enabled = True
    Grid_Productos.Enabled = True
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Chk_Producto_Click()
    If Chk_Producto.Value = 1 Then
        Cmb_Producto.Enabled = True
    Else
        Cmb_Producto.Enabled = False
        Cmb_Producto.Text = ""
        Cmb_Producto.Clear
    End If
End Sub

Private Sub Chk_Tipo_Producto_Click()
    If Chk_Tipo_Producto.Value = 1 Then
        Cmb_Tipo_Producto.Enabled = True
    Else
        Cmb_Tipo_Producto.Enabled = False
        Cmb_Tipo_Producto.ListIndex = -1
    End If
End Sub

Private Sub Cmb_Moneda_Click()
    Txt_Tipo_Cambio.Text = Conectar_Ayudante.Consulta_Tipo_Cambio(Format(Dtp_Fecha_Factura.Value, "MM/dd/yyyy"), Cmb_Moneda.Text)
    Txt_Costo_Sin_IVA_Change
End Sub

Private Sub Cmb_Producto_KeyPress(KeyAscii As Integer)
Dim Mi_SQL As String
Dim Rs_Consulta_Producto As rdoResultset

    If KeyAscii = 13 Then
        Select Case Cmb_Tipo_Producto.ListIndex
            Case -1
                MsgBox "Debe seleccionar el tipo de inventario", vbExclamation
                Exit Sub
            Case 0 'Productos de materia prima
                Mi_SQL = "SELECT Materia_Prima_ID,Codigo,Nombre"
                Mi_SQL = Mi_SQL & " FROM Cat_Materias_Primas"
                Mi_SQL = Mi_SQL & " WHERE (Nombre LIKE '%" & Cmb_Producto.Text & "%'"
                Mi_SQL = Mi_SQL & " OR Codigo LIKE '%" & Cmb_Producto.Text & "%')"
                Mi_SQL = Mi_SQL & " ORDER BY Nombre"
            Case 1  'Productos terminados
                Mi_SQL = "SELECT Producto_ID,Codigo,Nombre"
                Mi_SQL = Mi_SQL & " FROM Cat_Productos"
                Mi_SQL = Mi_SQL & " WHERE (Nombre LIKE '%" & Cmb_Producto.Text & "%'"
                Mi_SQL = Mi_SQL & " OR Codigo LIKE '%" & Cmb_Producto.Text & "%')"
                Mi_SQL = Mi_SQL & " ORDER BY Nombre"
        End Select
        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Producto.Clear
        With Rs_Consulta_Producto
            While Not .EOF
                'Agrega todos los valores obtenidos de la consulta anterior en el combo
                Cmb_Producto.AddItem .rdoColumns("Codigo") & " - " & .rdoColumns("Nombre")
                Cmb_Producto.ItemData(Cmb_Producto.NewIndex) = .rdoColumns(0)
                .MoveNext
            Wend
        End With
        Rs_Consulta_Producto.Close
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Dtp_Fecha_Factura_Change()
    Cmb_Moneda_Click
End Sub

Private Sub Dtp_Fecha_Factura_Click()
    Cmb_Moneda_Click
End Sub

Private Sub Form_Load()
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
    Me.Top = 0
    Chk_Tipo_Producto.Value = 1
    Cmb_Tipo_Producto.ListIndex = 0
    Cmb_Tipo_Producto.Locked = True
    Cmb_Moneda.ListIndex = 0
    Dtp_Fecha_Factura.Value = Now
End Sub

Private Sub Grid_Productos_DblClick()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Unidades As rdoResultset

    If Grid_Productos.Rows > -1 Then
        If Grid_Productos.ColSel = 10 Then      'Celda de mover productos
            Fra_Busqueda_Entradas.Enabled = False
            Grid_Productos.Enabled = False
            Fra_Cambio_Ubicacion.Visible = True
            Lbl_Descripcion_Producto.Caption = "PRODUCTO: " & Grid_Productos.TextMatrix(Grid_Productos.RowSel, 4)
            Lbl_Unidad.Caption = Grid_Productos.TextMatrix(Grid_Productos.RowSel, 10)
            Lbl_Ubicacion.Caption = Grid_Productos.TextMatrix(Grid_Productos.RowSel, 5)
            Txt_Costo_Sin_IVA.Text = Val(Grid_Productos.TextMatrix(Grid_Productos.RowSel, 9))
            Txt_Costo_Sin_IVA.SetFocus
            SendKeys "{Home}+{End}"
        Else                                    'Celdas para ordenar
            If Grid_Productos.MouseRow = 0 Then
                'Valida si intenta ordenar por fecha de caducidad
                'If Grid_Productos.ColSel = 8 Then
                '    Grid_Productos.ColSel = 14
                'End If
                'Ordena usando la columna a la que se le dio clic
                Grid_Productos.Row = 0
                Grid_Productos.RowSel = 0
                'Oculta el Grid
                Grid_Productos.Visible = False
                Grid_Productos.Refresh
                If Columna_Orden_Seleccionada <> Grid_Productos.ColSel Then
                    Modo_Orden_Grid_Productos = flexSortGenericAscending
                ElseIf Modo_Orden_Grid_Productos = flexSortGenericAscending Then
                    Modo_Orden_Grid_Productos = flexSortGenericDescending
                Else
                    Modo_Orden_Grid_Productos = flexSortGenericAscending
                End If
                Grid_Productos.Sort = Modo_Orden_Grid_Productos
                'Restaura el nombre de la otra columna
                If Grid_Productos.ColSel >= 0 Then
                    Grid_Productos.TextMatrix(0, Grid_Productos.ColSel) = Mid$(Grid_Productos.TextMatrix(0, Grid_Productos.ColSel), 1)
                End If
                'Muestra el caracter > en el nombre de la columna seleccionada
                Columna_Orden_Seleccionada = Grid_Productos.ColSel
                If Modo_Orden_Grid_Productos = flexSortGenericAscending Then
                    Grid_Productos.TextMatrix(0, Grid_Productos.ColSel) = Grid_Productos.TextMatrix(0, Grid_Productos.ColSel)
                Else
                    Grid_Productos.TextMatrix(0, Grid_Productos.ColSel) = Grid_Productos.TextMatrix(0, Grid_Productos.ColSel)
                End If
                'Muestra el Grid_Productos
                Grid_Productos.Visible = True
                Grid_Productos.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Txt_Costo_Sin_IVA_Change()
Dim Costo As Double
    'Pesos
    If Cmb_Moneda.Text = "PESOS" Or Cmb_Moneda.Text = "DOLARES" Then
        Txt_Costo_Pesos.Text = Format(Val(Txt_Costo_Sin_IVA.Text) * Val(Txt_Tipo_Cambio.Text), "#0.0000")
    Else
        Txt_Costo_Pesos.Text = Format((Val(Txt_Costo_Sin_IVA.Text) * Val(Txt_Tipo_Cambio.Text)) * Conectar_Ayudante.Consulta_Tipo_Cambio(Format(Dtp_Fecha_Factura.Value, "MM/dd/yyyy"), "DOLARES"), "#0.0000")
    End If
End Sub

Private Sub Txt_Costo_Sin_IVA_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Costo_Sin_IVA.Text, True)
End Sub

