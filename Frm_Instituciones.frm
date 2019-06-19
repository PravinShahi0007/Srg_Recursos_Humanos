VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Instituciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7653.125
   ScaleMode       =   0  'User
   ScaleWidth      =   7545
   Begin VB.PictureBox Pic_Cat_Empleados 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1680
         Picture         =   "Frm_Instituciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "M"
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3120
         Picture         =   "Frm_Instituciones.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "B"
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   6000
         Picture         =   "Frm_Instituciones.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   240
         Picture         =   "Frm_Instituciones.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "A"
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4560
         Picture         =   "Frm_Instituciones.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "C"
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Generales_Cat_Instituciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Instituciones"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   7200
         Begin VB.ComboBox Cmb_Estatus 
            Height          =   315
            ItemData        =   "Frm_Instituciones.frx":050A
            Left            =   4560
            List            =   "Frm_Instituciones.frx":0514
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   2370
         End
         Begin VB.TextBox Txt_Descripcion 
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1680
            Width           =   5800
         End
         Begin VB.TextBox Txt_Estado 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1320
            Width           =   2370
         End
         Begin VB.TextBox Txt_Ciudad 
            Height          =   315
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   5
            Top             =   960
            Width           =   2370
         End
         Begin VB.TextBox Txt_Direccion 
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   4
            Top             =   960
            Width           =   2370
         End
         Begin VB.TextBox Txt_Institucion_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   1
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Clave 
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   2
            Top             =   270
            Width           =   2370
         End
         Begin VB.TextBox Txt_Nombre 
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   3
            Top             =   600
            Width           =   5800
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Estatus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3600
            TabIndex        =   25
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   24
            Top             =   1770
            Width           =   930
         End
         Begin VB.Label Lbl_Ciudad 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Ciudad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3600
            TabIndex        =   23
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Lbl_Direccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Dirección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   22
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label Lbl_Estado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Estado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   1410
            Width           =   675
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Institución ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   20
            Top             =   330
            Width           =   930
         End
         Begin VB.Label Lbl_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Clave"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   19
            Top             =   330
            Width           =   570
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   18
            Top             =   690
            Width           =   735
         End
      End
      Begin VB.Frame Fra_Instituciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Instituciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Instituciones 
            Height          =   3120
            Left            =   75
            TabIndex        =   9
            Top             =   225
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5503
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "INSTITUCIONES"
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
         Left            =   2670
         TabIndex        =   15
         Top             =   15
         Width           =   3045
      End
   End
End
Attribute VB_Name = "Frm_Cat_Instituciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicializa()
Consulta_Cat_Instituciones ""
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Instituciones
    'DESCRIPCIÓN:           Consulta las Instituciones y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la Institución
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           21 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Instituciones(Nombre As String)
Dim Rs_Consulta_Cat_Instituciones As rdoResultset       'Informacion de los registros

    Grid_Cat_Instituciones.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Instituciones"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " OR Clave LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
'    MsgBox Mi_SQL

    Set Rs_Consulta_Cat_Instituciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Cat_Instituciones
        If Not .EOF Then

            Grid_Cat_Instituciones.AddItem "Institucion ID" & Chr(9) & "Clave" & Chr(9) & "Nombre" & Chr(9) & "Estatus"
            While Not .EOF
                Grid_Cat_Instituciones.AddItem .rdoColumns("Institucion_Id") & Chr(9) & .rdoColumns("Clave") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Estatus")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Instituciones.FixedRows = 1
            Grid_Cat_Instituciones.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Instituciones.ColWidth(1) = 1000   'Clave
            Grid_Cat_Instituciones.ColWidth(2) = 3700   'Nombre
            Grid_Cat_Instituciones.ColWidth(3) = 1500   'Estatus
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Instituciones = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o Clave para buscar las instituciones")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Cat_Instituciones Nombre
       
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo Fin
    If Txt_Institucion_Id.Text <> "" Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Mi_SQL = "DELETE FROM Cat_Instituciones WHERE Institucion_ID='" & Trim(Txt_Institucion_Id.Text) & "'"
            Conexion_Base.Execute Mi_SQL
            'Quita los datos del usuario contenidos en el Grid
            If Grid_Cat_Instituciones.Rows = 2 Then
                Grid_Cat_Instituciones.Rows = 0
            Else
                Grid_Cat_Instituciones.RemoveItem Grid_Cat_Instituciones.RowSel
            End If 'Grid_productos
            MsgBox "Institución Eliminada", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        MsgBox ("Es necesario seleccionar un registro para eliminar")
    End If
Exit Sub
Fin:
    If Err.Number <> 0 Then
        Conexion_Base.RollbackTrans
        If Err.Number = 40002 Then
            MsgBox "NO SE ELIMINO EL REGISTRO. El registro tiene dependencia en otro catalogo, elimine la relacion y vuelva a intentarlo.", vbExclamation, "Mensaje"
        Else
            MsgBox Err.Description, vbExclamation
        End If
    End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
If Txt_Institucion_Id.Text <> "" Then
Call Configurar_Formulario(True)
    Btn_Modificar.Enabled = True
    Btn_Modificar.Caption = "Guardar"
    Txt_Clave.SetFocus
Else
MsgBox ("Es necesario seleccionar un registro para modificar")
End If
Else
Modificar_Cat_Institucion
Limpiar_Formulario
        Btn_Modificar.Caption = "Modificar"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"
End If
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
    Call Configurar_Formulario(True)
    Limpiar_Formulario
    Btn_Nuevo.Enabled = True
    Btn_Nuevo.Caption = "Guardar"
    Txt_Clave.SetFocus
    Fra_Instituciones.Enabled = False
Else
    If Validar_Componentes Then
        Call Alta_Institucion
        Limpiar_Formulario
        Btn_Nuevo.Caption = "Nuevo"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"
        Fra_Instituciones.Enabled = True
    Else
        MsgBox ("Todos los campos marcados con * son necesarios")
    End If
End If



End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
Fra_Generales_Cat_Instituciones.Enabled = Habilitar
Btn_Nuevo.Enabled = Not Habilitar
Btn_Modificar.Enabled = Not Habilitar
Btn_Eliminar.Enabled = Not Habilitar
Btn_Buscar.Enabled = Not Habilitar
Btn_Salir.Caption = "Cancelar"

End Sub
Function Validar_Componentes() As Boolean
Validar_Componentes = True
If Txt_Clave.Text = "" Then
Validar_Componentes = False
End If
If Txt_Nombre.Text = "" Then
Validar_Componentes = False
End If
If Txt_Direccion.Text = "" Then
Validar_Componentes = False
End If
If Txt_Ciudad.Text = "" Then
Validar_Componentes = False
End If
If Txt_Estado.Text = "" Then
Validar_Componentes = False
End If
If Cmb_Estatus.ListIndex = -1 Then
Validar_Componentes = False
End If

End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Institucion
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos de la institución que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 21-Diciembre-2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Institucion()
'Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Instituciones As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans
    
    'Alta de Institución
    Set Rs_Alta_Cat_Instituciones = Conectar_Ayudante.Recordset_Agregar("Cat_Instituciones")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Instituciones
        .AddNew
            Txt_Institucion_Id.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Instituciones", "Institucion_Id"), "00000")
            .rdoColumns("Institucion_Id") = Txt_Institucion_Id.Text
            .rdoColumns("Clave") = Trim(Txt_Clave.Text)
            .rdoColumns("Nombre") = UCase(Txt_Nombre.Text)
            .rdoColumns("Direccion") = UCase(Txt_Direccion.Text)
            .rdoColumns("Ciudad") = UCase(Txt_Ciudad.Text)
            .rdoColumns("Estado") = UCase(Txt_Estado.Text)
            .rdoColumns("Descripcion") = UCase(Txt_Descripcion.Text)
            .rdoColumns("Estatus") = Cmb_Estatus.Text
'            .rdoColumns("No_Nomina") = Val(Txt_No_Nomina.Text)
'            .rdoColumns("Area_ID") = Format(Cmb_Area_ID.ItemData(Cmb_Area_ID.ListIndex), "00000")
'            .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
'            .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        'Guarda el password en la tabla
'        Mi_SQL = "INSERT INTO Cat_Instituciones ()"
'        Mi_SQL = Mi_SQL & " VALUES('" & Trim(Txt_Institucion_Id.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Clave.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Nombre.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Direccion.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Ciudad.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Estado.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Descripcion.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Trim(Cmb_Estatus.Text) & "'"
'        Mi_SQL = Mi_SQL & " , '" & Nombre_Usuario & "'"
'        Mi_SQL = Mi_SQL & " , '" & Format(Now, "yyyymmdd") & "'"
'        Mi_SQL = Mi_SQL & " , NULL, NULL)"
'        Mi_SQL = Mi_SQL & " , '" & Format(Now, "MM/dd/yyyy") & "')"
'        Conexion_Base.Execute Mi_SQL

    End With
    Rs_Alta_Cat_Instituciones.Close
    Conexion_Base.CommitTrans
    MsgBox "Institución agregada", vbInformation
    Consulta_Cat_Instituciones ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Institucion_Id.Text = ""
Txt_Clave.Text = ""
Txt_Nombre.Text = ""
Txt_Direccion.Text = ""
Txt_Ciudad.Text = ""
Txt_Estado.Text = ""
Txt_Descripcion.Text = ""
Cmb_Estatus.ListIndex = -1
End Sub

Private Sub Btn_Salir_Click()
If Btn_Salir.Caption = "Salir" Then
    Unload Me
Else
    Limpiar_Formulario
    Configurar_Formulario (False)
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Fra_Instituciones.Enabled = True
End If
    
End Sub

Private Sub Grid_Cat_Instituciones_Click()
Dim Rs_Consulta_Cat_Tipos_Notas_Credito As rdoResultset
    If Grid_Cat_Instituciones.Rows > 1 Then
        Txt_Institucion_Id.Text = Grid_Cat_Instituciones.TextMatrix(Grid_Cat_Instituciones.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Instituciones"
        Mi_SQL = Mi_SQL & "  WHERE Institucion_Id='" & Txt_Institucion_Id.Text & "'"
        Set Rs_Consulta_Cat_Instituciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Instituciones.EOF Then
            With Rs_Consulta_Cat_Instituciones
                Txt_Institucion_Id.Text = .rdoColumns("Institucion_Id")
                Txt_Clave.Text = .rdoColumns("Clave")
                Txt_Nombre.Text = .rdoColumns("Nombre")
                Txt_Direccion.Text = .rdoColumns("Direccion")
                Txt_Ciudad.Text = .rdoColumns("Ciudad")
                Txt_Estado.Text = .rdoColumns("Estado")
                If Not IsNull(.rdoColumns("Descripcion")) Then Txt_Descripcion.Text = .rdoColumns("Descripcion")
                If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Estatus")), Cmb_Estatus)
                Else
                    Cmb_Estatus.ListIndex = -1
                End If
            End With
        End If
        Rs_Consulta_Cat_Instituciones.Close
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Institucion
    'DESCRIPCIÓN:           Modifica el registro de la Institución
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    21 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Cat_Institucion()
Dim Rs_Modificacion_Cat_Institucion As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Instituciones"
    Mi_SQL = Mi_SQL & " WHERE Institucion_Id ='" & Trim(Txt_Institucion_Id.Text) & "'"
    Set Rs_Modificacion_Cat_Institucion = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Institucion
        .Edit
            .rdoColumns("Clave") = Trim(Txt_Clave.Text)
            .rdoColumns("Nombre") = Trim(Txt_Nombre.Text)
            .rdoColumns("Direccion") = Trim(Txt_Direccion.Text)
            .rdoColumns("Ciudad") = Trim(Txt_Ciudad.Text)
            .rdoColumns("Estado") = Trim(Txt_Estado.Text)
            .rdoColumns("Descripcion") = Trim(Txt_Descripcion.Text)
            .rdoColumns("Estatus") = Trim(Cmb_Estatus.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Institucion = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "La Institución ha sido modificada", vbInformation + vbOKOnly, Me.Caption
   Consulta_Cat_Instituciones ""
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Instituciones
    'DESCRIPCIÓN:           Consulta las Instituciones y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la institución
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           21 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
'Private Sub Consulta_Cat_Instituciones(Nombre As String)
'Dim Rs_Consulta_Cat_Instituciones As rdoResultset       'Informacion de los registros
'
'    Grid_Cat_Instituciones.Rows = 0
'
'    'Consulta los datos generales del usuario
'    Mi_SQL = "SELECT *"
'    Mi_SQL = Mi_SQL & " FROM Cat_Empresas"
'    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
'    Mi_SQL = Mi_SQL & " OR Clave LIKE '%" & Nombre & "%'"
'    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
'    Set Rs_Consulta_Cat_Instituciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'
'    With Rs_Consulta_Cat_Instituciones
'        If Not .EOF Then
'
'            Grid_Cat_Instituciones.AddItem "Institucion ID" & Chr(9) & "Nombre" & Chr(9) & "Clave"
'            While Not .EOF
'                Grid_Cat_Empresas.AddItem .rdoColumns("Instituciin_Id") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Clave")
'                .MoveNext
'            Wend
'            'Configura el tamaño de las columnas del grid_usuarios
'            Grid_Cat_Instituciones.FixedRows = 1
'            Grid_Cat_Instituciones.ColWidth(0) = 0      'Empresa_ID
'            Grid_Cat_Empresas.ColWidth(1) = 6000   'Nombre
'            Grid_Cat_Empresas.ColWidth(2) = 1800   'Acronimo
'            .Close
'        End If
'    End With
'    'Cierra el manejador del registro
'    Set Rs_Consulta_Cat_Instituciones = Nothing
'End Sub
