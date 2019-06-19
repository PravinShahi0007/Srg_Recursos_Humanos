VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Instructores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7273.768
   ScaleMode       =   0  'User
   ScaleWidth      =   12235.81
   Begin VB.PictureBox Pic_Cat_Empleados 
      BackColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   8340
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Instructores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Instructores"
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
         TabIndex        =   15
         Top             =   2280
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Instructores 
            Height          =   3120
            Left            =   75
            TabIndex        =   16
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
      Begin VB.Frame Fra_Generales_Cat_Instructores 
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
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7200
         Begin VB.ComboBox Cmb_Cat_Instructores_Intitucion_Id 
            Height          =   315
            ItemData        =   "Frm_Cat_Instructores.frx":0000
            Left            =   1200
            List            =   "Frm_Cat_Instructores.frx":000A
            TabIndex        =   21
            Top             =   1320
            Width           =   5850
         End
         Begin VB.TextBox Txt_Cat_Instructores_A_Materno 
            Height          =   315
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   19
            Top             =   960
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Instructores_A_Paterno 
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   10
            Top             =   960
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Instructores_Nombre 
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   9
            Top             =   600
            Width           =   5850
         End
         Begin VB.TextBox Txt_Cat_Instructores_Instructor_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Cat_Instructores_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Instructores.frx":0020
            Left            =   4680
            List            =   "Frm_Cat_Instructores.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   2370
         End
         Begin VB.Label Lbl_Direccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Institución"
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
            Left            =   120
            TabIndex        =   20
            Top             =   1410
            Width           =   960
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "A. Materno"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3720
            TabIndex        =   18
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*A. Paterno"
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
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   990
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
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Instructor ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   870
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
            Left            =   3720
            TabIndex        =   11
            Top             =   330
            Width           =   720
         End
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4560
         Picture         =   "Frm_Cat_Instructores.frx":0040
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "C"
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   240
         Picture         =   "Frm_Cat_Instructores.frx":0142
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "A"
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   6000
         Picture         =   "Frm_Cat_Instructores.frx":0244
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3120
         Picture         =   "Frm_Cat_Instructores.frx":0346
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "B"
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1680
         Picture         =   "Frm_Cat_Instructores.frx":0448
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "M"
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "INSTRUCTORES"
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
         Left            =   2655
         TabIndex        =   17
         Top             =   15
         Width           =   3075
      End
   End
End
Attribute VB_Name = "Frm_Cat_Instructores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
Consulta_Cat_Instructores ""
Call Conectar_Ayudante.Llena_Combo_Item("Institucion_Id, Nombre", "Cat_Instituciones WHERE Estatus='ACTIVO'", Cmb_Cat_Instructores_Intitucion_Id, 0, "Institucion_Id", "", False, "")
 
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Instructores
    'DESCRIPCIÓN:           Consulta las Instituciones y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la Institución
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           28 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Instructores(Nombre As String)
Dim Rs_Consulta_Cat_Instructores As rdoResultset       'Informacion de los registros

    Grid_Cat_Instructores.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Instructores"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " OR Apellido_Paterno LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
'    MsgBox Mi_SQL

    Set Rs_Consulta_Cat_Instructores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Cat_Instructores
        If Not .EOF Then

            Grid_Cat_Instructores.AddItem "Instructor ID" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno"
            While Not .EOF
                Grid_Cat_Instructores.AddItem .rdoColumns("Instructor_Id") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Instructores.FixedRows = 1
            Grid_Cat_Instructores.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Instructores.ColWidth(1) = 3000   'Nombre
            Grid_Cat_Instructores.ColWidth(2) = 1800   'apellido paterno
            Grid_Cat_Instructores.ColWidth(3) = 1800   'apellido materno
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Instructores = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o Apellido Paternoo para buscar los instructores")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Cat_Instructores Nombre
       
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo Fin
    If Txt_Cat_Instructores_Instructor_Id.Text <> "" Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Mi_SQL = "DELETE FROM Cat_Instructores WHERE Instructor_ID='" & Trim(Txt_Cat_Instructores_Instructor_Id.Text) & "'"
            Conexion_Base.Execute Mi_SQL
            'Quita los datos del usuario contenidos en el Grid
            If Grid_Cat_Instructores.Rows = 2 Then
                Grid_Cat_Instructores.Rows = 0
            Else
                Grid_Cat_Instructores.RemoveItem Grid_Cat_Instructores.RowSel
            End If 'Grid_productos
            MsgBox "Instructor Eliminado", vbInformation + vbOKOnly, Me.Caption
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
    If Txt_Cat_Instructores_Instructor_Id.Text <> "" Then
        Call Configurar_Formulario(True)
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Guardar"
        Fra_Instructores.Enabled = False
    Else
        MsgBox ("Es necesario seleccionar un registro para modificar")
    End If
Else
    Modificar_Cat_Instructores
    Limpiar_Formulario
    Btn_Modificar.Caption = "Modificar"
    Configurar_Formulario (False)
    Btn_Salir.Caption = "Salir"
    Fra_Instructores.Enabled = True
End If
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
    Call Configurar_Formulario(True)
    Limpiar_Formulario
    Btn_Nuevo.Enabled = True
    Btn_Nuevo.Caption = "Guardar"
    Cmb_Cat_Instructores_Estatus.SetFocus
    Fra_Instructores.Enabled = False
Else
    If Validar_Componentes Then
        Call Alta_Instructor
        Limpiar_Formulario
        Btn_Nuevo.Caption = "Nuevo"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"
        Fra_Instructores.Enabled = True
    Else
        MsgBox ("Todos los campos marcados con * son necesarios")
    End If
End If

End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
    Fra_Generales_Cat_Instructores.Enabled = Habilitar
    Btn_Nuevo.Enabled = Not Habilitar
    Btn_Modificar.Enabled = Not Habilitar
    Btn_Eliminar.Enabled = Not Habilitar
    Btn_Buscar.Enabled = Not Habilitar
    Btn_Salir.Caption = "Cancelar"
End Sub
Function Validar_Componentes() As Boolean
    Validar_Componentes = True
    If Txt_Cat_Instructores_Nombre.Text = "" Then
        Validar_Componentes = False
    End If
    If Txt_Cat_Instructores_A_Paterno.Text = "" Then
        Validar_Componentes = False
    End If
    If Cmb_Cat_Instructores_Intitucion_Id.ListIndex = -1 Then
        Validar_Componentes = False
    End If
    If Cmb_Cat_Instructores_Estatus.ListIndex = -1 Then
        Validar_Componentes = False
    End If
End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Instructor
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos de la institución que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 28-Diciembre-2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Instructor()
'Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Instructores As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans
    
    'Alta de Institución
    Set Rs_Alta_Cat_Instructores = Conectar_Ayudante.Recordset_Agregar("Cat_Instructores")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Instructores
        .AddNew
            Txt_Cat_Instructores_Instructor_Id.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Instructores", "Instructor_Id"), "00000")
            .rdoColumns("Instructor_Id") = Txt_Cat_Instructores_Instructor_Id.Text
            .rdoColumns("Nombre") = UCase(Txt_Cat_Instructores_Nombre.Text)
            .rdoColumns("Apellido_Paterno") = UCase(Txt_Cat_Instructores_A_Paterno.Text)
            .rdoColumns("Apellido_Materno") = UCase(Txt_Cat_Instructores_A_Materno.Text)
            .rdoColumns("Institucion_Id") = Format(Cmb_Cat_Instructores_Intitucion_Id.ItemData(Cmb_Cat_Instructores_Intitucion_Id.ListIndex), "00000")
            .rdoColumns("Estatus") = Cmb_Cat_Instructores_Estatus.Text
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Instructores.Close
    Conexion_Base.CommitTrans
    MsgBox "Instructor agregado", vbInformation
    Consulta_Cat_Instructores ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Cat_Instructores_Instructor_Id.Text = ""
Txt_Cat_Instructores_Nombre.Text = ""
Txt_Cat_Instructores_A_Paterno.Text = ""
Txt_Cat_Instructores_A_Materno.Text = ""
Cmb_Cat_Instructores_Intitucion_Id.ListIndex = -1
Cmb_Cat_Instructores_Estatus.ListIndex = -1
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
    Fra_Instructores.Enabled = True
End If
    
End Sub

Private Sub Grid_Cat_Instructores_Click()
Dim Rs_Consulta_Cat_Instructores As rdoResultset
    If Grid_Cat_Instructores.Rows > 1 Then
        Txt_Cat_Instructores_Instructor_Id.Text = Grid_Cat_Instructores.TextMatrix(Grid_Cat_Instructores.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Instructores"
        Mi_SQL = Mi_SQL & "  WHERE Instructor_Id='" & Txt_Cat_Instructores_Instructor_Id.Text & "'"
        Set Rs_Consulta_Cat_Instructores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Instructores.EOF Then
            With Rs_Consulta_Cat_Instructores
                Txt_Cat_Instructores_Instructor_Id.Text = .rdoColumns("Instructor_Id")
                Txt_Cat_Instructores_Nombre = .rdoColumns("Nombre")
                Txt_Cat_Instructores_A_Paterno.Text = .rdoColumns("Apellido_Paterno")
                Txt_Cat_Instructores_A_Materno.Text = .rdoColumns("Apellido_Materno")
                 If Not IsNull(.rdoColumns("Institucion_Id")) Then
'
                    Cmb_Cat_Instructores_Intitucion_Id.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Institucion_Id"), "Cat_Instituciones", "Nombre", "Institucion_Id"), Cmb_Cat_Instructores_Intitucion_Id)
                Else
                    Cmb_Cat_Instructores_Intitucion_Id.ListIndex = -1
                End If
                If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Cat_Instructores_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Estatus")), Cmb_Cat_Instructores_Estatus)
                Else
                    Cmb_Cat_Instructores_Estatus.ListIndex = -1
                End If
            End With
        End If
        Rs_Consulta_Cat_Instructores.Close
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Instructores
    'DESCRIPCIÓN:           Modifica el registro del instructor
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    28 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Cat_Instructores()
Dim Rs_Modificacion_Cat_Instructor As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Instructores"
    Mi_SQL = Mi_SQL & " WHERE Instructor_Id ='" & Trim(Txt_Cat_Instructores_Instructor_Id.Text) & "'"
    Set Rs_Modificacion_Cat_Instructor = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Instructor
        .Edit
            .rdoColumns("Nombre") = Trim(Txt_Cat_Instructores_Nombre.Text)
            .rdoColumns("Apellido_Paterno") = Trim(Txt_Cat_Instructores_A_Paterno.Text)
            .rdoColumns("Apellido_Materno") = Trim(Txt_Cat_Instructores_A_Materno.Text)
            .rdoColumns("Institucion_Id") = Format(Cmb_Cat_Instructores_Intitucion_Id.ItemData(Cmb_Cat_Instructores_Intitucion_Id.ListIndex), "00000")
            .rdoColumns("Estatus") = Trim(Cmb_Cat_Instructores_Estatus.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Instructor = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "El instructor ha sido modificado", vbInformation + vbOKOnly, Me.Caption
   Consulta_Cat_Instructores ""
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub



