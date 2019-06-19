VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Salas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6958.854
   ScaleMode       =   0  'User
   ScaleWidth      =   6938.088
   Begin VB.PictureBox Pic_Cat_Empleados 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7432
      Left            =   0
      ScaleHeight     =   7425
      ScaleWidth      =   7605
      TabIndex        =   0
      Top             =   0
      Width           =   7612
      Begin VB.Frame Fra_Salas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salas"
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
         TabIndex        =   18
         Top             =   2000
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Salas 
            Height          =   3120
            Left            =   75
            TabIndex        =   6
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
      Begin VB.Frame Fra_Generales_Cat_Salas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Salas"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7200
         Begin VB.TextBox Txt_Cat_Salas_Nombre 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Salas_Clave 
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   2
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Salas_Sala_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Salas_Descripcion 
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   5
            Top             =   960
            Width           =   5800
         End
         Begin VB.ComboBox Cmb_Cat_Salas_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Salas.frx":0000
            Left            =   4560
            List            =   "Frm_Cat_Salas.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2370
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
            TabIndex        =   17
            Top             =   690
            Width           =   735
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
            TabIndex        =   16
            Top             =   330
            Width           =   570
         End
         Begin VB.Label Lbl_Cat_Salas_Sala_Id 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Sala ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   1050
            Width           =   930
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
            TabIndex        =   13
            Top             =   690
            Width           =   720
         End
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Picture         =   "Frm_Cat_Salas.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "C"
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   120
         Picture         =   "Frm_Cat_Salas.frx":0122
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "A"
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   5880
         Picture         =   "Frm_Cat_Salas.frx":0224
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3000
         Picture         =   "Frm_Cat_Salas.frx":0326
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "B"
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1560
         Picture         =   "Frm_Cat_Salas.frx":0428
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "M"
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SALAS"
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
         Left            =   3570
         TabIndex        =   19
         Top             =   15
         Width           =   1245
      End
   End
End
Attribute VB_Name = "Frm_Cat_Salas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
Consulta_Cat_Salas ""
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Salas
    'DESCRIPCIÓN:           Consulta las Salas y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la Sala
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           21 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Salas(Nombre As String)
Dim Rs_Consulta_Cat_Salas As rdoResultset       'Informacion de los registros

    Grid_Cat_Salas.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Salas"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " OR Clave LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
'    MsgBox Mi_SQL

    Set Rs_Consulta_Cat_Salas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Cat_Salas
        If Not .EOF Then

            Grid_Cat_Salas.AddItem "Sala ID" & Chr(9) & "Clave" & Chr(9) & "Nombre" & Chr(9) & "Estatus"
            While Not .EOF
                Grid_Cat_Salas.AddItem .rdoColumns("Sala_Id") & Chr(9) & .rdoColumns("Clave") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Estatus")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Salas.FixedRows = 1
            Grid_Cat_Salas.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Salas.ColWidth(1) = 1000   'Nombre
            Grid_Cat_Salas.ColWidth(2) = 3000   'Clave
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Salas = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o Clave para buscar las Salas")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Cat_Salas Nombre
       
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo Fin
    If Txt_Cat_Salas_Sala_Id.Text <> "" Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Mi_SQL = "DELETE FROM Cat_Salas WHERE Sala_ID='" & Trim(Txt_Cat_Salas_Sala_Id.Text) & "'"
            Conexion_Base.Execute Mi_SQL
            'Quita los datos de la sala contenidos en el Grid
            If Grid_Cat_Salas.Rows = 2 Then
                Grid_Cat_Salas.Rows = 0
            Else
                Grid_Cat_Salas.RemoveItem Grid_Cat_Salas.RowSel
            End If 'Grid_productos
            MsgBox "Sala Eliminada", vbInformation + vbOKOnly, Me.Caption
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
    If Txt_Cat_Salas_Sala_Id.Text <> "" Then
        Call Configurar_Formulario(True)
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Guardar"
        Txt_Cat_Salas_Clave.SetFocus
        Fra_Salas.Enabled = False
    Else
        MsgBox ("Es necesario seleccionar un registro para modificar")
    End If
Else
    Modificar_Cat_Salas
    Limpiar_Formulario
    Btn_Modificar.Caption = "Modificar"
    Configurar_Formulario (False)
    Btn_Salir.Caption = "Salir"
    Fra_Salas.Enabled = True
End If
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
    Call Configurar_Formulario(True)
    Limpiar_Formulario
    Btn_Nuevo.Enabled = True
    Btn_Nuevo.Caption = "Guardar"
    Fra_Salas.Enabled = False
    Txt_Cat_Salas_Clave.SetFocus
Else
    If Validar_Componentes Then

        Call Alta_Salas
        Limpiar_Formulario
        Btn_Nuevo.Caption = "Nuevo"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"
        Fra_Salas.Enabled = True
    Else
        MsgBox ("Todos los campos marcados con * son necesarios")
    End If
End If



End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
Fra_Generales_Cat_Salas.Enabled = Habilitar
Btn_Nuevo.Enabled = Not Habilitar
Btn_Modificar.Enabled = Not Habilitar
Btn_Eliminar.Enabled = Not Habilitar
Btn_Buscar.Enabled = Not Habilitar
Btn_Salir.Caption = "Cancelar"

End Sub
Function Validar_Componentes() As Boolean
Validar_Componentes = True
If Txt_Cat_Salas_Clave.Text = "" Then
Validar_Componentes = False
End If
If Txt_Cat_Salas_Nombre.Text = "" Then
Validar_Componentes = False
End If
If Cmb_Cat_Salas_Estatus.ListIndex = -1 Then
Validar_Componentes = False
End If

End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Salas
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos de la sala que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 21-Diciembre-2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Salas()
'Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Salas As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans
    
    'Alta de Institución
    Set Rs_Alta_Cat_Salas = Conectar_Ayudante.Recordset_Agregar("Cat_Salas")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Salas
        .AddNew
            Txt_Cat_Salas_Sala_Id.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Salas", "Sala_Id"), "00000")
            .rdoColumns("Sala_Id") = Txt_Cat_Salas_Sala_Id.Text
            .rdoColumns("Clave") = Trim(Txt_Cat_Salas_Clave)
            .rdoColumns("Nombre") = UCase(Txt_Cat_Salas_Nombre.Text)
            .rdoColumns("Descripcion") = UCase(Txt_Cat_Salas_Descripcion.Text)
            .rdoColumns("Estatus") = Cmb_Cat_Salas_Estatus.Text
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Salas.Close
    Conexion_Base.CommitTrans
    MsgBox "Sala agregada", vbInformation
    Consulta_Cat_Salas ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Cat_Salas_Sala_Id.Text = ""
Txt_Cat_Salas_Clave.Text = ""
Txt_Cat_Salas_Nombre.Text = ""
Txt_Cat_Salas_Descripcion.Text = ""
Cmb_Cat_Salas_Estatus.ListIndex = -1
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
    Fra_Salas.Enabled = True
End If
    
End Sub

Private Sub Grid_Cat_Salas_Click()
Dim Rs_Consulta_Cat_Salas As rdoResultset
    If Grid_Cat_Salas.Rows > 1 Then
        Txt_Cat_Salas_Sala_Id.Text = Grid_Cat_Salas.TextMatrix(Grid_Cat_Salas.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Salas"
        Mi_SQL = Mi_SQL & "  WHERE Sala_Id='" & Txt_Cat_Salas_Sala_Id.Text & "'"
        Set Rs_Consulta_Cat_Salas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Salas.EOF Then
            With Rs_Consulta_Cat_Salas
                Txt_Cat_Salas_Sala_Id.Text = .rdoColumns("Sala_Id")
                Txt_Cat_Salas_Clave.Text = .rdoColumns("Clave")
                Txt_Cat_Salas_Nombre.Text = .rdoColumns("Nombre")
                Txt_Cat_Salas_Descripcion.Text = .rdoColumns("Descripcion")
                If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Cat_Salas_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Estatus")), Cmb_Cat_Salas_Estatus)
                Else
                    Cmb_Cat_Salas_Estatus.ListIndex = -1
                End If
            End With
        End If
        Rs_Consulta_Cat_Salas.Close
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Salas
    'DESCRIPCIÓN:           Modifica el registro de la Sala
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    22 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Cat_Salas()
Dim Rs_Modificacion_Cat_Salas As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Salas"
    Mi_SQL = Mi_SQL & " WHERE Sala_Id ='" & Trim(Txt_Cat_Salas_Sala_Id.Text) & "'"
    Set Rs_Modificacion_Cat_Salas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Salas
        .Edit
            .rdoColumns("Clave") = Trim(Txt_Cat_Salas_Clave.Text)
            .rdoColumns("Nombre") = Trim(Txt_Cat_Salas_Nombre.Text)
            .rdoColumns("Descripcion") = Trim(Txt_Cat_Salas_Descripcion.Text)
            .rdoColumns("Estatus") = Trim(Cmb_Cat_Salas_Estatus.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Salas = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "La Sala ha sido modificada", vbInformation + vbOKOnly, Me.Caption
   Consulta_Cat_Salas ""
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


