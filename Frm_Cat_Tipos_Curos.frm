VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Cat_Tipos_Curos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   4108.583
   ScaleMode       =   0  'User
   ScaleWidth      =   2585.08
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
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1560
         Picture         =   "Frm_Cat_Tipos_Curos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "M"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3000
         Picture         =   "Frm_Cat_Tipos_Curos.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "B"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   5880
         Picture         =   "Frm_Cat_Tipos_Curos.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   120
         Picture         =   "Frm_Cat_Tipos_Curos.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "A"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Picture         =   "Frm_Cat_Tipos_Curos.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "C"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Generales_Cat_Tipos_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Tipos Cursos"
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
         TabIndex        =   12
         Top             =   360
         Width           =   7200
         Begin VB.TextBox Txt_Cat_Tipos_Cursos_Descripcion 
            Height          =   795
            Left            =   1125
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   960
            Width           =   5800
         End
         Begin VB.TextBox Txt_Cat_Tipos_Cursos_Tipo_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Tipos_Cursos_Clave 
            Height          =   315
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   2
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Tipos_Cursos_Nombre 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   5800
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   16
            Top             =   1050
            Width           =   930
         End
         Begin VB.Label Lbl_Cat_Tipos_Cursos_Tipo_Id 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   330
            Width           =   525
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   690
            Width           =   735
         End
      End
      Begin VB.Frame Fra_Tipo_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipos Cursos"
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
         TabIndex        =   11
         Top             =   2235
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Tipos_Cursos 
            Height          =   3120
            Left            =   75
            TabIndex        =   5
            Top             =   225
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5503
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Tipos_Cursos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIPOS CURSOS"
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
         Left            =   2730
         TabIndex        =   17
         Top             =   15
         Width           =   2925
      End
   End
End
Attribute VB_Name = "Frm_Cat_Tipos_Curos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
Consulta_Cat_Tipos_Cursos ""
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Tipos_Cursos
    'DESCRIPCIÓN:           Consulta laos tipos de cursos y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre del tipo de curso
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           22 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Tipos_Cursos(Nombre As String)
Dim Rs_Consulta_Cat_Tipos_Cursos As rdoResultset       'Informacion de los registros

    Grid_Cat_Tipos_Cursos.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tipos_Cursos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " OR Clave LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Tipos_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Tipos_Cursos
        If Not .EOF Then
            Grid_Cat_Tipos_Cursos.AddItem "Tipo ID" & Chr(9) & "Clave" & Chr(9) & "Nombre"
            While Not .EOF
                Grid_Cat_Tipos_Cursos.AddItem .rdoColumns("Tipo_Curso_Id") & Chr(9) & .rdoColumns("Clave") & Chr(9) & .rdoColumns("Nombre")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Tipos_Cursos.FixedRows = 1
            Grid_Cat_Tipos_Cursos.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Tipos_Cursos.ColWidth(1) = 1200   'Clave
            Grid_Cat_Tipos_Cursos.ColWidth(2) = 4500   'Nombre
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Tipos_Cursos = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o Clave para buscar los tipos")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Cat_Tipos_Cursos Nombre
       
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo Fin
    If Txt_Cat_Tipos_Cursos_Tipo_Id.Text <> "" Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Mi_SQL = "DELETE FROM Cat_Tipos_Cursos WHERE Tipo_Curso_Id='" & Trim(Txt_Cat_Tipos_Cursos_Tipo_Id.Text) & "'"
            Conexion_Base.Execute Mi_SQL
            'Quita los datos de la sala contenidos en el Grid
            If Grid_Cat_Tipos_Cursos.Rows = 2 Then
                Grid_Cat_Tipos_Cursos.Rows = 0
            Else
                Grid_Cat_Tipos_Cursos.RemoveItem Grid_Cat_Tipos_Cursos.RowSel
            End If 'Grid_productos
            MsgBox "Tipo de curso eliminado", vbInformation + vbOKOnly, Me.Caption
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
        If Txt_Cat_Tipos_Cursos_Tipo_Id.Text <> "" Then
        Call Configurar_Formulario(True)
            Btn_Modificar.Enabled = True
            Btn_Modificar.Caption = "Guardar"
            Txt_Cat_Tipos_Cursos_Clave.SetFocus
        Else
            MsgBox ("Es necesario seleccionar un registro para modificar")
        End If
    Else
        Modificar_Cat_Tipos_Cursos
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
        Txt_Cat_Tipos_Cursos_Clave.SetFocus
        Fra_Tipo_Cursos.Enabled = False
    Else
        If Validar_Componentes Then
            Call Alta_Tipos_Cursos
            Limpiar_Formulario
            Btn_Nuevo.Caption = "Nuevo"
            Configurar_Formulario (False)
            Btn_Salir.Caption = "Salir"
            Fra_Tipo_Cursos.Enabled = True
        Else
            MsgBox ("Todos los campos marcados con * son necesarios")
        End If
    End If
End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
    Fra_Generales_Cat_Tipos_Cursos.Enabled = Habilitar
    Btn_Nuevo.Enabled = Not Habilitar
    Btn_Modificar.Enabled = Not Habilitar
    Btn_Eliminar.Enabled = Not Habilitar
    Btn_Buscar.Enabled = Not Habilitar
    Btn_Salir.Caption = "Cancelar"
End Sub
Function Validar_Componentes() As Boolean
Validar_Componentes = True
If Txt_Cat_Tipos_Cursos_Clave.Text = "" Then
Validar_Componentes = False
End If
If Txt_Cat_Tipos_Cursos_Nombre.Text = "" Then
Validar_Componentes = False
End If

End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Tipos_Cursos
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del tipo de curso que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laura Huichapa Ramírez
    'FECHA_CREO : 22-Diciembre-2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Tipos_Cursos()
'Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Tipos_Cursos As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans
    
    'Alta de Institución
    Set Rs_Alta_Cat_Tipos_Cursos = Conectar_Ayudante.Recordset_Agregar("Cat_Tipos_Cursos")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Tipos_Cursos
        .AddNew
            Txt_Cat_Tipos_Cursos_Tipo_Id.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Cursos", "Tipo_Curso_Id"), "00000")
            .rdoColumns("Tipo_Curso_Id") = Txt_Cat_Tipos_Cursos_Tipo_Id.Text
            .rdoColumns("Clave") = Trim(Txt_Cat_Tipos_Cursos_Clave)
            .rdoColumns("Nombre") = UCase(Txt_Cat_Tipos_Cursos_Nombre.Text)
            .rdoColumns("Descripcion") = UCase(Txt_Cat_Tipos_Cursos_Descripcion.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Tipos_Cursos.Close
    Conexion_Base.CommitTrans
    MsgBox "Registro agregado", vbInformation
    Consulta_Cat_Tipos_Cursos ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Cat_Tipos_Cursos_Tipo_Id.Text = ""
Txt_Cat_Tipos_Cursos_Clave.Text = ""
Txt_Cat_Tipos_Cursos_Nombre.Text = ""
Txt_Cat_Tipos_Cursos_Descripcion.Text = ""
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
    Fra_Tipo_Cursos.Enabled = True
End If
    
End Sub

Private Sub Grid_Cat_Tipos_Cursos_Click()
Dim Rs_Consulta_Cat_Tipos_Cursos As rdoResultset
    If Grid_Cat_Tipos_Cursos.Rows > 1 Then
        Txt_Cat_Tipos_Cursos_Tipo_Id.Text = Grid_Cat_Tipos_Cursos.TextMatrix(Grid_Cat_Tipos_Cursos.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Tipos_Cursos"
        Mi_SQL = Mi_SQL & "  WHERE Tipo_Curso_Id='" & Txt_Cat_Tipos_Cursos_Tipo_Id.Text & "'"
        Set Rs_Consulta_Cat_Tipos_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Tipos_Cursos.EOF Then
            With Rs_Consulta_Cat_Tipos_Cursos
                Txt_Cat_Tipos_Cursos_Tipo_Id.Text = .rdoColumns("Tipo_Curso_Id")
                Txt_Cat_Tipos_Cursos_Clave.Text = .rdoColumns("Clave")
                Txt_Cat_Tipos_Cursos_Nombre.Text = .rdoColumns("Nombre")
                Txt_Cat_Tipos_Cursos_Descripcion.Text = .rdoColumns("Descripcion")
            End With
        End If
        Rs_Consulta_Cat_Tipos_Cursos.Close
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Tipos_Cursos
    'DESCRIPCIÓN:           Modifica el registro de la tabla tipos de cursos
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    22 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Cat_Tipos_Cursos()
Dim Rs_Modificacion_Cat_Tipos_Cursos As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tipos_Cursos"
    Mi_SQL = Mi_SQL & " WHERE Tipo_Curso_Id ='" & Trim(Txt_Cat_Tipos_Cursos_Tipo_Id.Text) & "'"
    Set Rs_Modificacion_Cat_Tipos_Cursos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Tipos_Cursos
        .Edit
            .rdoColumns("Clave") = Trim(Txt_Cat_Tipos_Cursos_Clave.Text)
            .rdoColumns("Nombre") = Trim(Txt_Cat_Tipos_Cursos_Nombre.Text)
            .rdoColumns("Descripcion") = Trim(Txt_Cat_Tipos_Cursos_Descripcion.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Tipos_Cursos = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "El registro ha sido modificado", vbInformation + vbOKOnly, Me.Caption
   Consulta_Cat_Tipos_Cursos ""
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

