VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Ope_Programacion_Cursos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROGRAMACIÓN DE CURSOS"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   7560
   Begin VB.PictureBox Pic_Ope_Programacion_Cursos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   8400
      TabIndex        =   18
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Generales_Ope_Programacion_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   2895
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   7200
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Instructor 
            Height          =   315
            ItemData        =   "Frm_Ope_Programacion_Cursos.frx":0000
            Left            =   1200
            List            =   "Frm_Ope_Programacion_Cursos.frx":000A
            TabIndex        =   5
            Top             =   1320
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Institucion 
            Height          =   315
            ItemData        =   "Frm_Ope_Programacion_Cursos.frx":0020
            Left            =   1200
            List            =   "Frm_Ope_Programacion_Cursos.frx":002A
            TabIndex        =   4
            Top             =   960
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Sala 
            Height          =   315
            ItemData        =   "Frm_Ope_Programacion_Cursos.frx":0040
            Left            =   1200
            List            =   "Frm_Ope_Programacion_Cursos.frx":004A
            TabIndex        =   6
            Top             =   1680
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Curso 
            Height          =   315
            ItemData        =   "Frm_Ope_Programacion_Cursos.frx":0060
            Left            =   1200
            List            =   "Frm_Ope_Programacion_Cursos.frx":006A
            TabIndex        =   3
            Top             =   600
            Width           =   5850
         End
         Begin VB.TextBox Txt_Ope_Programacion_Cusos_No_Programa_Curso 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Estatus 
            Height          =   315
            ItemData        =   "Frm_Ope_Programacion_Cursos.frx":0080
            Left            =   4680
            List            =   "Frm_Ope_Programacion_Cursos.frx":008A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2370
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Fecha_Fin 
            Height          =   315
            Left            =   4680
            TabIndex        =   8
            Top             =   2040
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Format          =   108658689
            CurrentDate     =   42373
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Hora_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   2400
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Format          =   108658690
            CurrentDate     =   42373
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Hora_Fin 
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Top             =   2400
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Format          =   108658690
            CurrentDate     =   42373
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   2040
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Format          =   108658689
            CurrentDate     =   42373
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Hora Fin"
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
            TabIndex        =   30
            Top             =   2490
            Width           =   810
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*F. Fin"
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
            TabIndex        =   29
            Top             =   2130
            Width           =   570
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Curso"
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
            Left            =   90
            TabIndex        =   27
            Top             =   690
            Width           =   570
         End
         Begin VB.Label Lbl_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Sala"
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
            Left            =   90
            TabIndex        =   26
            Top             =   1770
            Width           =   465
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "No Curso"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   25
            Top             =   330
            Width           =   660
         End
         Begin VB.Label Lbl_Estado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*F. Inicio"
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
            Left            =   90
            TabIndex        =   24
            Top             =   2130
            Width           =   780
         End
         Begin VB.Label Lbl_Direccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Institucion"
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
            Left            =   90
            TabIndex        =   23
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Lbl_Ciudad 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Instructor"
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
            Left            =   90
            TabIndex        =   22
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Hora Inicio"
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
            Left            =   90
            TabIndex        =   21
            Top             =   2490
            Width           =   1020
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
            TabIndex        =   20
            Top             =   330
            Width           =   720
         End
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4560
         Picture         =   "Frm_Ope_Programacion_Cursos.frx":00A0
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "C"
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   240
         Picture         =   "Frm_Ope_Programacion_Cursos.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "A"
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   6000
         Picture         =   "Frm_Ope_Programacion_Cursos.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3120
         Picture         =   "Frm_Ope_Programacion_Cursos.frx":03A6
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "B"
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1680
         Picture         =   "Frm_Ope_Programacion_Cursos.frx":04A8
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "M"
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Ope_Grid_Programacion_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Programas Cursos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4080
         Left            =   120
         TabIndex        =   31
         Top             =   3360
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Ope_Programacion_Cursos 
            Height          =   3705
            Left            =   75
            TabIndex        =   0
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   6535
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
      Begin VB.Frame Fra_Ope_Programacion_Invitacion_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invitacion Empleados"
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
         Height          =   4080
         Left            =   120
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   7185
         Begin VB.CommandButton Btn_Eliminar_Inivtado 
            Appearance      =   0  'Flat
            Caption         =   "Quitar"
            Height          =   255
            Left            =   4560
            TabIndex        =   33
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton Btn_Agregar_Invitado 
            Caption         =   "Agregar"
            Height          =   255
            Left            =   5880
            TabIndex        =   12
            Top             =   3720
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Ope_Programacion_Invitacion_Empleados 
            Height          =   3360
            Left            =   75
            TabIndex        =   11
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5927
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Label Lbl_Programación_Cursos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PROGRAMACION DE CURSOS"
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
         Left            =   1425
         TabIndex        =   28
         Top             =   15
         Width           =   5535
      End
   End
End
Attribute VB_Name = "Frm_Ope_Programacion_Cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Empleado_Seleccion_Id As String

Public Sub Inicializa()
Consulta_Ope_Programacion_Cursos ""
Empleado_Seleccion_Id = ""
 Call Conectar_Ayudante.Llena_Combo_Item("Curso_Id, Nombre", "Cat_Cursos_Capacitaciones WHERE Estatus='ACTIVO'", Cmb_Ope_Prog_Cursos_Curso, 0, "Curso_Id", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Sala_Id, Nombre", "Cat_Salas WHERE Estatus='ACTIVO'", Cmb_Ope_Prog_Cursos_Sala, 0, "Sala_Id", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Institucion_Id, Nombre", "Cat_Instituciones WHERE Estatus='ACTIVO'", Cmb_Ope_Prog_Cursos_Institucion, 0, "Institucion_Id", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Instructor_Id, Nombre", "Cat_Instructores WHERE Estatus='ACTIVO'", Cmb_Ope_Prog_Cursos_Instructor, 0, "Instructor_Id", "", False, "")
Dt_Ope_Prog_Cursos_Hora_Inicio.Value = Time
Dt_Ope_Prog_Cursos_Hora_Fin.Value = Time
'Set Grid_Ope_Programacion_Invitacion_Empleados.DataSource = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Ope_Programacion_Cursos
    'DESCRIPCIÓN:           Consulta los Cursos programados y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre del curso
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Ope_Programacion_Cursos(Nombre As String)
Dim Rs_Consulta_Ope_Programacion_Cursos As rdoResultset       'Informacion de los registros

    Grid_Ope_Programacion_Cursos.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Ope_Programacion_Cursos.*, Cat_Cursos_Capacitaciones.Nombre as Curso, "
    Mi_SQL = Mi_SQL & " Cat_Instituciones.Nombre as Institucion, Cat_Instructores.Nombre as Instructor,Cat_Salas.Nombre as Sala "
    Mi_SQL = Mi_SQL & " FROM Ope_Programacion_Cursos, Cat_Cursos_Capacitaciones, Cat_Instituciones, Cat_Instructores,Cat_Salas"
    Mi_SQL = Mi_SQL & " WHERE Cat_Cursos_Capacitaciones.Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " AND  Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id"
    Mi_SQL = Mi_SQL & " AND Cat_Instituciones.Institucion_Id = Ope_Programacion_Cursos.Institucion_Id"
    Mi_SQL = Mi_SQL & " and Cat_Instructores.Instructor_Id = Ope_Programacion_Cursos.Instructor_Id"
    Mi_SQL = Mi_SQL & " and Ope_Programacion_Cursos.Sala_ID = Cat_Salas.Sala_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Estatus='ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Estatus='ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Cat_Instituciones.Estatus='ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Cat_Instructores.Estatus='ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Cat_Salas.Estatus='ACTIVO'"
    Set Rs_Consulta_Ope_Programacion_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Ope_Programacion_Cursos
        If Not .EOF Then
            Grid_Ope_Programacion_Cursos.AddItem "No Curso" & Chr(9) & "Nombre" & Chr(9) & "Institucion" & Chr(9) & "Sala"
            While Not .EOF
                Grid_Ope_Programacion_Cursos.AddItem .rdoColumns("No_Programa_Curso") & Chr(9) & .rdoColumns("Curso") & Chr(9) & .rdoColumns("Institucion") & Chr(9) & .rdoColumns("Sala")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Ope_Programacion_Cursos.FixedRows = 1
            Grid_Ope_Programacion_Cursos.ColWidth(0) = 800     'Intitución_ID
            Grid_Ope_Programacion_Cursos.ColWidth(1) = 4000   'curso
            Grid_Ope_Programacion_Cursos.ColWidth(2) = 4000   'institucion
            Grid_Ope_Programacion_Cursos.ColWidth(3) = 1300  'sala
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Programacion_Cursos = Nothing

End Sub

Private Sub Btn_Agregar_Invitado_Click()
Load Frm_Aux_Listar_Empleados
Frm_Aux_Listar_Empleados.Inicializa
'Unload Me

End Sub

Private Sub Btn_Buscar_Click()
    Dim Nombre As String 'Obtiene el nombre a consultar
    Nombre = InputBox("Proporcione el Nombre curso para buscar los cursos programados")
    Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
    Consulta_Ope_Programacion_Cursos Nombre
End Sub

Private Sub Btn_Eliminar_Click()
If Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text <> "" Then
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'        If Conectar_Ayudante.Elimina_Catalogo("Ope_Lista_Invitados", "No_Programa_Curso", Trim(Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text)) = True Then
'            If Conectar_Ayudante.Elimina_Catalogo("Ope_Programacion_Cursos", "No_Programa_Curso", Trim(Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text)) = True Then                        'Quita los datos del usuario contenidos en el Grid
                Modificar_Ope_Programacion_Cursos "CANCELADO"
                If Grid_Ope_Programacion_Cursos.Rows = 2 Then
                    Grid_Ope_Programacion_Cursos.Rows = 0
                Else
                    Grid_Ope_Programacion_Cursos.RemoveItem Grid_Ope_Programacion_Cursos.RowSel
                End If 'Grid_productos
                MsgBox "Registro Eliminado", vbInformation + vbOKOnly, Me.Caption
'            Else
'                MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
'            End If
'        End If
    End If
Else
    MsgBox ("Es necesario seleccionar un registro para eliminar")
End If
End Sub

Private Sub Btn_Eliminar_Inivtado_Click()
If (Grid_Ope_Programacion_Invitacion_Empleados.Rows > 0) Then
    If Grid_Ope_Programacion_Invitacion_Empleados.Rows = 2 Then
        Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0
    Else
        Grid_Ope_Programacion_Invitacion_Empleados.RemoveItem Grid_Ope_Programacion_Invitacion_Empleados.RowSel
    End If
Else
    MsgBox "Debes Seleccionar un empleado de la lista.", vbInformation, "Mensaje"
End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
    If Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text <> "" Then
        Call Configurar_Formulario(True)
        Cargar_Grid_Empleados_Invitados
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Guardar"
    Else
        MsgBox ("Es necesario seleccionar un registro para modificar")
    End If
Else
 On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Me.MousePointer = 11
    Modificar_Ope_Programacion_Cursos Trim(Cmb_Ope_Prog_Cursos_Estatus.Text)
    For I = 1 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1
        If Not Existe(Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 0)) And Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 0) <> "No_Tarjeta" Then
            Call Alta_Lista_Invitados(Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4))
        End If
    Next I

    Dim Rs_Consulta_Ope_Lista_Invitados_Registrados As rdoResultset       'Informacion de los registros
    Mi_SQL = "SELECT Ope_Lista_Invitados.* "
    Mi_SQL = Mi_SQL & " FROM Ope_Lista_Invitados "
    Mi_SQL = Mi_SQL & " WHERE No_Programa_Curso = " & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & " "
    Set Rs_Consulta_Ope_Lista_Invitados_Registrados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Ope_Lista_Invitados_Registrados.EOF
        If Not Existe_Grid(Rs_Consulta_Ope_Lista_Invitados_Registrados.rdoColumns("Empleado_Id")) Then
            Eliminar_Invitacion_Empleado (Rs_Consulta_Ope_Lista_Invitados_Registrados.rdoColumns("Empleado_Id"))
'            Dim Rs_Consutlta_No_Invitacion As rdoResultset
'            Mi_SQL = "SELECT * FROM Ope_Lista_Invitados where No_Programa_Curso = " & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & " AND Empleado_Id = " & Rs_Consulta_Ope_Lista_Invitados_Registrados.rdoColumns("Empleado_Id")
'            Set Rs_Consutlta_No_Invitacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'            Dim Resultado As Boolean
'            Resultado = Conectar_Ayudante.Elimina_Catalogo("Ope_Lista_Invitados", "No_Lista_Invitado", Rs_Consutlta_No_Invitacion.rdoColumns("No_Lista_Invitado"))
        End If
        Rs_Consulta_Ope_Lista_Invitados_Registrados.MoveNext
    Wend
    Rs_Consulta_Ope_Lista_Invitados_Registrados.Close
    Call Enviar_Correos
    Conexion_Base.CommitTrans
    MsgBox ("Programación del curso Modificada, correos enviados")
    Me.MousePointer = 0
    Limpiar_Formulario
    Btn_Modificar.Caption = "Modificar"
    Configurar_Formulario (False)
    Btn_Salir.Caption = "Salir"
End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    Me.MousePointer = 0
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
    Call Configurar_Formulario(True)
    Limpiar_Formulario
    Btn_Nuevo.Enabled = True
    Btn_Nuevo.Caption = "Guardar"
Else
    If Validar_Componentes Then
    On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Me.MousePointer = 11
        Call Alta_Ope_Programacion_Cursos
        For I = 1 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1
            If Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4) <> "Empleado_Id" Then
                Call Alta_Lista_Invitados(Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4))
'                Call Enviar_Correos(Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4))
            End If
        Next I
        Call Enviar_Correos
    Conexion_Base.CommitTrans
    MsgBox ("Programación del curso realizada, correos enviados")
    Me.MousePointer = 0
        Limpiar_Formulario
        Btn_Nuevo.Caption = "Nuevo"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"

    Else
        MsgBox ("Todos los campos marcados con * son necesarios")
    End If
End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    Me.MousePointer = 0
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
    Fra_Generales_Ope_Programacion_Cursos.Enabled = Habilitar
    Fra_Ope_Programacion_Invitacion_Empleados.Enabled = Habilitar
    Fra_Ope_Grid_Programacion_Cursos.Visible = Not Habilitar
    Fra_Ope_Programacion_Invitacion_Empleados.Visible = Habilitar
    Btn_Nuevo.Enabled = Not Habilitar
    Btn_Modificar.Enabled = Not Habilitar
    Btn_Eliminar.Enabled = Not Habilitar
    Btn_Buscar.Enabled = Not Habilitar
    Btn_Salir.Caption = "Cancelar"
End Sub
Function Validar_Componentes() As Boolean
Validar_Componentes = True
If Cmb_Ope_Prog_Cursos_Curso.ListIndex = -1 Then
Validar_Componentes = False
End If
If Cmb_Ope_Prog_Cursos_Estatus.ListIndex = -1 Then
Validar_Componentes = False
End If
If Cmb_Ope_Prog_Cursos_Institucion.ListIndex = -1 Then
Validar_Componentes = False
End If
If Cmb_Ope_Prog_Cursos_Instructor.Text = "" Then
Validar_Componentes = False
End If
If Cmb_Ope_Prog_Cursos_Sala.Text = "" Then
Validar_Componentes = False
End If
If Dt_Ope_Prog_Cursos_Fecha_Inicio.Value = "" Then
Validar_Componentes = False
End If
If Dt_Ope_Prog_Cursos_Fecha_Fin.Value = "" Then
Validar_Componentes = False
End If
If Dt_Ope_Prog_Cursos_Hora_Inicio.Value = "" Then
Validar_Componentes = False
End If
If Dt_Ope_Prog_Cursos_Hora_Fin.Value = "" Then
Validar_Componentes = False
End If
If Dt_Ope_Prog_Cursos_Hora_Inicio.Value > Dt_Ope_Prog_Cursos_Hora_Fin.Value Then
Validar_Componentes = False
End If

End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Ope_Programacion_Cursos
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del curso que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 04 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Ope_Programacion_Cursos()
'Dim Menus As Integer
'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Ope_Programacion_Cursos As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

'On Error GoTo HANDLER
'    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans

    'Alta de Institución
    Set Rs_Alta_Ope_Programacion_Cursos = Conectar_Ayudante.Recordset_Agregar("Ope_Programacion_Cursos")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Ope_Programacion_Cursos
        .AddNew
            Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Ope_Programacion_Cursos", "No_Programa_Curso"), "0000000000")
            .rdoColumns("No_Programa_Curso") = Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text
            .rdoColumns("Curso_Id") = Format(Cmb_Ope_Prog_Cursos_Curso.ItemData(Cmb_Ope_Prog_Cursos_Curso.ListIndex), "00000")
            .rdoColumns("Sala_Id") = Format(Cmb_Ope_Prog_Cursos_Sala.ItemData(Cmb_Ope_Prog_Cursos_Sala.ListIndex), "00000")
            .rdoColumns("Institucion_Id") = Format(Cmb_Ope_Prog_Cursos_Institucion.ItemData(Cmb_Ope_Prog_Cursos_Institucion.ListIndex), "00000")
            .rdoColumns("Instructor_Id") = Format(Cmb_Ope_Prog_Cursos_Instructor.ItemData(Cmb_Ope_Prog_Cursos_Instructor.ListIndex), "00000")
            .rdoColumns("Estatus") = Cmb_Ope_Prog_Cursos_Estatus.Text
'            .rdoColumns("Fecha_Inicio") = Dt_Ope_Prog_Cursos_Fecha_Inicio.Value
            .rdoColumns("Fecha_Inicio") = Dt_Ope_Prog_Cursos_Fecha_Inicio.Value
            .rdoColumns("Fecha_Fin") = Dt_Ope_Prog_Cursos_Fecha_Fin
            .rdoColumns("Hora_Inicio") = Dt_Ope_Prog_Cursos_Hora_Inicio.Value
            .rdoColumns("Hora_Fin") = Dt_Ope_Prog_Cursos_Hora_Fin.Value
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Ope_Programacion_Cursos.Close
'    Conexion_Base.CommitTrans
'    MsgBox "Programacion de curso agregada", vbInformation
'    Consulta_Ope_Programacion_Cursos ""
'    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
'HANDLER:
'    Conexion_Base.RollbackTrans
'    For Each Er In rdoErrors
'        MsgBox Er.Description
'    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = ""
'Set Grid_Ope_Programacion_Invitacion_Empleados.DataSource = Nothing
Cmb_Ope_Prog_Cursos_Curso.ListIndex = -1
Cmb_Ope_Prog_Cursos_Estatus.ListIndex = -1
Cmb_Ope_Prog_Cursos_Institucion.ListIndex = -1
Cmb_Ope_Prog_Cursos_Instructor.ListIndex = -1
Cmb_Ope_Prog_Cursos_Sala.ListIndex = -1
Dt_Ope_Prog_Cursos_Fecha_Inicio.Value = Date
Dt_Ope_Prog_Cursos_Fecha_Fin.Value = Date
Dt_Ope_Prog_Cursos_Hora_Inicio.Value = Time
Dt_Ope_Prog_Cursos_Hora_Fin.Value = Time
Grid_Ope_Programacion_Invitacion_Empleados.Clear
Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0
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
End If

End Sub

Private Sub Grid_Ope_Programacion_Cursos_Click()
Dim Rs_Consulta_Ope_Programacion_Cursos As rdoResultset
    If Grid_Ope_Programacion_Cursos.Rows > 1 Then
        Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = Grid_Ope_Programacion_Cursos.TextMatrix(Grid_Ope_Programacion_Cursos.RowSel, 0)
        Mi_SQL = "SELECT * FROM Ope_Programacion_Cursos"
        Mi_SQL = Mi_SQL & "  WHERE No_Programa_Curso='" & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & "'"
        Set Rs_Consulta_Ope_Programacion_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Ope_Programacion_Cursos.EOF Then
            With Rs_Consulta_Ope_Programacion_Cursos
              'CURSO
                 If Not IsNull(.rdoColumns("Curso_Id")) Then
'
                    Cmb_Ope_Prog_Cursos_Curso.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Curso_Id"), "Cat_Cursos_Capacitaciones", "Nombre", "Curso_Id"), Cmb_Ope_Prog_Cursos_Curso)
                Else
                    Cmb_Ope_Prog_Cursos_Curso.ListIndex = -1
                End If
                'SALA
                 If Not IsNull(.rdoColumns("Sala_Id")) Then
''
                    Cmb_Ope_Prog_Cursos_Sala.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Sala_Id"), "Cat_Salas", "Nombre", "Sala_Id"), Cmb_Ope_Prog_Cursos_Sala)
                Else
                    Cmb_Ope_Prog_Cursos_Sala.ListIndex = -1
                End If
                 'INSTITUCION
                 If Not IsNull(.rdoColumns("Institucion_Id")) Then
''
                    Cmb_Ope_Prog_Cursos_Institucion.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Institucion_Id"), "Cat_Instituciones", "Nombre", "Institucion_Id"), Cmb_Ope_Prog_Cursos_Institucion)
                Else
                    Cmb_Ope_Prog_Cursos_Institucion.ListIndex = -1
                End If
                'INSTRUCTOR
                 If Not IsNull(.rdoColumns("Instructor_Id")) Then
                    Cmb_Ope_Prog_Cursos_Instructor.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Instructor_Id"), "Cat_Instructores", "Nombre", "Instructor_Id"), Cmb_Ope_Prog_Cursos_Instructor)
                Else
                    Cmb_Ope_Prog_Cursos_Instructor.ListIndex = -1
                End If
             If Not IsNull(.rdoColumns("Fecha_Inicio")) Then
             Dt_Ope_Prog_Cursos_Fecha_Inicio = .rdoColumns("Fecha_Inicio")
             End If
              If Not IsNull(.rdoColumns("Fecha_Fin")) Then
             Dt_Ope_Prog_Cursos_Fecha_Fin = .rdoColumns("Fecha_Fin")
             End If
              If Not IsNull(.rdoColumns("Hora_Inicio")) Then
             Dt_Ope_Prog_Cursos_Hora_Inicio.Value = Format(.rdoColumns("Hora_Inicio"), "MM/dd/yyyy hh:mm:ss")
            End If
              If Not IsNull(.rdoColumns("Hora_Fin")) Then
             Dt_Ope_Prog_Cursos_Hora_Fin = Format(.rdoColumns("Hora_Fin"), "MM/dd/yyyy hh:mm:ss")
             End If
     If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Ope_Prog_Cursos_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Estatus")), Cmb_Ope_Prog_Cursos_Estatus)
                Else
                    Cmb_Ope_Prog_Cursos_Estatus.ListIndex = -1
                End If
        End With
        Rs_Consulta_Ope_Programacion_Cursos.Close
    End If
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Ope_Programacion_Cursos
    'DESCRIPCIÓN:           Modifica el registro de la tabla
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    04 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Ope_Programacion_Cursos(Estatus As String)
Dim Rs_Modificacion_Oe_Programacion_Cursos As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Ope_Programacion_Cursos"
    Mi_SQL = Mi_SQL & " WHERE No_Programa_Curso ='" & Trim(Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text) & "'"
    Set Rs_Modificacion_Oe_Programacion_Cursos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    If (Validar_Componentes) Then
    With Rs_Modificacion_Oe_Programacion_Cursos
        .Edit
            .rdoColumns("Curso_Id") = Format(Cmb_Ope_Prog_Cursos_Curso.ItemData(Cmb_Ope_Prog_Cursos_Curso.ListIndex), "00000")
            .rdoColumns("Sala_Id") = Format(Cmb_Ope_Prog_Cursos_Sala.ItemData(Cmb_Ope_Prog_Cursos_Sala.ListIndex), "00000")
            .rdoColumns("Institucion_Id") = Format(Cmb_Ope_Prog_Cursos_Institucion.ItemData(Cmb_Ope_Prog_Cursos_Institucion.ListIndex), "00000")
            .rdoColumns("Instructor_Id") = Format(Cmb_Ope_Prog_Cursos_Instructor.ItemData(Cmb_Ope_Prog_Cursos_Instructor.ListIndex), "00000")
            .rdoColumns("Fecha_Inicio") = Dt_Ope_Prog_Cursos_Fecha_Inicio
            .rdoColumns("Fecha_Fin") = Dt_Ope_Prog_Cursos_Fecha_Fin
            .rdoColumns("Hora_Inicio") = Dt_Ope_Prog_Cursos_Hora_Inicio.Value
            .rdoColumns("Hora_Fin") = Dt_Ope_Prog_Cursos_Hora_Fin.Value
            .rdoColumns("Estatus") = Estatus
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
   End If
    Set Rs_Modificacion_Oe_Programacion_Cursos = Nothing
    'Agrega los checadores

    Conexion_Base.CommitTrans
'   MsgBox "El curso programado ha sido modificado", vbInformation + vbOKOnly, Me.Caption
   Consulta_Ope_Programacion_Cursos ""

    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Public Sub Llenar_Grid_Invitados()
If Empleado_Seleccion_Id <> "" Then
If Verificar_Id_Empleado Then
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros

'    Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM  Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Empleado_Id = '" & Empleado_Seleccion_Id & "'"
'    MsgBox Mi_SQL
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Empleados
        If Not .EOF Then
            If Grid_Ope_Programacion_Invitacion_Empleados.Rows <= 0 Then
            Grid_Ope_Programacion_Invitacion_Empleados.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno" & Chr(9) & "Empleado_ID"
           End If
           While Not .EOF
                Grid_Ope_Programacion_Invitacion_Empleados.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno") & Chr(9) & .rdoColumns("Empleado_ID")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Ope_Programacion_Invitacion_Empleados.FixedRows = 1
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(0) = 800     'Intitución_ID
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(1) = 2000   'Nombre
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(2) = 2000   'Clave
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(3) = 2000
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(4) = 0
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Empleados = Nothing
    End If
Else

End If
Empleado_Seleccion_Id = ""
End Sub

Function Verificar_Id_Empleado() As Boolean
Verificar_Id_Empleado = True
Dim I As Integer
Dim Id_Emplado_Grid As String
For I = 0 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1

Id_Emplado_Grid = Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4)
If Empleado_Seleccion_Id = Id_Emplado_Grid Then
Verificar_Id_Empleado = False
If Id_Emplado_Grid <> "Empleado_Id" Then
'MsgBox ("El empleado ya ha sido agregado anteriormente")
End If
Load Frm_Aux_Listar_Empleados

'--------------------------------------------
Dim Empresa As String 'Obtiene el nombre a consultar
Dim Departamento As String 'Obtiene el nombre a consultar
Dim Puesto As String 'Obtiene el nombre a consultar
Dim Turno As String 'Obtiene el nombre a consultar
If Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Empresa.ListIndex > -1 Then
Empresa = Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Empresa.Text
End If
If Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Departamento.ListIndex > -1 Then
Departamento = Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Departamento.Text
End If
If Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Puesto.ListIndex > -1 Then
Puesto = Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Puesto.Text
End If
If Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Turno.ListIndex > -1 Then
Turno = Frm_Aux_Listar_Empleados.Cmb_Filtro_Empleado_Turno.Text
End If
       Frm_Aux_Listar_Empleados.Consulta_Cat_Empleados Empresa, Departamento, Puesto, Turno
'.-----------------------------------------
'Frm_Aux_Listar_Empleados.Inicializa
Exit For
End If
Next I
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Lista_Invitados
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del curso que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 05 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Lista_Invitados(Empleado_ID As String)
'For i = 1 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1

'Id_Emplado_Grid = Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(i, 0)
'Dim Menus As Integer
'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Lista_Invitados As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

'On Error GoTo HANDLER
'    Conexion_Base.BeginTrans
''    Conexion_Servidor.BeginTrans

    'Alta de Institución
    Set Rs_Alta_Lista_Invitados = Conectar_Ayudante.Recordset_Agregar("Ope_Lista_Invitados")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Lista_Invitados
        .AddNew
            .rdoColumns("No_Programa_Curso") = Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text
            '.rdoColumns("Empleado_Id") = Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 0)
            .rdoColumns("Empleado_Id") = Empleado_ID
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Lista_Invitados.Close
'    Conexion_Base.CommitTrans
'    Next i
'    MsgBox "Programacion de curso agregada", vbInformation
    Consulta_Ope_Programacion_Cursos ""
'    Exit Sub
'
''Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
'HANDLER:
'    Conexion_Base.RollbackTrans
'    For Each Er In rdoErrors
'        MsgBox Er.Description
'    Next Er

End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Cargar_Grid_Empleados_Invitados
    'DESCRIPCIÓN:           Consulta los Cursos programados y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre del curso
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           05 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Cargar_Grid_Empleados_Invitados()
Dim Rs_Consulta_Ope_Empleados_Invitados As rdoResultset       'Informacion de los registros

    Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Cat_Empleados.* "
    Mi_SQL = Mi_SQL & "FROM Ope_Lista_Invitados, Cat_Empleados "
    Mi_SQL = Mi_SQL & "WHERE No_Programa_Curso = '" & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & "' "
    Mi_SQL = Mi_SQL & "AND Cat_Empleados.Empleado_ID = Ope_Lista_Invitados.Empleado_Id"
'    MsgBox Mi_SQL

    Set Rs_Consulta_Ope_Empleados_Invitados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Ope_Empleados_Invitados
        If Not .EOF Then

            Grid_Ope_Programacion_Invitacion_Empleados.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno" & Chr(9) & "Empleado_ID"
            While Not .EOF
                Grid_Ope_Programacion_Invitacion_Empleados.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno") & Chr(9) & .rdoColumns("Empleado_ID")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Ope_Programacion_Invitacion_Empleados.FixedRows = 1
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(0) = 800     'Intitución_ID
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(1) = 6000   'Nombre
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(2) = 1800   'Clave
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(3) = 1800
            Grid_Ope_Programacion_Invitacion_Empleados.ColWidth(4) = 0

            .Close

        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Grid_Ope_Programacion_Invitacion_Empleados = Nothing

End Sub

Private Sub Grid_Ope_Programacion_Invitacion_Empleados_Click()
If Grid_Ope_Programacion_Invitacion_Empleados.ColSel = 4 Then
Grid_Ope_Programacion_Invitacion_Empleados.RemoveItem (Grid_Ope_Programacion_Invitacion_Empleados.RowSel)
End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Enviar_Correos
    'DESCRIPCIÓN:           Consulta los registos del grid para enviar el correo a cada empleado
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           06 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Enviar_Correos()
Dim Temario As String
Dim Objetivo As String
Dim Materiales As String
Dim Descripcion As String

  Frm_Email.Obtener_Parametros_Correos

    Dim Rs_Consulta_Cursos As rdoResultset
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & " FROM Cat_Cursos_Capacitaciones"
    Mi_SQL = Mi_SQL & " WHERE Curso_Id = '" & Format(Cmb_Ope_Prog_Cursos_Curso.ItemData(Cmb_Ope_Prog_Cursos_Curso.ListIndex), "00000") & "'"
    Set Rs_Consulta_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cursos
        If Not .EOF Then
            While Not .EOF
                 Temario = .rdoColumns("Temario")
                 Objetivo = .rdoColumns("Objetivo")
                 Materiales = .rdoColumns("Lista_Material")
                 Descripcion = .rdoColumns("Descripcion")
                .MoveNext
            Wend
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cursos = Nothing
    For I = 1 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1
        Dim Usuariooo As String
        Usuariooo = Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 0)
        Dim Rs_Consulta_Ope_Email As rdoResultset       'Informacion de los registros
        'Consulta los datos generales del usuario
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE No_Tarjeta = " & Usuariooo
        Set Rs_Consulta_Ope_Email = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Ope_Email
        If Not .EOF Then
            While Not .EOF
             If Not IsNull(.rdoColumns("Email")) Then
                Frm_Email.Correo = .rdoColumns("Email")
                End If
                Frm_Email.Asunto = "Invitación a Curso"
                Frm_Email.Mensaje = "Ha sido invitado a tomar el curso " & Cmb_Ope_Prog_Cursos_Curso.Text & _
                " que se llevará a cabo en " & Cmb_Ope_Prog_Cursos_Sala.Text & "  del día " & Dt_Ope_Prog_Cursos_Fecha_Inicio.Value & _
                " al día " & Dt_Ope_Prog_Cursos_Fecha_Fin.Value & " de " & Dt_Ope_Prog_Cursos_Hora_Inicio.Value & " a " & Dt_Ope_Prog_Cursos_Hora_Fin.Value & _
                " impartido por " & Cmb_Ope_Prog_Cursos_Institucion.Text & "  y el instructor  " & Cmb_Ope_Prog_Cursos_Instructor.Text & _
                 vbLf & "Materiales " & Materiales & vbLf & "Objetivo: " & Objetivo & vbLf & "Temario: " & Temario & vbLf & "Descripción del curso: " & Descripcion
               Dim enviar As Boolean
               If Not IsNull(.rdoColumns("Email")) And .rdoColumns("Email") <> "" Then
               enviar = Frm_Email.Mandar_Correo
               End If
                Unload Frm_Email
                .MoveNext
            Wend
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Email = Nothing

Next I

End Sub


Function Existe(Id_Empleado_Grid As String) As Boolean
Dim Id_Empleado_Tabla As String
Existe = False
Dim Rs_Consulta_Ope_Lista_Invitados As rdoResultset       'Informacion de los registros

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Cat_Empleados.* "
    Mi_SQL = Mi_SQL & " FROM Ope_Lista_Invitados, Cat_Empleados "
    Mi_SQL = Mi_SQL & " WHERE No_Programa_Curso = " & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & " "
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID =Ope_Lista_Invitados.Empleado_Id"
    Set Rs_Consulta_Ope_Lista_Invitados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

   Do While Not Rs_Consulta_Ope_Lista_Invitados.EOF
    Id_Empleado_Tabla = Rs_Consulta_Ope_Lista_Invitados.rdoColumns("No_Tarjeta")
    If Id_Empleado_Grid = Id_Empleado_Tabla Then
     Existe = True
   Exit Do
    End If
   Rs_Consulta_Ope_Lista_Invitados.MoveNext
   Loop
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Programacion_Cursos = Nothing



End Function

Function Existe_Grid(Id_Empleado_Tabla) As Boolean
Dim Id_Empleado_Grid
Existe_Grid = False


For I = 0 To Grid_Ope_Programacion_Invitacion_Empleados.Rows - 1
Id_Empleado_Grid = Grid_Ope_Programacion_Invitacion_Empleados.TextMatrix(I, 4)
If (Id_Empleado_Grid = Id_Empleado_Tabla) Then
Existe_Grid = True
Exit For
End If
Next I

End Function
 Sub Eliminar_Invitacion_Empleado(Empleado_ID As String)
 Dim No_Lista_Invitado As String
            Dim Rs_Consulta_No_Invitacion As rdoResultset
            Mi_SQL = "SELECT * FROM Ope_Lista_Invitados where No_Programa_Curso = " & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & " AND Empleado_Id = " & Empleado_ID
            Set Rs_Consulta_No_Invitacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_No_Invitacion
        If Not .EOF Then
           No_Lista_Invitado = .rdoColumns("No_Lista_Invitado")
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_No_Invitacion = Nothing

Dim Resultado As Boolean
Resultado = Conectar_Ayudante.Elimina_Catalogo_2("Ope_Lista_Invitados", "No_Lista_Invitado", No_Lista_Invitado)
' Conectar_Ayudante.Elimina_Catalogo "Ope_Luis"
 End Sub

