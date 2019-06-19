VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Aux_Listar_Empleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INVITAR EMPLEADOS"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7440
   Begin VB.PictureBox Pic_Ope_Programacion_Cursos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Filtros_Cat_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   7200
         Begin VB.ComboBox Cmb_Filtro_Empleado_Empresa 
            Height          =   315
            ItemData        =   "Frm_Aux_Listar_Empleados.frx":0000
            Left            =   765
            List            =   "Frm_Aux_Listar_Empleados.frx":000A
            TabIndex        =   10
            Top             =   240
            Width           =   2610
         End
         Begin VB.ComboBox Cmb_Filtro_Empleado_Departamento 
            Height          =   315
            ItemData        =   "Frm_Aux_Listar_Empleados.frx":0020
            Left            =   4680
            List            =   "Frm_Aux_Listar_Empleados.frx":002A
            TabIndex        =   9
            Top             =   240
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Filtro_Empleado_Puesto 
            Height          =   315
            ItemData        =   "Frm_Aux_Listar_Empleados.frx":0040
            Left            =   765
            List            =   "Frm_Aux_Listar_Empleados.frx":004A
            TabIndex        =   8
            Top             =   600
            Width           =   2610
         End
         Begin VB.ComboBox Cmb_Filtro_Empleado_Turno 
            Height          =   315
            ItemData        =   "Frm_Aux_Listar_Empleados.frx":0060
            Left            =   4680
            List            =   "Frm_Aux_Listar_Empleados.frx":006A
            TabIndex        =   7
            Top             =   600
            Width           =   2370
         End
         Begin VB.Label Lbl_Ciudad 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Turno"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   14
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Lbl_Direccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Puesto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   13
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Lbl_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Departamento"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   12
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empresa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   11
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Agregar"
         Height          =   555
         Left            =   4560
         Picture         =   "Frm_Aux_Listar_Empleados.frx":0080
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "A"
         Top             =   4440
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   6000
         Picture         =   "Frm_Aux_Listar_Empleados.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "C"
         Top             =   4440
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Cat_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Empleados 
            Height          =   2160
            Left            =   75
            TabIndex        =   2
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   3810
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
      Begin VB.Label Lbl_Programación_Cursos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "LISTA DE EMPLEADOS"
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
         Left            =   2100
         TabIndex        =   5
         Top             =   15
         Width           =   4185
      End
   End
End
Attribute VB_Name = "Frm_Aux_Listar_Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicializa()
Consulta_Cat_Empleados "", "", "", ""
 Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Filtro_Empleado_Empresa, 0, "Empresa_ID", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Filtro_Empleado_Departamento, 0, "Departamento_ID", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Puesto_ID, Nombre", "Cat_Puestos", Cmb_Filtro_Empleado_Puesto, 0, "Puesto_ID", "", False, "")
 Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Filtro_Empleado_Turno, 0, "Turno_ID", "", False, "")

End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Empleados
    'DESCRIPCIÓN:           Consulta los empleados y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre del curso
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Cat_Empleados(Empresa As String, Departamento As String, Puesto As String, Turno As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros

    Grid_Cat_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Cat_Empleados.* "
    Mi_SQL = Mi_SQL & "FROM Cat_Empleados, Cat_Empresas, Cat_Departamentos, Cat_Puestos, Cat_Turnos "
    Mi_SQL = Mi_SQL & "Where Cat_Empleados.Empresa_ID = Cat_Empresas.Empresa_ID "
    Mi_SQL = Mi_SQL & "AND Cat_Empleados.Departamento_ID = Cat_Departamentos.Departamento_ID "
    Mi_SQL = Mi_SQL & "AND Cat_Empleados.Puesto_ID = Cat_Puestos.Puesto_ID "
    Mi_SQL = Mi_SQL & "AND Cat_Empleados.Turno_ID = Cat_Turnos.Turno_ID "
    Mi_SQL = Mi_SQL & "AND Cat_Empresas.Nombre LIKE '%" & Empresa & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Departamentos.Nombre LIKE '%" & Departamento & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Puestos.Nombre LIKE '%" & Puesto & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Turnos.Nombre LIKE '%" & Turno & "%'"
    Mi_SQL = Mi_SQL & "AND Cat_Empleados.Estatus = 'A' "

    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Cat_Empleados
        If Not .EOF Then

            Grid_Cat_Empleados.AddItem "Empleado_Id" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno"
            While Not .EOF
                Grid_Cat_Empleados.AddItem .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Empleados.FixedRows = 1
            Grid_Cat_Empleados.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Empleados.ColWidth(1) = 2000   'Nombre
            Grid_Cat_Empleados.ColWidth(2) = 2000   'Clave
           Grid_Cat_Empleados.ColWidth(3) = 2000
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Empleados = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Empresa As String 'Obtiene el nombre a consultar
Dim Departamento As String 'Obtiene el nombre a consultar
Dim Puesto As String 'Obtiene el nombre a consultar
Dim Turno As String 'Obtiene el nombre a consultar
If Cmb_Filtro_Empleado_Empresa.ListIndex > -1 Then
Empresa = Cmb_Filtro_Empleado_Empresa.Text
End If
If Cmb_Filtro_Empleado_Departamento.ListIndex > -1 Then
Departamento = Cmb_Filtro_Empleado_Departamento.Text
End If
If Cmb_Filtro_Empleado_Puesto.ListIndex > -1 Then
Puesto = Cmb_Filtro_Empleado_Puesto.Text
End If
If Cmb_Filtro_Empleado_Turno.ListIndex > -1 Then
Turno = Cmb_Filtro_Empleado_Turno.Text
End If
        Consulta_Cat_Empleados Empresa, Departamento, Puesto, Turno
End Sub

Private Sub Btn_Nuevo_Click()
'Dim Rs_Consulta_Cat_Tipos_Notas_Credito As rdoResultset

    If Grid_Cat_Empleados.Rows > 1 Then
    For I = 0 To Grid_Cat_Empleados.Rows - 1

    Frm_Ope_Programacion_Cursos.Empleado_Seleccion_Id = Grid_Cat_Empleados.TextMatrix(I, 0)
    Frm_Ope_Programacion_Cursos.Llenar_Grid_Invitados
    Next I
    
    
'       Frm_Ope_Programacion_Cursos.Empleado_Seleccion_Id = Grid_Cat_Empleados.TextMatrix(Grid_Cat_Empleados.RowSel, 0)
        Unload Me


    Else
            MsgBox ("No existen empleados con estos filtros")
    End If
End Sub



