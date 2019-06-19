VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Ope_Asistencia_Cursos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASISTENCIA A CURSOS"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   7455
   Begin VB.PictureBox Pic_Ope_Programacion_Cursos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   6000
         Picture         =   "Frm_Ope_Asistencia_Cursos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6960
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar Cursos"
         Height          =   555
         Left            =   4560
         Picture         =   "Frm_Ope_Asistencia_Cursos.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "C"
         Top             =   6960
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Generales_Cat_Empleados 
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
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   7200
         Begin VB.CommandButton Btn_Agregar_Invitado 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   5640
            TabIndex        =   31
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox Txt_Empleado_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.ComboBox Cmb_Cat_Empleados 
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":0204
            Left            =   840
            List            =   "Frm_Ope_Asistencia_Cursos.frx":020E
            TabIndex        =   24
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empleado ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   810
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nombre"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Fra_Grid_Ope_Programacion_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Asistencias Cursos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Ope_Programacion_Cursos 
            Height          =   3585
            Left            =   75
            TabIndex        =   29
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   6324
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
      Begin VB.Frame Fra_Generales_Ope_Programacion_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7200
         Begin VB.CheckBox Chk_Busqueda_Fechas 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   22
            Top             =   2040
            Width           =   375
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Estatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":0224
            Left            =   4440
            List            =   "Frm_Ope_Asistencia_Cursos.frx":022E
            TabIndex        =   7
            Top             =   960
            Width           =   2490
         End
         Begin VB.TextBox Txt_Ope_Programacion_Cusos_No_Programa_Curso 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Curso 
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":0244
            Left            =   1125
            List            =   "Frm_Ope_Asistencia_Cursos.frx":024E
            TabIndex        =   5
            Top             =   600
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Sala 
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":0264
            Left            =   1125
            List            =   "Frm_Ope_Asistencia_Cursos.frx":026E
            TabIndex        =   4
            Top             =   960
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Institucion 
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":0284
            Left            =   1125
            List            =   "Frm_Ope_Asistencia_Cursos.frx":028E
            TabIndex        =   3
            Top             =   1320
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Ope_Prog_Cursos_Instructor 
            Height          =   315
            ItemData        =   "Frm_Ope_Asistencia_Cursos.frx":02A4
            Left            =   1125
            List            =   "Frm_Ope_Asistencia_Cursos.frx":02AE
            TabIndex        =   2
            Top             =   1680
            Width           =   5850
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Fecha_Fin 
            Height          =   315
            Left            =   4560
            TabIndex        =   8
            Top             =   2040
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Format          =   121962497
            CurrentDate     =   42373
         End
         Begin MSComCtl2.DTPicker Dt_Ope_Prog_Cursos_Fecha_Inicio 
            Height          =   315
            Left            =   1605
            TabIndex        =   9
            Top             =   2040
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   393216
            Format          =   121962497
            CurrentDate     =   42373
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estatus"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3600
            TabIndex        =   17
            Top             =   1050
            Width           =   525
         End
         Begin VB.Label Lbl_Ciudad 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Instructor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   16
            Top             =   1800
            Width           =   660
         End
         Begin VB.Label Lbl_Direccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Institucion"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label Lbl_Estado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha Inicio"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   495
            TabIndex        =   14
            Top             =   2130
            Width           =   870
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "No Curso"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   13
            Top             =   330
            Width           =   660
         End
         Begin VB.Label Lbl_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Sala"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   12
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Curso"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   11
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha Fin"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3600
            TabIndex        =   10
            Top             =   2130
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Grid_Cat_Empleados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lista Empleados"
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
         ForeColor       =   &H80000008&
         Height          =   5760
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   7185
         Begin VB.CommandButton Btn_Quitar_Empleado 
            Caption         =   "Quitar"
            Height          =   255
            Left            =   5640
            TabIndex        =   32
            Top             =   5400
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Empleados 
            Height          =   5040
            Left            =   75
            TabIndex        =   30
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   8890
            _Version        =   393216
            Rows            =   0
            Cols            =   4
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
         Caption         =   "REGISTRO DE ASISTENCIA"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   15
         Width           =   5025
      End
   End
End
Attribute VB_Name = "Frm_Ope_Asistencia_Cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Empleado_Seleccion_Id As String
Dim No_Curso_Seleccionado As String

Option Explicit
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim Crear_Features As DPFPFeatureExtraction
Dim Verificacion As DPFPVerification
Dim Convertir_Sample As DPFPSampleConversion
Dim Template_BD As DPFPTemplate
Dim Segundos_Espera As Integer


Public Sub Inicializa()
    Consulta_Ope_Programacion_Cursos "", "", "", "", "19100101", "21001212"
    Empleado_Seleccion_Id = ""
    Call Conectar_Ayudante.Llena_Combo_Item("Curso_Id, Nombre", "Cat_Cursos_Capacitaciones", Cmb_Ope_Prog_Cursos_Curso, 0, "Curso_Id", "", False, "")
    Call Conectar_Ayudante.Llena_Combo_Item("Sala_Id, Nombre", "Cat_Salas", Cmb_Ope_Prog_Cursos_Sala, 0, "Sala_Id", "", False, "")
    Call Conectar_Ayudante.Llena_Combo_Item("Institucion_Id, Nombre", "Cat_Instituciones", Cmb_Ope_Prog_Cursos_Institucion, 0, "Institucion_Id", "", False, "")
    Call Conectar_Ayudante.Llena_Combo_Item("Instructor_Id, Nombre", "Cat_Instructores", Cmb_Ope_Prog_Cursos_Instructor, 0, "Instructor_Id", "", False, "")
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
Private Sub Consulta_Ope_Programacion_Cursos(Curso As String, Sala As String, Institucion As String, Instructor As String, Fecha_Inicio As String, Fecha_Fin As String)
Dim Rs_Consulta_Ope_Programacion_Cursos As rdoResultset       'Informacion de los registros

    Grid_Ope_Programacion_Cursos.Rows = 0
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Ope_Programacion_Cursos.*, Cat_Cursos_Capacitaciones.Nombre as Curso, Cat_Instituciones.Nombre as Institucion, (Cat_Instructores.Nombre + ' ' + Cat_Instructores.Apellido_Paterno + ' ' + Cat_Instructores.Apellido_Materno) as Instructor "
    Mi_SQL = Mi_SQL & "FROM Ope_Programacion_Cursos, Cat_Cursos_Capacitaciones, Cat_Salas, Cat_Instituciones, Cat_Instructores "
    Mi_SQL = Mi_SQL & "WHERE Cat_Cursos_Capacitaciones.Curso_Id = Ope_Programacion_Cursos.Curso_Id "
    Mi_SQL = Mi_SQL & "AND Cat_Salas.Sala_Id = Ope_Programacion_Cursos.Sala_Id "
    Mi_SQL = Mi_SQL & "AND Cat_Instituciones.Institucion_Id = Ope_Programacion_Cursos.Institucion_Id "
    Mi_SQL = Mi_SQL & "AND Cat_Instructores.Instructor_Id = Ope_Programacion_Cursos.Instructor_Id "
    Mi_SQL = Mi_SQL & "AND Cat_Cursos_Capacitaciones.Nombre LIKE '%" & Curso & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Salas.Nombre LIKE '%" & Sala & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Instituciones.Nombre LIKE '%" & Institucion & "%' "
    Mi_SQL = Mi_SQL & "AND Cat_Instructores.Nombre LIKE '%" & Instructor & "%' "
    Mi_SQL = Mi_SQL & "AND Fecha_Inicio >= '" & Fecha_Inicio & "' "
    Mi_SQL = Mi_SQL & "AND Fecha_Fin <= '" & Fecha_Fin & "' "
    Mi_SQL = Mi_SQL & "AND Ope_Programacion_Cursos.Estatus = 'ACTIVO' "
    Mi_SQL = Mi_SQL & "AND Cat_Cursos_Capacitaciones.Estatus = 'ACTIVO' "
    Set Rs_Consulta_Ope_Programacion_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Ope_Programacion_Cursos
        If Not .EOF Then
            Grid_Ope_Programacion_Cursos.AddItem "No Curso" & Chr(9) & "Nombre" & Chr(9) & "Institucion" & Chr(9) & "Instructor"
            While Not .EOF
                Grid_Ope_Programacion_Cursos.AddItem .rdoColumns("No_Programa_Curso") & Chr(9) & .rdoColumns("Curso") & Chr(9) & .rdoColumns("Institucion") & Chr(9) & .rdoColumns("Instructor")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Ope_Programacion_Cursos.FixedRows = 1
            Grid_Ope_Programacion_Cursos.ColWidth(0) = 800     'Intitución_ID
            Grid_Ope_Programacion_Cursos.ColWidth(1) = 4000   'Nombre
            Grid_Ope_Programacion_Cursos.ColWidth(2) = 4000   'Clave
           Grid_Ope_Programacion_Cursos.ColWidth(3) = 4000
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Programacion_Cursos = Nothing

End Sub


Private Sub Btn_Agregar_Invitado_Click()
If Txt_Empleado_Id.Text <> "" Then
    If (Verificar_Id_Empleado) Then
        Alta_Ope_Lista_Asistencia
        Llenar_Grid_Lista_Aistencia
    End If
Else
    MsgBox ("Seleccione un empleado para agregar")
End If
End Sub

Private Sub Btn_Buscar_Click()
If Btn_Buscar.Caption = "Buscar Cursos" Then
Dim Curso As String 'Obtiene el nombre a consultar
Dim Sala As String 'Obtiene el nombre a consultar
Dim Institucion As String 'Obtiene el nombre a consultar
Dim Instructor As String 'Obtiene el nombre a consultar
Dim Fecha_Inicio As String
Dim Fecha_Fin As String
Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = ""

If Cmb_Ope_Prog_Cursos_Curso.ListIndex > -1 Then
Curso = Cmb_Ope_Prog_Cursos_Curso.Text
End If
If Cmb_Ope_Prog_Cursos_Sala.ListIndex > -1 Then
Sala = Cmb_Ope_Prog_Cursos_Sala.Text
End If
If Cmb_Ope_Prog_Cursos_Institucion.ListIndex > -1 Then
Institucion = Cmb_Ope_Prog_Cursos_Institucion.Text
End If
If Cmb_Ope_Prog_Cursos_Instructor.ListIndex > -1 Then
Instructor = Cmb_Ope_Prog_Cursos_Instructor.Text
End If
If Chk_Busqueda_Fechas.Value = 1 Or Chk_Busqueda_Fechas.Value = True Then
Fecha_Inicio = Dt_Ope_Prog_Cursos_Fecha_Inicio.Value
Fecha_Fin = Dt_Ope_Prog_Cursos_Fecha_Fin.Value
Else
Fecha_Inicio = ""
Fecha_Fin = ""
End If

Consulta_Ope_Programacion_Cursos Curso, Sala, Institucion, Instructor, Fecha_Inicio, Fecha_Fin
          Else
          Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o apellidos para buscar los empleados")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Dim Condicion As String
        Condicion = "((Nombre Like '%" & Nombre & "%' or Apellido_Paterno like '%" & Nombre & "%' or Apellido_Materno like '%" & Nombre & "%')"
       Condicion = Condicion & " or (Nombre +' '+Apellido_Paterno+' ' + Apellido_Materno) LIKE '%" & Nombre & "%') AND Estatus = 'A'"
       Llena_Combo "Empleado_Id, Nombre+' '+Apellido_Paterno+' '+Apellido_Materno as Nombre", "Cat_Empleados", Cmb_Cat_Empleados, 1, Condicion, "", False, ""
          End If
          
End Sub

Private Sub Btn_Quitar_Empleado_Click()
 If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Dim No_Lista_Seleccion As String
        If (Eliminar_No_Lista_Asistencia(Grid_Cat_Empleados.TextMatrix(Grid_Cat_Empleados.RowSel, 0))) Then
            MsgBox ("El empleado ha sido eliminado del curso")
            Llenar_Grid_Lista_Aistencia
        End If
        'Grid_Cat_Empleados.RemoveItem (Grid_Ope_Programacion_Invitacion_Empleados.RowSel)
    End If
End Sub

Private Sub Btn_Salir_Click()
If Btn_Salir.Caption = "Salir" Then
Unload Me
Else
Configuracion_Empleados False
Consulta_Ope_Programacion_Cursos "", "", "", "", "19100101", "21001212"

Capture.StopCapture
End If
End Sub

Private Sub Cmb_Cat_Empleados_Click()
Dim Indice As Integer
Indice = Cmb_Cat_Empleados.ListIndex
If Indice > -1 Then
Txt_Empleado_Id.Text = Format(Cmb_Cat_Empleados.ItemData(Cmb_Cat_Empleados.ListIndex), "00000")
End If

End Sub

Private Sub Grid_Cat_Empleados_Click()
If Grid_Cat_Empleados.ColSel = 4 Then
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Dim No_Lista_Seleccion As String
        If (Eliminar_No_Lista_Asistencia(Grid_Cat_Empleados.TextMatrix(Grid_Cat_Empleados.RowSel, 0))) Then
            MsgBox ("El empleado ha sido eliminado del curso")
            Llenar_Grid_Lista_Aistencia
        End If
        'Grid_Cat_Empleados.RemoveItem (Grid_Ope_Programacion_Invitacion_Empleados.RowSel)
    End If
End If
End Sub

Private Sub Grid_Ope_Programacion_Cursos_Click()
Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = Grid_Ope_Programacion_Cursos.TextMatrix(Grid_Ope_Programacion_Cursos.RowSel, 0)
End Sub

'
''*******************************************************************************
'    'NOMBRE DE LA FUNCIÓN:  Grid_Ope_Programacion_Cursos_DblClick
'    'DESCRIPCIÓN:           Carga el frm para agregar los empleados que asistieron al curso
'    'PARÁMETROS :
'    'CREO       :           Ana Laura Huichapa Ramírez
'    'FECHA_CREO :           11 Enero 2016
'    'MODIFICO          :
'    'FECHA_MODIFICO    :
'    'CAUSA_MODIFICACIÓN:
''*******************************************************************************

Private Sub Grid_Ope_Programacion_Cursos_DblClick()
Configuracion_Empleados (True)
' Call Conectar_Ayudante.Llena_Combo_Item("Empleado_Id, (Nombre+' ' + Apellido_Paterno+ ' ' + Apellido_Materno) as Nombre", "Cat_Empleados", Cmb_Cat_Empleados, 2, "Nombre", "", False, "")
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = ""
'        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Dim Condicion As String
        Condicion = "((Nombre Like '%" & Nombre & "%' or Apellido_Paterno like '%" & Nombre & "%' or Apellido_Materno like '%" & Nombre & "%')"
       Condicion = Condicion & " or (Nombre +' '+Apellido_Paterno+' ' + Apellido_Materno) LIKE '%" & Nombre & "%') AND Estatus = 'A'"
        Llena_Combo "Empleado_Id, Nombre+' '+Apellido_Paterno+' '+Apellido_Materno as Nombre", "Cat_Empleados", Cmb_Cat_Empleados, 1, Condicion, "", False, ""
       

No_Curso_Seleccionado = Grid_Ope_Programacion_Cursos.TextMatrix(Grid_Ope_Programacion_Cursos.RowSel, 0)
Fra_Grid_Cat_Empleados.Caption = "Curso: " & Grid_Ope_Programacion_Cursos.TextMatrix(Grid_Ope_Programacion_Cursos.RowSel, 1)
Llenar_Grid_Lista_Aistencia
    Registrar_Huella
End Sub
''*******************************************************************************
'    'NOMBRE DE LA FUNCIÓN:  Configuracion_Empleados
'    'DESCRIPCIÓN:           Configura el formulario para mostrar los controles necesarios
'    'PARÁMETROS :
'    'CREO       :           Ana Laura Huichapa Ramírez
'    'FECHA_CREO :           11 Enero 2016
'    'MODIFICO          :
'    'FECHA_MODIFICO    :
'    'CAUSA_MODIFICACIÓN:
''*******************************************************************************

Private Sub Configuracion_Empleados(Estatus As Boolean)
Fra_Generales_Cat_Empleados.Visible = Estatus
Fra_Generales_Ope_Programacion_Cursos.Visible = Not Estatus
Fra_Grid_Cat_Empleados.Visible = Estatus
Fra_Grid_Ope_Programacion_Cursos.Visible = Not Estatus
Fra_Grid_Cat_Empleados.Enabled = Estatus
Fra_Generales_Cat_Empleados.Enabled = Estatus
If Estatus Then
Btn_Buscar.Caption = "Buscar Empleados"
Btn_Salir.Caption = "Regresar"
Else
Btn_Buscar.Caption = "Buscar Cursos"
Btn_Salir.Caption = "Salir"
Grid_Ope_Programacion_Cursos.Visible = True
'Consulta_Ope_Programacion_Cursos "", "", "", "", ""
End If


End Sub

Private Sub Llenar_Grid_Lista_Aistencia()
Dim Dia As Integer
Dim Mes As Integer
Dim año As Integer
Dia = Day(Now)
Mes = Month(Now)
año = Year(Now)
Dim Rs_Consulta_Lista_Asistentes As rdoResultset       'Informacion de los registros

    Grid_Cat_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Ope_Lista_Asistencia.*,  Nombre, Apellido_Paterno, Apellido_Materno "
    Mi_SQL = Mi_SQL & " FROM  Ope_Lista_Asistencia, Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE No_Programa_Curso = '" & No_Curso_Seleccionado & "' "
    Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Empleado_Id = Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " and DATEPART (DAY, Ope_Lista_Asistencia.Fecha_Hora_Registro) = " & Dia
    Mi_SQL = Mi_SQL & " and DATEPART (MONTH, Ope_Lista_Asistencia.Fecha_Hora_Registro) = " & Mes
    Mi_SQL = Mi_SQL & " and DATEPART (YEAR, Ope_Lista_Asistencia.Fecha_Hora_Registro) = " & año
'    MsgBox Mi_SQL

    Set Rs_Consulta_Lista_Asistentes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Lista_Asistentes
        If Not .EOF Then
            If Grid_Cat_Empleados.Rows <= 0 Then
            Grid_Cat_Empleados.AddItem "Empleado Id" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno"
           End If
           While Not .EOF
                Grid_Cat_Empleados.AddItem .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Empleados.FixedRows = 1
            Grid_Cat_Empleados.FixedCols = 1
            Grid_Cat_Empleados.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Empleados.ColWidth(1) = 2000   'Nombre
            Grid_Cat_Empleados.ColWidth(2) = 2000   'Clave
            Grid_Cat_Empleados.ColWidth(3) = 2000
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Lista_Asistentes = Nothing

End Sub



Function Verificar_Id_Empleado() As Boolean
Verificar_Id_Empleado = True
Dim I As Integer
Dim Id_Emplado_Grid As String
For I = 0 To Grid_Cat_Empleados.Rows - 1

Id_Emplado_Grid = Grid_Cat_Empleados.TextMatrix(I, 0)
If Txt_Empleado_Id.Text = Id_Emplado_Grid Then
Verificar_Id_Empleado = False
If Id_Emplado_Grid <> "Empleado_Id" Then
'MsgBox ("El empleado ya ha sido agregado anteriormente")
End If
'Load Frm_Aux_Listar_Empleados
'Frm_Aux_Listar_Empleados.Inicializa
Exit For
End If
Next I
End Function



'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Ope_Lista_Asistencia
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del curso que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 11 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Ope_Lista_Asistencia()
'Dim Menus As Integer
'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Ope_Lista_Asistencia As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans

    'Alta de Institución
    Set Rs_Alta_Ope_Lista_Asistencia = Conectar_Ayudante.Recordset_Agregar("Ope_Lista_Asistencia")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Ope_Lista_Asistencia
        .AddNew
'            Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Ope_Programacion_Cursos", "No_Programa_Curso"), "0000000000")
'            .rdoColumns("No_Lista_Asistencia") = Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text
            .rdoColumns("Empleado_Id") = Txt_Empleado_Id.Text
            .rdoColumns("Fecha_Hora_Registro") = Now
            .rdoColumns("No_Programa_Curso") = Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Ope_Lista_Asistencia.Close
    Conexion_Base.CommitTrans
'    MsgBox "Programacion de curso agregada", vbInformation
'    Consulta_Ope_Programacion_Cursos ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Function Eliminar_No_Lista_Asistencia(Empleado_Id As String) As Boolean
Dim No_Lista_Asistencia As String
Dim Rs_Consulta_Ope_Lista_Asistencia As rdoResultset       'Informacion de los registros

'    Grid_Ope_Programacion_Cursos.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Ope_Lista_Asistencia.* "
    Mi_SQL = Mi_SQL & "FROM Ope_Lista_Asistencia "
    Mi_SQL = Mi_SQL & "WHERE No_Programa_Curso = '" & Txt_Ope_Programacion_Cusos_No_Programa_Curso.Text & "' "
    Mi_SQL = Mi_SQL & "AND Empleado_Id ='" & Empleado_Id & "'"
    
   
    Set Rs_Consulta_Ope_Lista_Asistencia = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Ope_Lista_Asistencia
        If Not .EOF Then
            No_Lista_Asistencia = .rdoColumns("No_Lista_Asistencia")
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Lista_Asistencia = Nothing
    Dim Resultado As Boolean
Resultado = Conectar_Ayudante.Elimina_Catalogo("Ope_Lista_Asistencia", "No_Lista_Asistencia", No_Lista_Asistencia)
Eliminar_No_Lista_Asistencia = Resultado
End Function



''*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo_Item
    'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
    'PARÁMETROS:
    '             1. Campos: Campo a consultar para el llenado del ComboBox
    '             2. Tabla: Nombre de la tabla a consultar
    '             3. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             4. Tipo: Para saber si esta consultando
    '             5. Campo_Con: Para consultar y llenar el campo con las palabras introducidas
    '                           por el usuario
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Llena_Combo(Campos As String, Tabla As String, Combo_Control As ComboBox, Tipo As Integer, Condicion As String, Optional Condicion_Adicional As String = "", Optional Mostrar_Todos As Boolean, Optional Mensaje_Todos As String)
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & Condicion
        End If
       
'        .SQL = .SQL & " ORDER BY " & Campo_con
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    Combo_Control.Clear
    If Mostrar_Todos Then
        If Mensaje_Todos <> "" Then
            Combo_Control.AddItem Mensaje_Todos
        Else
            Combo_Control.AddItem "TODOS"
        End If
        Combo_Control.ItemData(Combo_Control.NewIndex) = 0
    End If
    If Not Rs_Combo.EOF Then
        While Not Rs_Combo.EOF
            Combo_Control.AddItem Rs_Combo(1)
            Combo_Control.ItemData(Combo_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Sub

Private Sub Registrar_Huella()
    Me.Top = 0
    Me.Left = 0
    'Inicializa las librerías del lector
    'Create capture operation.
    Set Capture = New DPFPCapture
    'Start capture operation.
    Capture.StartCapture
    'Create DPFPFeatureExtraction object.
    Set Crear_Features = New DPFPFeatureExtraction
    'Create DPFPVerification object.
    Set Verificacion = New DPFPVerification
    'Create DPFPSampleConversion object.
    Set Convertir_Sample = New DPFPSampleConversion
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
Dim Feedback As DPFPCaptureFeedbackEnum
Dim Resultado As DPFPVerificationResult
Dim Template_Consulta As Object
Dim Template_Imagen() As Byte
Dim Rs_Consulta_Huellas As rdoResultset
Dim Ruta_Almacenamiento As String

    'Process sample and create feature set for purpose of verification.
    Feedback = Crear_Features.CreateFeatureSet(Sample, DataPurposeVerification)
    'Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
        'Consulta los registros de huella digital
        Mi_SQL = "SELECT Empleado_ID,No_Tarjeta,Huella_Ruta,Huella_Digital"
        Mi_SQL = Mi_SQL & " FROM Cat_Empleados_Huellas"
        Set Rs_Consulta_Huellas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Huellas.EOF
            
'            If Not IsNull(Rs_Consulta_Huellas.rdoColumns("Huella_Digital")) Then
'                'Lectura desde la base de datos
                Template_Imagen = Rs_Consulta_Huellas.rdoColumns("Huella_Digital")
'            Else
                'Lectura desde un archivo físico
                'Ruta_Almacenamiento = App.Path & "\Huellas\" & Rs_Consulta_Huellas.rdoColumns("Huella_Ruta")
                'Read binary data from file.
                'Open Ruta_Almacenamiento For Binary As #1
                 '   ReDim Template_Imagen(LOF(1))
                 '   Get #1, , Template_Imagen()
                'Close #1
'            End If

            'Template can be empty, it must be created first.
            If Template_BD Is Nothing Then Set Template_BD = New DPFPTemplate
            'Import binary data to template.
            Template_BD.Deserialize Template_Imagen
            Set Template_Consulta = Template_BD
            'Compare feature set with template.
            Set Resultado = Verificacion.Verify(Crear_Features.FeatureSet, Template_Consulta)
            If Resultado.Verified = True Then
                Txt_Empleado_Id.Text = Rs_Consulta_Huellas.rdoColumns("Empleado_ID")
                If Verificar_Estatus_Empleado(Txt_Empleado_Id.Text) Then
                    If (Verificar_Id_Empleado) Then
                        Alta_Ope_Lista_Asistencia
                        Llenar_Grid_Lista_Aistencia
                    End If
                Else
                    MsgBox ("No es posible registrar asistencia, el empleado no se encuentra activo")
                End If
                    Rs_Consulta_Huellas.Close
                    Exit Sub
            End If
            Rs_Consulta_Huellas.MoveNext
        Wend
        Rs_Consulta_Huellas.Close
    Else
        MsgBox "La calidad de muestra del lector es pobre"
    End If
End Sub

Function Verificar_Estatus_Empleado(Empleado_Id) As Boolean
Dim Rs_Consulta_Ope_Programacion_Cursos As rdoResultset       'Informacion de los registros

    Grid_Ope_Programacion_Cursos.Rows = 0
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT * "
    Mi_SQL = Mi_SQL & "FROM Cat_Empleados "
    Mi_SQL = Mi_SQL & "WHERE Empleado_Id = '" & Empleado_Id & "'"
    Set Rs_Consulta_Ope_Programacion_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Ope_Programacion_Cursos
        If .rdoColumns("Estatus") = "A" Then
        Verificar_Estatus_Empleado = True
        Else
        Verificar_Estatus_Empleado = False
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Programacion_Cursos = Nothing



End Function
