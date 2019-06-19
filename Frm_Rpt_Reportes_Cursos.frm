VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Rpt_Reportes_Cursos 
   Caption         =   "Reportes Cursos"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   6000
   Begin VB.Frame Fra_Cursos_Indices_Asistencia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indices de Asistencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   5850
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   405
         Width           =   825
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   660
         Left            =   4440
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reporte"
         Height          =   660
         Left            =   3120
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "C"
         Top             =   2880
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   870
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
   End
   Begin VB.Frame Fra_Cursos_Horas_Hombre 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte Cursos Horas/Hombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
      Begin VB.CommandButton Btn_Rpt_Generar 
         Caption         =   "Reporte"
         Height          =   660
         Left            =   3120
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "C"
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Salir_Reporte 
         Caption         =   "Salir"
         Height          =   660
         Left            =   4440
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   870
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Rango de fechas"
         Height          =   735
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Fra_Cursos_Reporte_General 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas"
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   1125
         Width           =   825
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
         Height          =   660
         Left            =   4440
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reporte"
         Height          =   660
         Left            =   3120
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "C"
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000E&
         Caption         =   "Auditable"
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frm_Rpt_Reportes_Cursos.frx":213C
         Left            =   1200
         List            =   "Frm_Rpt_Reportes_Cursos.frx":2146
         TabIndex        =   21
         Top             =   480
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   1590
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   1080
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo Curso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.Frame Fra_Cursos_Resumen_Mensual 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resumen Mensual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
      Begin VB.ComboBox Cmb_Cat_Cursos_Tipo_Curso 
         Height          =   315
         ItemData        =   "Frm_Rpt_Reportes_Cursos.frx":215C
         Left            =   1200
         List            =   "Frm_Rpt_Reportes_Cursos.frx":2166
         TabIndex        =   18
         Top             =   480
         Width           =   2745
      End
      Begin VB.CheckBox Chk_Auditable 
         BackColor       =   &H8000000E&
         Caption         =   "Auditable"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Reporte"
         Height          =   660
         Left            =   3120
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":217C
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "C"
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   660
         Left            =   4440
         Picture         =   "Frm_Rpt_Reportes_Cursos.frx":2706
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1125
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   1590
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   1080
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   103022595
         CurrentDate     =   41039
      End
      Begin VB.Label Lbl_Tipo_Curso_id 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo Curso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   765
      End
   End
End
Attribute VB_Name = "Frm_Rpt_Reportes_Cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Reporte As String                   'Nombre del reporte a manejar en la forma

Public Sub Cargar_Frame(Frame As Frame, Formulario As Form)
Dim Control As Control          'Toma la forma del objeto al que esta apuntando en ese momento

    'Oculta los pictures contenidos en la forma
    For Each Control In Formulario.Controls
        If TypeOf Control Is Frame Then
            Control.Visible = False
        End If
    Next
    Frame.Visible = True
    Frame.Top = 0
    Frame.Left = 0
    Formulario.Width = Frame.Width + 200
    Formulario.Height = Frame.Height + 400
    Formulario.Left = (Screen.Width - Formulario.Width) \ 2
    Formulario.Top = (Screen.Height - Formulario.Height) \ 2
End Sub

Public Sub Inicializar()
Dim Rs_Empleados_Departamento As rdoResultset
Dim Rs_Empleados_Supervisor As rdoResultset

'    Btn_Imprimir.Enabled = False
'    Btn_Exportar.Enabled = False
'    Btn_Regresar.Enabled = False
'    Btn_Salir.Enabled = False
'
    Select Case Reporte
        Case "Cursos_Horas_Hombre":
'            'Carga las empresas
'            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa, 0, "Nombre", , True, "TODAS")
'            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListCount > 0 Then
'                Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex = 0
'            End If
'            'Departamentos
''            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento, 0, "Nombre", , True, "TODOS")
'
'            'Consulta Departamento.
'            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
'            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
'            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
'            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'            Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.Clear
'            While Not Rs_Empleados_Departamento.EOF
'                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
'                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
'                Rs_Empleados_Departamento.MoveNext
'            Wend
'            Rs_Empleados_Departamento.Close
'            Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.Text = "<-SELECCIONE->"
'
'            If Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListCount > 0 Then
'                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex = 0
'            End If
'            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
'                Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
'            End If
'            Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.ListIndex = 0
'            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value = Now
'            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value = Now
'            Btn_Rpt_Generar.Top = 3000
'            Btn_Salir_Reporte.Top = 3000
'            Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.SetFocus
'
            
        Case "Cursos_Por_Empleado"
'            Call Conectar_Ayudante.Llena_Combo_Con_Items("Empleado_ID, Nombre, Apellido_Paterno, Apellido_Materno", "Cat_Empleados", Cmb_Empleado_CurTomEmp, 0, "Nombre")
'             Call Conectar_Ayudante.Llena_Combo_Item("Empleado_Id, (Nombre+' ' + Apellido_Paterno+ ' ' + Apellido_Materno) as Nombre", "Cat_Empleados", Cmb_Empleado_CurTomEmp, 0, "Empleado_Id", "", False, "")
    End Select
End Sub
