VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Asistencias_Mantenimiento 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Adm_Asistencia_Mantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   4845
      ScaleWidth      =   7740
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      Begin VB.Frame Fra_Adm_Asistencias_General 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General"
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
         Height          =   3435
         Left            =   45
         TabIndex        =   24
         Top             =   450
         Width           =   7620
         Begin VB.ComboBox Cmb_Adm_Asistencias_Tipo_Permiso 
            Height          =   315
            ItemData        =   "Frm_Adm_Asistencias_Mantenimiento.frx":0000
            Left            =   1305
            List            =   "Frm_Adm_Asistencias_Mantenimiento.frx":0002
            TabIndex        =   42
            Top             =   1431
            Width           =   6090
         End
         Begin VB.ComboBox Cmb_Adm_Asistencias_Permiso 
            Height          =   315
            ItemData        =   "Frm_Adm_Asistencias_Mantenimiento.frx":0004
            Left            =   1290
            List            =   "Frm_Adm_Asistencias_Mantenimiento.frx":0029
            TabIndex        =   41
            Top             =   1929
            Width           =   6090
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Movimiento 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   4095
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Referencia_Siguiente 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   3780
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Referencia 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_No_tarjeta 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   405
            Width           =   1770
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Horas_Aprobada 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1305
            MaxLength       =   5
            TabIndex        =   11
            Top             =   2947
            Width           =   1770
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Retardo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5645
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2955
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_No_Asistencias 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Empleado_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5645
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Nombre 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   918
            Width           =   6090
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Asistencias_Hora_Salida 
            Height          =   315
            Left            =   5640
            TabIndex        =   10
            Top             =   2415
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   117637122
            CurrentDate     =   39986
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Asistencias_Fecha_Asistencia 
            Height          =   315
            Left            =   5640
            TabIndex        =   7
            Top             =   413
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   117637123
            CurrentDate     =   39986
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Asistencias_Hora_Entrada 
            Height          =   330
            Left            =   1305
            TabIndex        =   9
            Top             =   2400
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   117637122
            CurrentDate     =   39986
         End
         Begin VB.Label Lbl_Turno 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Turno"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1995
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Permiso"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1485
            Width           =   915
         End
         Begin VB.Label Lbl_No_Empleado 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Empleado"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   975
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   4380
            TabIndex        =   31
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hora Entrada"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   2475
            Width           =   945
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hora Salida"
            Height          =   195
            Left            =   4380
            TabIndex        =   29
            Top             =   2475
            Width           =   825
         End
         Begin VB.Label Lbl_Horas_Extra 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hrs Extra"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   3015
            Width           =   645
         End
         Begin VB.Label Lbl_Horas_Trabajadaas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hrs. Trabajadas"
            Height          =   195
            Left            =   4380
            TabIndex        =   27
            Top             =   3015
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No_Asistencia"
            Height          =   195
            Left            =   45
            TabIndex        =   26
            Top             =   3525
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado ID"
            Height          =   195
            Left            =   4680
            TabIndex        =   25
            Top             =   3525
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   645
         Left            =   45
         Picture         =   "Frm_Adm_Asistencias_Mantenimiento.frx":00EA
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "M"
         Top             =   3960
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   645
         Left            =   6465
         Picture         =   "Frm_Adm_Asistencias_Mantenimiento.frx":0674
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3960
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Consultar 
         Caption         =   "Consultar"
         Height          =   645
         Left            =   3240
         Picture         =   "Frm_Adm_Asistencias_Mantenimiento.frx":0BFE
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "C"
         Top             =   3960
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MANTENIMIENTO ASISTENCIAS"
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
         Left            =   975
         TabIndex        =   23
         Top             =   0
         Width           =   5925
      End
   End
   Begin VB.PictureBox Pic_Adm_Asistencias_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   4845
      ScaleWidth      =   7740
      TabIndex        =   34
      Top             =   0
      Width           =   7740
      Begin VB.Frame Fra_Adm_Inasistencias_Consulta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultas"
         Height          =   4785
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   7665
         Begin VB.Frame Fra_Inasistencias_Consulta_Resultados 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Resultados"
            Height          =   3390
            Left            =   45
            TabIndex        =   36
            Top             =   1350
            Width           =   7575
            Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Asistencias_Consulta_Resultados 
               Height          =   3120
               Left            =   90
               TabIndex        =   22
               Top             =   225
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   5503
               _Version        =   393216
               Rows            =   0
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               ScrollBars      =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               Appearance      =   0
            End
         End
         Begin VB.ComboBox Cmb_Adm_Asistencias_Consulta_Supervisor 
            Height          =   315
            Left            =   1215
            TabIndex        =   14
            Top             =   225
            Width           =   4920
         End
         Begin VB.CommandButton Btn_Regresar 
            Cancel          =   -1  'True
            Caption         =   "Regresar"
            Height          =   510
            Left            =   6420
            Picture         =   "Frm_Adm_Asistencias_Mantenimiento.frx":1188
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   810
            UseMaskColor    =   -1  'True
            Width           =   1200
         End
         Begin VB.CheckBox Chk_Adm_Asistencias_Consulta_Periodo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Periodo"
            Height          =   315
            Left            =   135
            TabIndex        =   17
            Top             =   945
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CheckBox Chk_Adm_Asistencias_Consulta_Empleado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            Height          =   315
            Left            =   135
            TabIndex        =   15
            Top             =   585
            Width           =   1050
         End
         Begin VB.ComboBox Cmb_Adm_Asistencias_Consulta_Empleado 
            Height          =   315
            Left            =   1215
            TabIndex        =   16
            Top             =   585
            Width           =   4920
         End
         Begin VB.CheckBox Chk_Adm_Asistencias_Consulta_Supervisor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            Height          =   315
            Left            =   135
            TabIndex        =   13
            Top             =   225
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   510
            Left            =   6420
            Picture         =   "Frm_Adm_Asistencias_Mantenimiento.frx":1712
            Style           =   1  'Graphical
            TabIndex        =   20
            Tag             =   "C"
            Top             =   225
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Asistencias_Consulta_Fecha_Inicio 
            Height          =   315
            Left            =   1215
            TabIndex        =   18
            Top             =   945
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   117637123
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Asistencias_Consulta_Fecha_Termino 
            Height          =   315
            Left            =   4410
            TabIndex        =   19
            Top             =   945
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   117637123
            CurrentDate     =   39940
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   3607
            TabIndex        =   37
            Top             =   990
            Width           =   135
         End
      End
   End
End
Attribute VB_Name = "Frm_Adm_Asistencias_Mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Movimiento_Empleado As String           'Indica el movimiento del empleado seleccionado
Dim Referencia As String                    'Indica la referencia con la que se asociara la asistencia
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean

Private Sub Btn_Buscar_Click()
    Consulta_Adm_Asistencias
End Sub

Private Sub Btn_Consultar_Click()
    Pic_Adm_Asistencia_Mantenimiento.ZOrder vbSendToBack
    Pic_Adm_Asistencia_Mantenimiento.Visible = False
    Pic_Adm_Asistencias_Consulta.Visible = True
    Pic_Adm_Asistencias_Consulta.ZOrder vbBringToFront
    Chk_Adm_Asistencias_Consulta_Periodo.Value = 1
    Grid_Adm_Asistencias_Consulta_Resultados.Rows = 0
End Sub


Public Sub Inicializa()
    Dtp_Adm_Asistencias_Fecha_Asistencia.Value = Now
    Dtp_Adm_Asistencias_Hora_Entrada.Value = "00:00:00"
    Dtp_Adm_Asistencias_Hora_Salida.Value = "00:00:00"
    Dtp_Adm_Asistencias_Consulta_Fecha_Inicio.Value = Now
    Dtp_Adm_Asistencias_Consulta_Fecha_Termino.Value = Now
    Referencia = ""
    Llena_Combo_Turnos_Incidencias
    'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo = 'S' AND Estatus = 'A'", Cmb_Adm_Asistencias_Consulta_Supervisor, 0, "Apellido_paterno", , False, "TODOS")
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
        If Trim(Txt_Adm_Asistencias_No_Asistencias.Text) <> "" Then
            Fra_Adm_Asistencias_General.Enabled = True
            Cmb_Adm_Asistencias_Tipo_Permiso.SetFocus
        Else
            MsgBox "Seleccione una asistencia para modificar", vbOKOnly + vbInformation, Me.Caption
            Exit Sub
        End If
        Btn_Modificar.Caption = "Actualizar"
        Btn_Consultar.Enabled = False
        Btn_Salir.Caption = "Regresar"
    Else
        If Cmb_Adm_Asistencias_Tipo_Permiso.ListIndex > -1 Then
            If Cmb_Adm_Asistencias_Permiso.ListIndex > -1 Then
                If Val(Txt_Adm_Asistencias_Horas_Aprobada.Text) >= 0 Then
                    Modifica_Adm_Asistencias
                Else
                    MsgBox "Ingrese el no. de horas, por favor", vbExclamation
                End If
            Else
                MsgBox "Debe seleccionar el turno", vbExclamation
            End If
        Else
            MsgBox "Seleccione el tipo de movimiento", vbExclamation
        End If
    End If
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Adm_Asistencias_Consulta.ZOrder vbSendToBack
    Pic_Adm_Asistencias_Consulta.Visible = False
    Pic_Adm_Asistencia_Mantenimiento.ZOrder vbBringToFront
    Pic_Adm_Asistencia_Mantenimiento.Visible = True
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Modificar.Enabled = True
        Btn_Consultar.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Salir.Caption = "Salir"
        Fra_Adm_Asistencias_General.Enabled = False
        Dtp_Adm_Asistencias_Fecha_Asistencia.Value = Now
        Dtp_Adm_Asistencias_Hora_Entrada.Value = "00:00:00"
        Dtp_Adm_Asistencias_Hora_Salida.Value = "00:00:00"
        Cmb_Adm_Asistencias_Tipo_Permiso.Text = ""
        Cmb_Adm_Asistencias_Permiso.Text = ""
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Asistencias_Mantenimiento", Me)
    End If
End Sub

Private Sub Chk_Adm_Asistencias_Consulta_Empleado_Click()
Dim Rs_Empleados_Supervisor As rdoResultset
    If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 Then
        Cmb_Adm_Asistencias_Consulta_Empleado.Clear
        Cmb_Adm_Asistencias_Consulta_Empleado.SetFocus
        If Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex < 0 And Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
            MsgBox "Seleccione primero al supervisor", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 And Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex > -1 Then
            
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " AND Supervisor_ID = '" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Adm_Asistencias_Consulta_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Adm_Asistencias_Consulta_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
'         Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex = 0
'        Cmb_Adm_Asistencias_Consulta_Supervisor.Text = ""
            
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Asistencias_Consulta_Empleado, 1, "Apellido_Paterno", "AND (Nombre like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%') AND Supervisor_ID = '" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'", False, "")
        End If
    Else
        Cmb_Adm_Asistencias_Consulta_Empleado.Clear
    End If
End Sub

Private Sub Chk_Adm_Asistencias_Consulta_Supervisor_Click()
    If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
        Cmb_Adm_Asistencias_Consulta_Supervisor.Clear
        Cmb_Adm_Asistencias_Consulta_Supervisor.SetFocus
    Else
        Cmb_Adm_Asistencias_Consulta_Supervisor.Clear
    End If

End Sub

Private Sub Chk_Adm_Asistencias_Consulta_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Supervisor As rdoResultset
    If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
        'Consulta Supervisor.
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Adm_Asistencias_Consulta_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Adm_Asistencias_Consulta_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
        Cmb_Adm_Asistencias_Consulta_Supervisor.Text = ""
        
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo = 'S' AND Estatus = 'A' AND Nombre like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%'", Cmb_Adm_Asistencias_Consulta_Supervisor, 0, "")
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Asistencias_Consulta_Empleado_KeyPress(KeyAscii As Integer)
Dim Rs_Consulta_Empleado As rdoResultset

    If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            If Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex > -1 Then
                'Consulta el filtro de los empleados del supervisor
                Mi_SQL = "SELECT Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre"
                Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'"
                
                If IsNumeric(Cmb_Adm_Asistencias_Consulta_Empleado.Text) Then
                    Mi_SQL = Mi_SQL & " AND  No_Tarjeta='" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "'"
                Else
                    Mi_SQL = Mi_SQL & " AND (Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Nombre LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%')"
                End If
            
            Else
            
                'Consulta el filtro de los empleados del supervisor
                Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre"
                Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
                Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
                
                If IsNumeric(Cmb_Adm_Asistencias_Consulta_Empleado.Text) Then
                    Mi_SQL = Mi_SQL & " AND  No_Tarjeta='" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "'"
                Else
                    Mi_SQL = Mi_SQL & " AND (Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Nombre LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%')"
                End If
            
            End If
            Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Adm_Asistencias_Consulta_Empleado.Clear
            While Not Rs_Consulta_Empleado.EOF
                Cmb_Adm_Asistencias_Consulta_Empleado.AddItem Rs_Consulta_Empleado.rdoColumns("Nombre")
                Cmb_Adm_Asistencias_Consulta_Empleado.ItemData(Cmb_Adm_Asistencias_Consulta_Empleado.NewIndex) = Rs_Consulta_Empleado.rdoColumns("Empleado_ID")
                Rs_Consulta_Empleado.MoveNext
            Wend
            Rs_Consulta_Empleado.Close
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Asistencias_Consulta_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 Then
        Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Asistencias_Consulta_Empleado, KeyCode)
    Else
        KeyCode = 0
    End If
End Sub

Private Sub Cmb_Adm_Asistencias_Consulta_Supervisor_Click()
Dim Rs_Consulta_Empleado As rdoResultset
    
    If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 Then
        'Consulta el filtro de los empleados del supervisor
        Mi_SQL = "SELECT Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre"
        Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'"
        If IsNumeric(Cmb_Adm_Asistencias_Consulta_Empleado.Text) Then
            Mi_SQL = Mi_SQL & " AND No_Tarjeta='" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "'"
        Else
            Mi_SQL = Mi_SQL & " AND (Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Nombre LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%')"
        End If
        Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Adm_Asistencias_Consulta_Empleado.Clear
        While Not Rs_Consulta_Empleado.EOF
            Cmb_Adm_Asistencias_Consulta_Empleado.AddItem Rs_Consulta_Empleado.rdoColumns("Nombre")
            Cmb_Adm_Asistencias_Consulta_Empleado.ItemData(Cmb_Adm_Asistencias_Consulta_Empleado.NewIndex) = Rs_Consulta_Empleado.rdoColumns("Empleado_ID")
            Rs_Consulta_Empleado.MoveNext
        Wend
        Rs_Consulta_Empleado.Close
    End If
End Sub

Private Sub Cmb_Adm_Asistencias_Consulta_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Consulta_Empleado As rdoResultset

    If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            'Consulta el filtro del supervisor
            Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
            Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
            Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
            Mi_SQL = Mi_SQL & " AND Tipo='S'"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            
            If IsNumeric(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) Then
                Mi_SQL = Mi_SQL & " AND No_Tarjeta='" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "'"
            Else
                Mi_SQL = Mi_SQL & " AND (Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%' OR Nombre LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Supervisor.Text) & "%')"
            End If
            Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Adm_Asistencias_Consulta_Supervisor.Clear
            While Not Rs_Consulta_Empleado.EOF
                Cmb_Adm_Asistencias_Consulta_Supervisor.AddItem Rs_Consulta_Empleado.rdoColumns("Nombre")
                Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.NewIndex) = Rs_Consulta_Empleado.rdoColumns("Empleado_ID")
                Rs_Consulta_Empleado.MoveNext
            Wend
            Rs_Consulta_Empleado.Close
            If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 And Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex > -1 Then
                'Consulta el filtro de los empleados del supervisor
                Mi_SQL = "SELECT Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre"
                Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'"
                If IsNumeric(Cmb_Adm_Asistencias_Consulta_Empleado.Text) Then
                    Mi_SQL = Mi_SQL & " AND No_Tarjeta='" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "'"
                Else
                    Mi_SQL = Mi_SQL & " AND (Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%' OR Nombre LIKE '%" & Trim(Cmb_Adm_Asistencias_Consulta_Empleado.Text) & "%')"
                End If
                Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Adm_Asistencias_Consulta_Empleado.Clear
                While Not Rs_Consulta_Empleado.EOF
                    Cmb_Adm_Asistencias_Consulta_Empleado.AddItem Rs_Consulta_Empleado.rdoColumns("Nombre")
                    Cmb_Adm_Asistencias_Consulta_Empleado.ItemData(Cmb_Adm_Asistencias_Consulta_Empleado.NewIndex) = Rs_Consulta_Empleado.rdoColumns("Empleado_ID")
                    Rs_Consulta_Empleado.MoveNext
                Wend
                Rs_Consulta_Empleado.Close
            End If
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Asistencias_Consulta_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
        Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Asistencias_Consulta_Supervisor, KeyCode)
    Else
        KeyCode = 0
    End If
End Sub

Private Sub Grid_Adm_Asistencias_Consulta_Resultados_DblClick()
    If Grid_Adm_Asistencias_Consulta_Resultados.Rows > 0 Then
        Consulta_Informacion_Adm_Asistencias
    End If
End Sub

Private Sub Txt_Adm_Asistencias_Horas_Aprobada_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Adm_Asistencias_Horas_Aprobada, True)
End Sub

'************************************************Inicio Asistencias***************************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Adm_Asistencias
    'DESCRIPCIÓN:           Consulta la informacion de inasistencias
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           18 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Adm_Asistencias()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios
Dim Asistencia As String

On Error GoTo handler:
    Grid_Adm_Asistencias_Consulta_Resultados.Rows = 0
    Grid_Adm_Asistencias_Consulta_Resultados.Cols = 6
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT AA.No_Asistencia,AA.Empleado_ID,AA.No_Tarjeta,"
    Mi_SQL = Mi_SQL & " AA.Fecha,AA.Hora_Entrada,AA.Hora_Salida, AA.Simbologia,AA.SubSimbologia, AA.Tiempo_Retardo,"
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA, Cat_Empleados CE"
    Mi_SQL = Mi_SQL & " WHERE AA.Empleado_Id = CE.Empleado_ID"
    'Manejo de Filtros
    'Supervisor
    If Chk_Adm_Asistencias_Consulta_Supervisor.Value = 1 Then
        If Cmb_Adm_Asistencias_Consulta_Supervisor.Text <> "" Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Adm_Asistencias_Consulta_Supervisor.ItemData(Cmb_Adm_Asistencias_Consulta_Supervisor.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado el supervisor", vbInformation + vbOKOnly, Me.Caption
            Cmb_Adm_Asistencias_Consulta_Supervisor.SetFocus
            Exit Sub
        End If
    End If
    'Empleados
    If Chk_Adm_Asistencias_Consulta_Empleado.Value = 1 Then
        If Cmb_Adm_Asistencias_Consulta_Empleado.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AA.Empleado_ID = '" & Format(Cmb_Adm_Asistencias_Consulta_Empleado.ItemData(Cmb_Adm_Asistencias_Consulta_Empleado.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun empleado", vbInformation + vbOKOnly, Me.Caption
            Cmb_Adm_Asistencias_Consulta_Empleado.SetFocus
            Exit Sub
        End If
    End If
    'Periodo
    If Chk_Adm_Asistencias_Consulta_Periodo.Value = 1 Then
        If DateDiff("d", Format(Dtp_Adm_Asistencias_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Dtp_Adm_Asistencias_Consulta_Fecha_Termino, "MM/dd/yyyy")) < 0 Then
            MsgBox "Rango de Fechas Incorrecto", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        Else
            Mi_SQL = Mi_SQL & " AND AA.Fecha >= " & Par_Fecha & Format(Dtp_Adm_Asistencias_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND AA.Fecha <= " & Par_Fecha & Format(Dtp_Adm_Asistencias_Consulta_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
        End If
    End If
    Mi_SQL = Mi_SQL & " ORDER BY CE.Apellido_Paterno, AA.Fecha"
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    With Rs_Consulta_Adm_Asistencias
    If Not .EOF Then
        'Coloca un encabezado en el grid
        Grid_Adm_Asistencias_Consulta_Resultados.AddItem "No Movimiento" & Chr(9) & "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "Tipo" & Chr(9) & "SubTipo"
        While Not .EOF
            'Valores para las incindecias
            Grid_Adm_Asistencias_Consulta_Resultados.AddItem .rdoColumns("No_Asistencia") _
                & Chr(9) & .rdoColumns("No_Tarjeta") _
                & Chr(9) & .rdoColumns("Nombre") _
                & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                & Chr(9) & .rdoColumns("Simbologia") _
                & Chr(9) & .rdoColumns("SubSimbologia")
            .MoveNext
        Wend
        .Close
        'Configura el tamaño de las columnas del grid_usuarios
        With Grid_Adm_Asistencias_Consulta_Resultados
            .FixedRows = 1
            .ColWidth(0) = 0    'No Asistencia
            .ColWidth(1) = 1000 'No tarjeta
            .ColWidth(2) = 3000 'Nombre
            .ColWidth(3) = 1200 'Fecha
            .ColWidth(4) = 700 'Tipo
            .ColWidth(5) = 700 'SubTipo
        End With
    End If
    End With
Set Rs_Consulta_Adm_Asistencias = Nothing
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Informacion_Adm_Inasistencias
'DESCRIPCION: Consulta la información de la assitencia generada
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 03-Septiembre-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Informacion_Adm_Asistencias()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios

    Referencia = -1
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT * FROM Adm_Asistencias"
    Mi_SQL = Mi_SQL & " WHERE No_Asistencia='" & Grid_Adm_Asistencias_Consulta_Resultados.TextMatrix(Grid_Adm_Asistencias_Consulta_Resultados.RowSel, 0) & "'"
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    Movimiento_Empleado = ""
    If Not Rs_Consulta_Adm_Asistencias.EOF Then
        With Rs_Consulta_Adm_Asistencias
            Txt_Adm_Asistencias_No_Asistencias.Text = .rdoColumns("No_Asistencia")
            Txt_Adm_Asistencias_Empleado_ID.Text = .rdoColumns("Empleado_ID")
            Txt_Adm_Asistencias_No_tarjeta.Text = .rdoColumns("No_Tarjeta")
            Txt_Adm_Asistencias_Nombre.Text = Grid_Adm_Asistencias_Consulta_Resultados.TextMatrix(Grid_Adm_Asistencias_Consulta_Resultados.RowSel, 2)
            Dtp_Adm_Asistencias_Fecha_Asistencia.Value = .rdoColumns("Fecha")
            Call Conectar_Ayudante.Asigna_Item_Combo(Trim(.rdoColumns("Simbologia")), Cmb_Adm_Asistencias_Tipo_Permiso)
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Turno_ID"), Cmb_Adm_Asistencias_Permiso)
            Referencia = Cmb_Adm_Asistencias_Tipo_Permiso.ListIndex
            Dtp_Adm_Asistencias_Hora_Entrada.Value = .rdoColumns("Hora_Entrada")
            Dtp_Adm_Asistencias_Hora_Salida.Value = .rdoColumns("Hora_Salida")
            Txt_Adm_Asistencias_Horas_Aprobada.Text = .rdoColumns("Horas_Extra")
            If Not IsNull(.rdoColumns("Horas_Aprobadas")) Then
                Txt_Adm_Asistencias_Retardo.Text = .rdoColumns("Horas_Aprobadas")
            Else
                Txt_Adm_Asistencias_Retardo.Text = 0
            End If
            Btn_Regresar_Click
        End With
    End If
    Rs_Consulta_Adm_Asistencias.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Adm_Inasistencias
'DESCRIPCION: Actualiza el registro de la falta
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 03-Septiembre-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Adm_Asistencias()
Dim Mi_SQL As String
Dim Rs_Modifica_Adm_Inasistencias As rdoResultset 'Informacion del Maquinas
Dim Rs_Consulta_Adm_Movimientos As rdoResultset   'Informacion  del movimiento
Dim Rs_Consulta_Turno As rdoResultset
Dim Movimiento As String
Dim Fecha_Inicio As Date
Dim Fecha_Termino As Date
Dim Horas_Turno As Double

On Error GoTo handler
    Conexion_Base.BeginTrans
    'Identifica la fecha del registro
    Mi_SQL = "SELECT * FROM Adm_Asistencias"
    Mi_SQL = Mi_SQL & " WHERE No_Asistencia='" & Trim(Txt_Adm_Asistencias_No_Asistencias.Text) & "'"
    Set Rs_Modifica_Adm_Inasistencias = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Adm_Inasistencias.EOF Then
        With Rs_Modifica_Adm_Inasistencias
            'Modifica la incidencia cancelando el registro
            Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
            Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & .rdoColumns("Referencia") & "'"
            Mi_SQL = Mi_SQL & " AND Tipo_Incidencia='" & .rdoColumns("Tipo_Incidencia") & "'"
            Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Consulta_Adm_Movimientos.EOF Then
                Rs_Consulta_Adm_Movimientos.Edit
                    Rs_Consulta_Adm_Movimientos.rdoColumns("Estatus") = "C"
                Rs_Consulta_Adm_Movimientos.Update
            End If
            Rs_Consulta_Adm_Movimientos.Close
            'Cambia el registro de la asistencia
            .Edit
                If .rdoColumns("Turno_ID") <> Format(Cmb_Adm_Asistencias_Permiso.ItemData(Cmb_Adm_Asistencias_Permiso.ListIndex), "00000") Then
                    'Consulta datos del turno seleccionado
                    Mi_SQL = "SELECT * FROM Cat_Turnos WHERE Turno_ID='" & Format(Cmb_Adm_Asistencias_Permiso.ItemData(Cmb_Adm_Asistencias_Permiso.ListIndex), "00000") & "'"
                    Set Rs_Consulta_Turno = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Turno.EOF Then
                        .rdoColumns("Turno_ID") = Format(Cmb_Adm_Asistencias_Permiso.ItemData(Cmb_Adm_Asistencias_Permiso.ListIndex), "00000")
                        .rdoColumns("Hora_Entrada_Turno") = Format(Rs_Consulta_Turno.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                        .rdoColumns("Hora_Salida_Turno") = Format(Rs_Consulta_Turno.rdoColumns("Hora_Termino"), "HH:mm:ss")
                    End If
                    Rs_Consulta_Turno.Close
                End If
                .rdoColumns("Hora_Entrada") = Format(Dtp_Adm_Asistencias_Hora_Entrada.Value, "HH:mm:ss")
                .rdoColumns("Hora_Salida") = Format(Dtp_Adm_Asistencias_Hora_Salida.Value, "HH:mm:ss")
                Horas_Turno = Format(DateDiff("n", Format(Dtp_Adm_Asistencias_Hora_Entrada.Value, "HH:mm:ss"), Format(Dtp_Adm_Asistencias_Hora_Salida.Value, "HH:mm:ss")) / 60, "#0.00")
                If Horas_Turno >= 0 Then
                    '.rdoColumns("Horas_Aprobadas") = Horas_Turno + Val(Txt_Adm_Asistencias_Horas_Aprobada.Text)
                    .rdoColumns("Horas_Aprobadas") = Horas_Turno
                Else
                    '.rdoColumns("Horas_Aprobadas") = (24 + Horas_Turno) + Val(Txt_Adm_Asistencias_Horas_Aprobada.Text)
                    .rdoColumns("Horas_Aprobadas") = (24 + Horas_Turno)
                End If
                .rdoColumns("Horas_Extra") = Val(Txt_Adm_Asistencias_Horas_Aprobada.Text)
                .rdoColumns("Simbologia") = Cmb_Adm_Asistencias_Tipo_Permiso.Text
                .rdoColumns("SubSimbologia") = ""
                .rdoColumns("Referencia") = ""
                .rdoColumns("Tipo_Incidencia") = ""
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modifica_Adm_Inasistencias.Close
    MsgBox "La asistencia ha sido modificada", vbInformation + vbOKOnly, Me.Caption
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir o modificar los valoes
    Fra_Adm_Asistencias_General.Enabled = False
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Asistencias_Mantenimiento", Me)
    Movimiento_Empleado = ""
Exit Sub
handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbExclamation + vbOKOnly, Me.Caption
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Enviar_Correo
    'DESCRIPCIÓN:           Envia el correo con los parametros establecido
    'PARÁMETROS :           From_Email: correo de quien envia
    '                       Nombre_From: Nombre quien envia
    '                       To_Email:correo a quien se envia
    '                       Nombre_To: nombre a quien se envia
    '                       Asunto: asunto del correo
    '                       Mensaje_Email: mensaje del correo
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           19 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Enviar_Correo(From_Email As String, Nombre_From As String, To_Email As String, Nombre_To As String, Asunto As String, Mensaje_Email As String)
    Set poSendMail = New clsSendMail
    Me.MousePointer = vbHourglass

    With poSendMail
        ' Propiedades opcionales para envio de correo, deberan ser primero configuradas si se utilizan
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
        ' Propiedades básicas para envio de correos
        .SMTPHost = Servidor_SMTP           ' Required the fist time, optional thereafter
        .From = From_Email                  ' Required the fist time, optional thereafte
        .FromDisplayName = Nombre_From      ' Optional, saved after first use
        .Recipient = To_Email               ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = Nombre_To   ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName          ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc              ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text      ' Optional, used when different than 'From' address
        .Subject = Asunto                   ' Optional
        .Message = Mensaje_Email            ' Optional
        '.Attachment = Trim(txtAttach.Text)  ' Optional, separate multiple entries with delimiter character

        ' Propiedades opcionales adicionales, utilizar si son requeridas por la aplicacion
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        '.UserName = txtUserName                     ' Optional, default = Null String
        '.Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' Propiedades avanzadas, cambiar solo si tienes una buena razon para hacerlos
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
         .SMTPPort = Puerto_SMTP                    ' Optional, default = 25

        ' Envio de correo
        ' .Connect                                  ' Optional, use when sending bulk mail
        .send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'txtServer.Text = .SMTPhost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Set poSendMail = Nothing
    Me.MousePointer = vbDefault
End Function

Private Sub Llena_Combo_Turnos_Incidencias()
Dim Rs_Cat_Tipos_Faltas As rdoResultset             'Informacion de los tipos de faltas

    'Consulta los tipos de incidencias
    Mi_SQL = "SELECT Tipo_Falta_ID,Simbologia FROM Cat_Tipos_Faltas ORDER BY Simbologia"
    Set Rs_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Cat_Tipos_Faltas.EOF
        Cmb_Adm_Asistencias_Tipo_Permiso.AddItem Rs_Cat_Tipos_Faltas.rdoColumns("Simbologia")
        Cmb_Adm_Asistencias_Tipo_Permiso.ItemData(Cmb_Adm_Asistencias_Tipo_Permiso.NewIndex) = Rs_Cat_Tipos_Faltas.rdoColumns("Tipo_Falta_ID")
        Rs_Cat_Tipos_Faltas.MoveNext
    Wend
    Rs_Cat_Tipos_Faltas.Close
    'Consulta los turnos
    Mi_SQL = "SELECT Turno_ID,Nombre FROM Cat_Turnos"
    Set Rs_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Cat_Tipos_Faltas.EOF
        Cmb_Adm_Asistencias_Permiso.AddItem Rs_Cat_Tipos_Faltas.rdoColumns("Nombre")
        Cmb_Adm_Asistencias_Permiso.ItemData(Cmb_Adm_Asistencias_Permiso.NewIndex) = Rs_Cat_Tipos_Faltas.rdoColumns("Turno_ID")
        Rs_Cat_Tipos_Faltas.MoveNext
    Wend
    Rs_Cat_Tipos_Faltas.Close
End Sub

