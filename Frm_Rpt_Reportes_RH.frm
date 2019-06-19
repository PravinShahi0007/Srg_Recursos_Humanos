VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Rpt_Reportes_RH 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Exportar_PDF 
      Caption         =   "Exportar PDF"
      Enabled         =   0   'False
      Height          =   660
      Left            =   1920
      Picture         =   "Frm_Rpt_Reportes_RH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   124
      Tag             =   "A"
      Top             =   6960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton Btn_Regresar 
      Caption         =   "Regresar"
      Enabled         =   0   'False
      Height          =   660
      Left            =   7920
      Picture         =   "Frm_Rpt_Reportes_RH.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   660
      Left            =   0
      Picture         =   "Frm_Rpt_Reportes_RH.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "A"
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Height          =   660
      Left            =   11880
      Picture         =   "Frm_Rpt_Reportes_RH.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton Btn_Exportar 
      Caption         =   "Exportar EXCEL"
      Enabled         =   0   'False
      Height          =   660
      Left            =   3960
      Picture         =   "Frm_Rpt_Reportes_RH.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "A"
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComDlg.CommonDialog Cmd_Exportar 
      Left            =   9045
      Top             =   6975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Prbar_Exportacion 
      Height          =   285
      Left            =   5130
      TabIndex        =   38
      Top             =   7410
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Pic_Reportes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6900
      Left            =   0
      ScaleHeight     =   6870
      ScaleWidth      =   6045
      TabIndex        =   36
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton Btn_Salir_Reporte 
         Caption         =   "Salir"
         Height          =   660
         Left            =   4590
         Picture         =   "Frm_Rpt_Reportes_RH.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3150
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Rpt_Generar 
         Caption         =   "Reporte"
         Height          =   660
         Left            =   990
         Picture         =   "Frm_Rpt_Reportes_RH.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "C"
         Top             =   3120
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Rpt_Informacion 
         Height          =   1050
         Left            =   90
         TabIndex        =   39
         Top             =   3420
         Visible         =   0   'False
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   1852
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Rpt_Informacion_Tmp 
         Height          =   1080
         Left            =   90
         TabIndex        =   40
         Top             =   4485
         Visible         =   0   'False
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   1905
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Rpt_Informacion_Tmp_2 
         Height          =   1170
         Left            =   90
         TabIndex        =   45
         Top             =   5580
         Visible         =   0   'False
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   2064
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin VB.Frame Fra_Accesos_Almacen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accesos al Almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   160
         Top             =   0
         Visible         =   0   'False
         Width           =   6210
         Begin VB.ComboBox Cmb_Rpt_Empleado_Accesos_Almacenes 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":26C6
            Left            =   1200
            List            =   "Frm_Rpt_Reportes_RH.frx":26C8
            TabIndex        =   162
            Top             =   720
            Width           =   4545
         End
         Begin VB.CheckBox Chk_Rpt_Accesos_Almacen_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   120
            TabIndex        =   161
            Top             =   1245
            Width           =   825
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Accesos_Almacenes_Fecha_Termino 
            Height          =   315
            Left            =   1200
            TabIndex        =   163
            Top             =   1680
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Accesos_Almacenes_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   164
            Top             =   1200
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Lbl_Empleado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empleado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   165
            Top             =   720
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Cursos_Tomados_Por_Empleado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Cursos Tomados por Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   111
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.TextBox Txt_No_Tarjeta_Cursos_Por_Empleado 
            Height          =   285
            Left            =   1200
            TabIndex        =   155
            Top             =   600
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado 
            Height          =   315
            Left            =   1200
            TabIndex        =   113
            Top             =   960
            Width           =   4515
         End
         Begin VB.CheckBox Chk_Rpt_Cursos_Tomados_Por_Empleado_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   120
            TabIndex        =   112
            Top             =   1485
            Width           =   825
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Fin 
            Height          =   315
            Left            =   1200
            TabIndex        =   114
            Top             =   1950
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   115
            Top             =   1380
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Label46 
            BackColor       =   &H8000000E&
            Caption         =   "No. Tarjeta"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   960
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Rpt_Empleados_No_Validados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Empleados No Validados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Empleados_No_Validados_Departamento 
            Height          =   315
            Left            =   1110
            TabIndex        =   31
            Top             =   360
            Width           =   4680
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_No_Validados_Supervisor 
            Height          =   315
            Left            =   1110
            TabIndex        =   32
            Top             =   870
            Width           =   4680
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_No_Validados_Fecha_Termino 
            Height          =   315
            Left            =   1110
            TabIndex        =   34
            Top             =   1890
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_No_Validados_Fecha_Inicio 
            Height          =   315
            Left            =   1110
            TabIndex        =   33
            Top             =   1380
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            Height          =   195
            Left            =   90
            TabIndex        =   75
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   70
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino"
            Height          =   195
            Left            =   90
            TabIndex        =   69
            Top             =   1950
            Width           =   570
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   90
            TabIndex        =   68
            Top             =   930
            Width           =   750
         End
      End
      Begin VB.Frame Fra_Rpt_Asistencia_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Asistencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Departamento 
            Height          =   315
            Left            =   975
            TabIndex        =   7
            Top             =   738
            Width           =   4770
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Supervisor 
            Height          =   315
            Left            =   975
            TabIndex        =   8
            Top             =   1161
            Width           =   4770
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Empleado 
            Height          =   315
            Left            =   975
            TabIndex        =   9
            Top             =   1584
            Width           =   4770
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Empresa 
            Height          =   315
            Left            =   975
            TabIndex        =   6
            Top             =   315
            Width           =   4770
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Asistencia_Empleados_Fecha_Termino 
            Height          =   315
            Left            =   975
            TabIndex        =   11
            Top             =   2430
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio 
            Height          =   315
            Left            =   975
            TabIndex        =   10
            Top             =   2007
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departameto"
            Height          =   195
            Left            =   45
            TabIndex        =   71
            Top             =   795
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   45
            TabIndex        =   66
            Top             =   1215
            Width           =   750
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino"
            Height          =   195
            Left            =   45
            TabIndex        =   61
            Top             =   2490
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Left            =   45
            TabIndex        =   60
            Top             =   2070
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
            Height          =   195
            Left            =   45
            TabIndex        =   59
            Top             =   1650
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   45
            TabIndex        =   58
            Top             =   375
            Width           =   615
         End
      End
      Begin VB.Frame Fra_Rpt_Historico_Permisos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Historico de Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Historico_Permisos_Departamento 
            Height          =   315
            Left            =   1065
            TabIndex        =   13
            Top             =   858
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Permisos_Supervisor 
            Height          =   315
            Left            =   1065
            TabIndex        =   14
            Top             =   1266
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Permisos_Empleado 
            Height          =   315
            Left            =   1065
            TabIndex        =   15
            Top             =   1674
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Permisos_Empresa 
            Height          =   315
            Left            =   1065
            TabIndex        =   12
            Top             =   450
            Width           =   4725
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Historico_Permisos_Fecha_inicio 
            Height          =   315
            Left            =   1065
            TabIndex        =   16
            Top             =   2085
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Historico_Permisos_Fecha_Termino 
            Height          =   315
            Left            =   1065
            TabIndex        =   17
            Top             =   2490
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Departamento"
            Height          =   195
            Left            =   45
            TabIndex        =   72
            Top             =   915
            Width           =   1005
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   45
            TabIndex        =   64
            Top             =   1326
            Width           =   750
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   45
            TabIndex        =   56
            Top             =   510
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            Height          =   195
            Left            =   45
            TabIndex        =   55
            Top             =   1734
            Width           =   705
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Inicio"
            Height          =   195
            Left            =   45
            TabIndex        =   54
            Top             =   2142
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Termino"
            Height          =   195
            Left            =   45
            TabIndex        =   53
            Top             =   2550
            Width           =   570
         End
      End
      Begin VB.Frame Fra_Faltas_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Faltas del Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   105
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Turno_Faltas 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":26CA
            Left            =   1050
            List            =   "Frm_Rpt_Reportes_RH.frx":26CC
            TabIndex        =   110
            Top             =   1140
            Width           =   4710
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Faltas_Empleados 
            Height          =   315
            Left            =   1065
            TabIndex        =   106
            Top             =   2010
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Faltas_Empleados_Fin 
            Height          =   315
            Left            =   1065
            TabIndex        =   108
            Top             =   2445
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Lbl_Turno_Faltas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Turnos"
            Height          =   195
            Left            =   195
            TabIndex        =   109
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Lbl_Fechas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   195
            TabIndex        =   107
            Top             =   2070
            Width           =   525
         End
      End
      Begin VB.Frame Fra_Rpt_Empleados_Baja 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Empleados Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Empleados_Baja_Empresa 
            Height          =   315
            Left            =   1170
            TabIndex        =   92
            Top             =   360
            Width           =   4635
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_Baja_Puesto 
            Height          =   315
            Left            =   1170
            TabIndex        =   78
            Top             =   1320
            Width           =   4635
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_Baja_Departamento 
            Height          =   315
            Left            =   1170
            TabIndex        =   77
            Top             =   840
            Width           =   4635
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_Baja_Fecha_Termino 
            Height          =   315
            Left            =   1170
            TabIndex        =   79
            Top             =   2295
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755843
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_Baja_Fecha_Inicio 
            Height          =   315
            Left            =   1170
            TabIndex        =   80
            Top             =   1815
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755843
            CurrentDate     =   39872
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   93
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puesto"
            Height          =   195
            Left            =   90
            TabIndex        =   84
            Top             =   1386
            Width           =   495
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino"
            Height          =   195
            Left            =   90
            TabIndex        =   83
            Top             =   2355
            Width           =   570
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   82
            Top             =   1869
            Width           =   375
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            Height          =   195
            Left            =   90
            TabIndex        =   81
            Top             =   903
            Width           =   1005
         End
      End
      Begin VB.Frame Fra_Rpt_Historico_Faltas_Retardos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Historico de Faltas y Retardos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Historico_Faltas_Retardos_Departamento 
            Height          =   315
            Left            =   1065
            TabIndex        =   19
            Top             =   840
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Faltas_Retardos_Supervisor 
            Height          =   315
            Left            =   1065
            TabIndex        =   20
            Top             =   1230
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Faltas_Retardos_Empleado 
            Height          =   315
            Left            =   1065
            TabIndex        =   21
            Top             =   1620
            Width           =   4725
         End
         Begin VB.ComboBox Cmb_Rpt_Historico_Faltas_Retardos_Empresa 
            Height          =   315
            Left            =   1065
            TabIndex        =   18
            Top             =   450
            Width           =   4725
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Inicio 
            Height          =   315
            Left            =   1065
            TabIndex        =   22
            Top             =   2010
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Termino 
            Height          =   315
            Left            =   1065
            TabIndex        =   23
            Top             =   2400
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Departamento"
            Height          =   195
            Left            =   45
            TabIndex        =   73
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   45
            TabIndex        =   65
            Top             =   1290
            Width           =   750
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Termino"
            Height          =   195
            Left            =   45
            TabIndex        =   51
            Top             =   2460
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Inicio"
            Height          =   195
            Left            =   45
            TabIndex        =   50
            Top             =   2070
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            Height          =   195
            Left            =   45
            TabIndex        =   49
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   48
            Top             =   510
            Width           =   615
         End
      End
      Begin VB.Frame Fra_Rpt_Horas_Trabajadas_Empleado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Horas Trabajadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento 
            Height          =   315
            Left            =   1110
            TabIndex        =   25
            Top             =   752
            Width           =   4680
         End
         Begin VB.ComboBox Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":26CE
            Left            =   1110
            List            =   "Frm_Rpt_Reportes_RH.frx":26DB
            TabIndex        =   28
            Top             =   1928
            Width           =   4680
         End
         Begin VB.ComboBox Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor 
            Height          =   315
            Left            =   1110
            TabIndex        =   26
            Top             =   1144
            Width           =   4680
         End
         Begin VB.ComboBox Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa 
            Height          =   315
            Left            =   1110
            TabIndex        =   24
            Top             =   360
            Width           =   4680
         End
         Begin VB.ComboBox Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado 
            Height          =   315
            Left            =   1110
            TabIndex        =   27
            Top             =   1536
            Width           =   4680
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio 
            Height          =   315
            Left            =   1110
            TabIndex        =   29
            Top             =   2325
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino 
            Height          =   315
            Left            =   1110
            TabIndex        =   30
            Top             =   2715
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            Height          =   195
            Left            =   90
            TabIndex        =   74
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo"
            Height          =   195
            Left            =   90
            TabIndex        =   63
            Top             =   1995
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   90
            TabIndex        =   62
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   46
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   1590
            Width           =   705
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   43
            Top             =   2385
            Width           =   375
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino"
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   2775
            Width           =   570
         End
      End
      Begin VB.Frame Fra_Reporte_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   97
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Estatus 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":26FC
            Left            =   90
            List            =   "Frm_Rpt_Reportes_RH.frx":2709
            TabIndex        =   153
            Top             =   1560
            Width           =   5685
         End
         Begin VB.ComboBox Cmb_Empleado_Curso 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":272D
            Left            =   90
            List            =   "Frm_Rpt_Reportes_RH.frx":272F
            TabIndex        =   100
            Top             =   945
            Width           =   5685
         End
         Begin VB.ComboBox Cmb_Curso 
            Height          =   315
            Left            =   1185
            TabIndex        =   99
            Top             =   360
            Width           =   4560
         End
         Begin VB.CheckBox Chk_Fechas_Curso 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   105
            TabIndex        =   98
            Top             =   2205
            Width           =   825
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Curso 
            Height          =   315
            Left            =   1185
            TabIndex        =   102
            Top             =   2550
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Curso 
            Height          =   315
            Left            =   1185
            TabIndex        =   101
            Top             =   2100
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Label44 
            BackColor       =   &H8000000E&
            Caption         =   "Estatus"
            Height          =   255
            Left            =   90
            TabIndex        =   152
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Lbl_Empleado_Curso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
            Height          =   195
            Left            =   90
            TabIndex        =   104
            Top             =   645
            Width           =   705
         End
         Begin VB.Label Lbl_Curso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Curso"
            Height          =   195
            Left            =   90
            TabIndex        =   103
            Top             =   420
            Width           =   405
         End
      End
      Begin VB.Frame Fra_Rpt_General_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte General de Cursos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   131
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_General_Cursos_Sala 
            Height          =   315
            Left            =   1200
            TabIndex        =   140
            Top             =   1560
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_General_Cursos_Institucion 
            Height          =   315
            Left            =   1200
            TabIndex        =   139
            Top             =   960
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_General_Cursos_Instructor 
            Height          =   315
            Left            =   1200
            TabIndex        =   138
            Top             =   360
            Width           =   4455
         End
         Begin VB.CheckBox Chk_Rpt_General_Cursos_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   120
            TabIndex        =   132
            Top             =   2085
            Width           =   825
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_General_Cursos_Fecha_Fin 
            Height          =   315
            Left            =   1200
            TabIndex        =   133
            Top             =   2550
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Genera_Cursosl_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   134
            Top             =   2040
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Label43 
            BackColor       =   &H8000000E&
            Caption         =   "Sala"
            Height          =   255
            Left            =   360
            TabIndex        =   137
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label42 
            BackColor       =   &H8000000E&
            Caption         =   "Institución"
            Height          =   255
            Left            =   360
            TabIndex        =   136
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label41 
            BackColor       =   &H8000000E&
            Caption         =   "Instructor"
            Height          =   255
            Left            =   360
            TabIndex        =   135
            Top             =   480
            Width           =   855
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
         Height          =   4005
         Left            =   120
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Rpt_Cursos_Resumen_Mensual_Auditable 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":2731
            Left            =   1200
            List            =   "Frm_Rpt_Reportes_RH.frx":273E
            TabIndex        =   143
            Top             =   1320
            Width           =   4575
         End
         Begin VB.CheckBox Chk_Rpt_Cursos_Resumen_Mensual_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   120
            TabIndex        =   127
            Top             =   1965
            Width           =   825
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":2762
            Left            =   1200
            List            =   "Frm_Rpt_Reportes_RH.frx":2764
            TabIndex        =   126
            Top             =   720
            Width           =   4545
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin 
            Height          =   315
            Left            =   1200
            TabIndex        =   128
            Top             =   2430
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   129
            Top             =   1920
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Label36 
            BackColor       =   &H8000000E&
            Caption         =   "Auditable"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Lbl_Tipo_Curso_id 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo Curso"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   720
            Width           =   765
         End
      End
      Begin VB.Frame Fra_Rpt_Empleados_Alta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   120
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.CheckBox Chk_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   105
            TabIndex        =   96
            Top             =   1845
            Width           =   825
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_Alta_Empresa 
            Height          =   315
            Left            =   1185
            TabIndex        =   94
            Top             =   360
            Width           =   4560
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_Alta_Departamento 
            Height          =   315
            Left            =   1185
            TabIndex        =   87
            Top             =   825
            Width           =   4560
         End
         Begin VB.ComboBox Cmb_Rpt_Empleados_Alta_Puesto 
            Height          =   315
            Left            =   1185
            TabIndex        =   86
            Top             =   1275
            Width           =   4560
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_Alta_Fecha_Termino 
            Height          =   315
            Left            =   1185
            TabIndex        =   88
            Top             =   2190
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Empleados_Alta_Fecha_Inicio 
            Height          =   315
            Left            =   1185
            TabIndex        =   89
            Top             =   1740
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   110755840
            CurrentDate     =   39872
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   95
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            Height          =   195
            Left            =   90
            TabIndex        =   91
            Top             =   881
            Width           =   1005
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puesto"
            Height          =   195
            Left            =   90
            TabIndex        =   90
            Top             =   1342
            Width           =   495
         End
      End
      Begin VB.Frame Fra_Rpt_Cursos_Indice_Asistencias 
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
         Height          =   4005
         Left            =   120
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.TextBox Txt_No_Tarjeta_Indices_Asistencia 
            Height          =   285
            Left            =   1200
            TabIndex        =   159
            Top             =   960
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":2766
            Left            =   1200
            List            =   "Frm_Rpt_Reportes_RH.frx":2776
            TabIndex        =   149
            Top             =   600
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda 
            Height          =   315
            Left            =   1200
            TabIndex        =   148
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CheckBox Chk_Rpt_Cursos_Indice_Asistencias_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   210
            Left            =   120
            TabIndex        =   145
            Top             =   1965
            Width           =   825
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin 
            Height          =   315
            Left            =   1200
            TabIndex        =   146
            Top             =   2430
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   147
            Top             =   1920
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Lbl_No_Tarjeta_Indices_Asistencia 
            BackColor       =   &H8000000E&
            Caption         =   "No Tarjeta"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label45 
            BackColor       =   &H8000000E&
            Caption         =   "Tipo de Busqueda"
            Height          =   495
            Left            =   120
            TabIndex        =   151
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda 
            BackColor       =   &H8000000E&
            Caption         =   "Empleado"
            Height          =   375
            Left            =   120
            TabIndex        =   150
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Fra_Cursos_Horas_Hombre 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cursos Horas Hombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   117
         Top             =   0
         Visible         =   0   'False
         Width           =   5850
         Begin VB.TextBox Txt_No_Tarjeta_Cursos_Hioras_Hombre 
            Height          =   285
            Left            =   1200
            TabIndex        =   157
            Top             =   960
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CheckBox Chk_Rpt_Cursos_Hora_Hombre_Fechas 
            BackColor       =   &H8000000E&
            Caption         =   "Fechas"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Hora_Hombre_Busqueda 
            Height          =   315
            Left            =   1200
            TabIndex        =   123
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.ComboBox Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes_RH.frx":27AE
            Left            =   1200
            List            =   "Frm_Rpt_Reportes_RH.frx":27BE
            TabIndex        =   122
            Top             =   600
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino 
            Height          =   315
            Left            =   1200
            TabIndex        =   118
            Top             =   2640
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   119
            Top             =   2040
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMMM yyyy"
            Format          =   110755843
            CurrentDate     =   41039
         End
         Begin VB.Label Lbl_No_Tarjeta_Cursos_Horas_Hombre 
            BackColor       =   &H8000000E&
            Caption         =   "No_Tarjeta"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Lbl_Rpt_Cursos_Hora_Hombre_Busqueda 
            BackColor       =   &H8000000E&
            Caption         =   "Empleado"
            Height          =   375
            Left            =   120
            TabIndex        =   121
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label40 
            BackColor       =   &H8000000E&
            Caption         =   "Tipo de Busqueda"
            Height          =   495
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   1095
         End
      End
   End
   Begin RichTextLib.RichTextBox Rich_Reporte 
      Height          =   6915
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   12197
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"Frm_Rpt_Reportes_RH.frx":27F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Lbl_Progreso_Exportacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exportando..."
      Height          =   195
      Left            =   5160
      TabIndex        =   37
      Top             =   7065
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "Frm_Rpt_Reportes_RH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim linea As String                     'Lleva el control del nuemro de linea que se esta acediendo en la lectura de un archivo
Dim Contador As Integer                 'Sirve para iterra por los bucles
Dim Orientacion_Reporte As Orientacion  'Indica la orientacion del reporte para su exportacion a excel
Public Reporte As String                   'Nombre del reporte a manejar en la forma
Public Archivo_Reporte_Abierto As Boolean   'Indica si el archivo para reporte esta abierto
Private Enum Orientacion
    Horizontal = 1
    Vertical = 2
End Enum

Private Sub Btn_Exportar_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String

On Error GoTo HANDLER
    Cmd_Exportar.CancelError = True
    Cmd_Exportar.DialogTitle = "Seleccione el directorio"
    Cmd_Exportar.Flags = cdlOFNHideReadOnly
    Cmd_Exportar.Filter = "Archivos de Excel(*.xls)|*.xls"
    Cmd_Exportar.FilterIndex = 2
    Cmd_Exportar.FileName = Reporte & ".xls"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
        Call Exportar_Excel_Bien(Ruta_Temporal & Reporte & "xls.txt", Ruta_Exportacion)
    End If
    'Display name of selected file
    Exit Sub
HANDLER:
    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub Exportar_Excel_Bien(Archivo_Exportar As String, Ruta As String)
Dim obj_Excel As Object
Dim Celda As Object
Dim Fila As Integer, Columna As Integer
Dim Contenido As String, Lineas As Variant
Dim Datos As Variant, MC As Integer
Dim Encabezado As Boolean
Dim Fila_Encabezado As Integer

Dim Izquierda As Single
Dim Arriba As Single
Dim Ancho As Double
Dim Alto As Double
Dim Carpeta_Foto As String

On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    'Lee el contenido del reporte
    Open Archivo_Exportar For Input As #2
    Contenido = Input$(LOF(2), #2)
    Close
    Lbl_Progreso_Exportacion.Caption = "Exportando ..."
    Lbl_Progreso_Exportacion.Visible = True
    Prbar_Exportacion.Visible = True
    Prbar_Exportacion.Value = 0
    Prbar_Exportacion.Min = 0
    'Nuevo objeto Excel
    Set obj_Excel = CreateObject("Excel.Application")
    With obj_Excel
        'Agrega un libro
        .Workbooks.Add
        ' Obtiene el número de líneas del Csv con la función split
        Lineas = Split(Contenido, vbCrLf)
        Prbar_Exportacion.Max = UBound(Lineas) + 1
        For Fila = 0 To UBound(Lineas)
            'Encabezado = False
            'Separa los datos de la linea
            Datos = Split(Lineas(Fila), "|")
            'Recorre los datos de esta fila que corresponden a cada campo
            For Columna = 0 To UBound(Datos)
                ' Agrega el dato a la celda de la hoja activa
                .ActiveSheet.Cells(Fila + 1, Columna + 1) = Datos(Columna)
                If Trim(Mid(Datos(Columna), 1, 1)) = "." Then
                    Encabezado = True
                    Fila_Encabezado = Fila
                End If
                If Encabezado Then
                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Borders.LineStyle = 1
                    '.ActiveSheet.Cells(Fila + 1, Columna + 1).Borders.Weight = 0
                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Borders.Color = RGB(0, 0, 0)
                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Font.FontStyle = "Bold"
                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Font.Size = 12
                End If
            Next
            If UBound(Datos) > 0 Then
                If UCase(Trim(.ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1))) Like "PERFIL*" Then
                    Set Celda = .ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1)
'                    Ancho = Celda.Offset(0, 1).Left - Celda.Left
'                    Izquierda = Celda.Left + Ancho / 2 - 340 / 2
'                    If Izquierda < 1 Then
'                        Izquierda = 1
'                    End If
'                    Alto = Celda.Offset(1, 0).Top - Celda.Top
'                    Arriba = Celda.Top + Alto / 2 - 255 / 2
'                    If Arriba < 1 Then
'                        Arriba = 1
'                    End If
                    Carpeta_Foto = Replace(Trim(.ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1)), "Perfil\", "")
                    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & Carpeta_Foto, "ARCHIVO") = True Then
                        .ActiveSheet.Shapes.AddPicture PG_Ruta_Fotos & "\" & Carpeta_Foto, 0, 1, Celda.Left, Celda.Top, 50, 50
                        .ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1).RowHeight = 50
                    End If
                    .ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1) = ""
'                    .ActiveSheet.Cells(Fila + 1, UBound(Datos) + 1) = PG_Ruta_Fotos & "\" & Carpeta_Foto
                End If
            End If
            If MC < Columna Then
               MC = Columna
            End If
            Prbar_Exportacion.Value = Prbar_Exportacion.Value + 1
        Next
        'Selecciona toda la hoja
        .ActiveSheet.UsedRange.Select
        'Autoajusta las columnas
        '.Selection.Columns.AutoFit
        'Selecciona el encabezado
    End With
    ' Aplica atributos a la fuente a la selección anterior ( los encabezados )
    With obj_Excel.Selection.Font
        '.Name = "Verdana"
        '.FontStyle = "Bold"
        '.Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        '.Underline = xlUnderlineStyleNone
    End With
    'se ajusta el documento a una hoja
    'With obj_Excel.Selection
    Lbl_Progreso_Exportacion.Caption = "Guardando ..."
    ' Guarda el documento Xls
    obj_Excel.ActiveWorkbook.SaveAs _
        FileName:=Ruta, _
        Password:="", _
        WriteResPassword:="", _
        ReadOnlyRecommended:=False, _
        CreateBackup:=False
    'obj_Excel.ActiveWorkbook.Close False
    Lbl_Progreso_Exportacion.Visible = False
    Prbar_Exportacion.Visible = False
    MDIFrm_Apl_Principal.MousePointer = 0
     'Cierra el archivo y elimina la variable
     If MsgBox("¿Desea abrir el archivo?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        obj_Excel.Visible = True
     Else
        obj_Excel.Quit
        MsgBox "Reportes exportado", vbInformation + vbOKOnly, Me.Caption
    End If
    'obj_Excel.Quit
    Set obj_Excel = Nothing
    Exit Sub
Exit Sub
 ' Error
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox Err.Description
    On Error Resume Next
    If Not obj_Excel Is Nothing Then
        obj_Excel.Quit
        Set obj_Excel = Nothing
    End If
    Lbl_Progreso_Exportacion.Visible = False
    Prbar_Exportacion.Visible = False
End Sub

Private Sub Btn_Exportar_PDF_Click()
Dim Nombre As String
Dim Nombre_RPT As String
Dim Hoora As Date
Hoora = Format$(Now, "d-mmmm-yy h:mm:ss")
Dim hora As String
hora = Replace(Hoora, " ", "")
hora = Replace(hora, ":", "_")
hora = Replace(hora, ".", "")
hora = Replace(hora, "/", "")
    Select Case Reporte
        Case "Cursos_Por_Empleado":
            Nombre_RPT = "Rpt_Cursos_Por_Empleado"
            Nombre = "Cursos_Por_Empleado_" & hora
            
        Case "Cursos_Hora_Hombre"
            If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0 Then
                Nombre_RPT = "Rpt_Cursos_Horas_Hombre_General"
                Nombre = "Cursos_Horas_Hombre_" & hora
            End If
            
            If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 1 Then
                Nombre_RPT = "Rpt_Cursos_Horas_Hombre_Empleados"
                Nombre = "Cursos_Horas_Hombre_" & hora
            End If
            
            If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 2 Then
         
                Nombre_RPT = "Rpt_Cursos_Horas_Hombre_Cursos"
                Nombre = "Cursos_Horas_Hombre_" & hora
             End If
             
             If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 3 Then
         
                Nombre_RPT = "Rpt_Cursos_Horas_Hombre_Departamentos"
                Nombre = "Cursos_Horas_Hombre_" & hora
             End If
    
            Case "Cursos_Indice_Asistencia"
                Nombre_RPT = "Rpt_Cursos_Indices_Asistencia"
                Nombre = "Cursos_Indices_Asistencia_" & hora
            
            
            Case "Cursos_Resumen_Mensual"
                Nombre_RPT = "Rpt_Cursos_Resumen_Mensual"
                Nombre = "Cursos_Resumen_Mensual_" & hora
           
        
         Case "Reporte_General_Cursos"
                Nombre_RPT = "Rpt_Cursos_Reporte_General"
                Nombre = "Cursos_Reporte_General_" & hora
            End Select
        
Crea_PDF Nombre_RPT, Nombre
End Sub

Private Sub Btn_Imprimir_Click()
Dim linea As String 'Obtiene el texto a imprimir
Dim X As Printer
Dim contar_linea As Integer
Dim Foto_Empleado As New StdPicture
Dim No_Tarjeta As String
Dim Cont_Fila As Integer
Dim Cordenada_Y_Imagen  As Double
Dim Cont_Saltos As Integer
Dim Mi_SQL As String
Dim Rs_Conssultar_Foto As rdoResultset
Dim Contar_Filas As Integer
Dim Numero_Pagina As Integer


On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    Cordenada_Y_Imagen = 0
    'If Reporte = "Reporte_COI" Or Reporte = "Asistencias_Empleados" Then Printer.Orientation = vbPRORLandscape
    If Reporte = "Reporte_COI" Then Printer.Orientation = vbPRORLandscape
    'If (Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "MENSUAL" And Reporte = "Horas_Trabajadas_Empleado") Or Reporte = "Asistencias_Empleados" Then
    If (Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "MENSUAL" And Reporte = "Horas_Trabajadas_Empleado") Then
        Printer.Orientation = vbPRORLandscape
        Printer.FontSize = 7
        Printer.Font = "COURIER NEW"
        Printer.Print
        Printer.FontSize = 11
        Printer.Font = "COURIER NEW"
        Printer.Print
        Printer.FontSize = 7
        Printer.Font = "Courier New"
    Else
        Printer.FontSize = 8
        Printer.Font = "COURIER NEW"
        Printer.Print
        Printer.FontSize = 11
        Printer.Font = "COURIER NEW"
        Printer.Print
        Printer.FontSize = 8
        Printer.Font = "Courier New"
    End If
'    Printer.FontSize = 8
'    Printer.Font = "COURIER NEW"
'    Printer.Print
'    Printer.FontSize = 11
'    Printer.Font = "COURIER NEW"
'    Printer.Print
'    Printer.FontSize = 8
'    Printer.Font = "Courier New"

''    If Reporte = "Historico_Faltas_Retardos" Then
''        Open Ruta_Temporal & Reporte & ".txt" For Input As #1
''        Numero_Pagina = 1
''        Do While Not EOF(1)
''            ''Si es cambio de pagina reinicia los contadores
''            If Contar_Filas >= 90 Then
''                Contar_Filas = 0
''                Cordenada_Y_Imagen = 0
''                Cont_Saltos = 2
''            End If
''            contar_linea = contar_linea + 1
''            Contar_Filas = Contar_Filas + 1
''            If contar_linea = 90 Then
''                Printer.NewPage
''            End If
''            Line Input #1, linea
''            If Val(Mid(Trim(linea), 1, 4)) > 4000 Then
''                ''Se consulta la foto
''                Mi_SQL = "  SELECT Imagen_Perfil  FROM Cat_Empleados "
''                Mi_SQL = Mi_SQL & " WHERE No_Tarjeta = '" & Mid(Trim(linea), 1, 4) & "'"
''                Set Rs_Conssultar_Foto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
''                If Not Rs_Conssultar_Foto.EOF Then
''                    If Not IsNull(Rs_Conssultar_Foto!Imagen_Perfil) Then
''                        Set Foto_Empleado = LoadPicture(App.Path & "\Perfil\" & Rs_Conssultar_Foto!Imagen_Perfil)
''                        ''se imprimen espacios en blanco
''                        For Cont_Fila = 1 To 6 Step 1
''                            Printer.Print ""
''                            Contar_Filas = Contar_Filas + 1
''                            contar_linea = contar_linea + 1
''                        Next
''                        ''Se imprime la fila
''                        Printer.Print linea
''                        Cordenada_Y_Imagen = ((Contar_Filas + Cont_Saltos) * 153)
''                        Cont_Saltos = Cont_Saltos + 2
''                        ''Se hace la impresion de la foto
''                        Call Printer.PaintPicture(Foto_Empleado, 5300, Cordenada_Y_Imagen, 800, 800)
''                    Else
''                        Printer.Print linea
''                    End If
''                End If
''                Rs_Conssultar_Foto.Close
''            Else
''                Printer.Print linea
''            End If
''        Loop
''        Printer.EndDoc
''        Close #1
''    Else
        Open Ruta_Temporal & Reporte & ".txt" For Input As #1
        Do While Not EOF(1)
            contar_linea = contar_linea + 1
            If contar_linea = 90 Then
                Printer.NewPage
            End If
            Line Input #1, linea
            If Reporte = "Asistencias_Empleados" And linea = "." Then
                Printer.NewPage
                linea = ""
            End If
            Printer.Print linea
        Loop
        Printer.EndDoc
        Close #1
''    End If
    MsgBox "Reporte enviado a impresora", vbInformation + vbOKOnly, Me.Caption
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
HANDLER:
    Printer.EndDoc
    Close #1
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cargar_Frame
    'DESCRIPCIÓN: Este proceso hace visible la frame que necesitamos ocultando
    '             todo las demas a la vez ajusta el tamaño de la forma
    'PARÁMETROS:
    '             1. Frame: Nombre del Frame el cual vamos hacer visible
    '             2. Formulario: Nombre del Formulario al cual nos estamos refiriendo
    'CREO: Ruben García
    'FECHA_CREO:18 sep 05
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
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

Private Sub Btn_Regresar_Click()
Dim Control As Control
    Me.Width = 5180
    Me.Height = 3650
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
    Me.Top = 100
    For Each Control In Me.Controls
        If TypeOf Control Is Frame Then
            If Control.Visible = True Then
                Exit For
            End If
        End If
    Next
    Control.Top = 0
    Control.Left = 0
    Me.Width = Control.Width + 200
    Me.Height = Control.Height + 600
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Rich_Reporte.Visible = False
    Pic_Reportes.Visible = True
    Btn_Imprimir.Enabled = False
    Btn_Exportar.Enabled = False
    Btn_Regresar.Enabled = False
    Btn_Salir.Enabled = False
End Sub

Private Sub Btn_Rpt_Generar_Click()
On Error GoTo HANDLER
    Select Case Reporte
        Case "Horas_Trabajadas_Empleado":
            Select Case Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text
                Case "ACUMULADO":
                    Generar_Reporte_Horas_Trabajadas_Empleado
                Case "SEMANAL":
                    If DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value) > 6 Then
                        MsgBox "El periodo seleccionado rebasa una semana," + vbCrLf + _
                               "favor de verificar", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                    Generar_Reporte_Horas_Trabajadas_Empleado_Periodo
                Case "MENSUAL":
                    If DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value) > 31 Then
                        MsgBox "El periodo seleccionado rebasa un mes," + vbCrLf + _
                               "favor de verificar", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                    Generar_Reporte_Horas_Trabajadas_Empleado_Periodo
            End Select
        Case "Historico_Faltas_Retardos":
            Generar_Reporte_Historico_Faltas_Retardos
        Case "Historico_Permisos":
            Generar_Reporte_Historico_Permisos
        Case "Empleados_No_Validados"
            Generar_Reporte_Empleados_No_Validados
        Case "Asistencias_Empleados"
            Generar_Reporte_Asistencia_Empleados
        Case "Empleados_Baja"
            Generar_Reporte_Empleados_Baja
        Case "Empleados_Alta"
            Generar_Reporte_Empleados_Alta
        Case "Reporte_Curso"
            If Cmb_Curso.ListIndex = -1 And Cmb_Empleado_Curso.ListIndex = -1 Then
                MsgBox "Seleccione un curso o un empleado", vbExclamation
            Else
                If Cmb_Curso.ListIndex > -1 Then
                    Generar_Reporte_Curso_Empleados
                Else
                    Generar_Reporte_Empleado_Cursos
                End If
            End If
        Case "Empleados_Faltas":
            Generar_Reporte_Faltas_Empleados
        Case "Empleados_Faltas_Validadas":
            Generar_Reporte_Faltas_Empleados_Validadas
        Case "Reporte_Comedor"
            If Cmb_Empleado_Curso.ListIndex > -1 Then
                Generar_Reporte_Empleado_Comidas
                
            Else
                Generar_Reporte_Comidas_Empleados
            End If
        Case "Empleados_Huella_Comedor"
            Generar_Reporte_Empleados_Huella_Comedor
        Case "Cursos_Por_Empleado"
            Generar_Reporte_Cursos_Por_Empleado
        Case "Cursos_Hora_Hombre"
            Generar_Reporte_Cursos_Hora_Hombre
        Case "Cursos_Indice_Asistencia"
            Generar_Reporte_Cursos_Indice_Asistencias
        Case "Cursos_Resumen_Mensual"
            Generar_Reporte_Cursos_Resumen_Mensual
        Case "Reporte_General_Cursos"
            Generar_Reporte_General_Cursos
        Case "Accesos_Almacenes"
            Generar_Reporte_Accesos_Almacenes
    End Select
Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    If Archivo_Reporte_Abierto Then
        Close #1, #2
        Archivo_Reporte_Abierto = False
    End If
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Btn_Salir_Reporte_Click()
    Unload Me
End Sub


Private Sub Cmb_Curso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID,Nombre", "Cat_Cursos", Cmb_Curso, "1", "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub



Private Sub Cmb_Empleado_Curso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(Cmb_Empleado_Curso.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Empleado_Curso, 1, "Nombre", "AND Estatus='A'")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Empleado_Curso, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR " & "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%')", False, "")
          
        End If
        If Cmb_Empleado_Curso.ListCount > 0 Then
            Cmb_Empleado_Curso.ListIndex = 0
        Else
            Cmb_Empleado_Curso.Text = ""
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
        
    End If
    
    
    

End Sub

Private Sub Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda, KeyCode)
End Sub

Private Sub Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda_Click()
Txt_No_Tarjeta_Indices_Asistencia.Visible = False
Lbl_No_Tarjeta_Indices_Asistencia.Visible = False

    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Caption = "Empleado"
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda, 0, "Apellido_Paterno")
        Txt_No_Tarjeta_Indices_Asistencia.Visible = True
        Lbl_No_Tarjeta_Indices_Asistencia.Visible = True
 
    ElseIf Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Caption = "Curso"
        Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID, Nombre", "Cat_Cursos_Capacitaciones", Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda, 0, "Nombre")
     ElseIf Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Caption = "Departamento"
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda, 0, "Nombre")
   Else
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Clear
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Visible = False
        Lbl_Rpt_Cursos_Indice_Asistencia_Busqueda.Visible = False
    End If
    If Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListCount > 0 Then
        Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)

End Sub

Private Sub Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empleado_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then
'            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, 0, True, "TODOS")
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%') AND SUpervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
'            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 1 Then
'                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 1
'            End If
'        Else
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%')", False, "")
'            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 1 Then
'                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 1
'            End If
'        End If
'    Else
'        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
'    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
'    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_Click()
'Dim Rs_Empleados_Supervisor As rdoResultset
'    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > -1 Then
'        If Trim(Empleado_Supervisor_ID) = "" Then
'            'Consulta Supervisor.
'            Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
'            Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
'            Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
'            Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'            Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
'            Mi_SQL = Mi_SQL & " AND Tipo='S'"
'            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
'            Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
'            Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'            Cmb_Rpt_Asistencia_Empleados_Supervisor.Clear
'            While Not Rs_Empleados_Supervisor.EOF
'                Cmb_Rpt_Asistencia_Empleados_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
'                Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
'                Rs_Empleados_Supervisor.MoveNext
'            Wend
'            Rs_Empleados_Supervisor.Close
''            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'", True, "TODOS")
'                Cmb_Rpt_Asistencia_Empleados_Supervisor.Enabled = True
'        Else
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Empleado_ID='" & Empleado_Supervisor_ID & "'")
'            Cmb_Rpt_Asistencia_Empleados_Supervisor.Enabled = False
'        End If
'        If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListCount > 0 Then Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex = 0
'    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Asistencia_Empleados_Empresa, 1, "Nombre", True, "TODAS")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_Click()
'Dim Rs_Empleados_Supervisor As rdoResultset
'
'    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Or Trim(Empleado_Supervisor_ID) <> "" Then
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
'        If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
'            Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
'        End If
'    Else
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_paterno", "", True, "TODOS")
'        If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
'            Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
'        End If
'    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_KeyPress(KeyAscii As Integer)
'Dim Rs_Empleados_Supervisor As rdoResultset
'    If KeyAscii = 13 Then
'    'Consulta Supervisor.
'        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
'        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
'        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
'        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
'        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
'        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
'        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%'"
'        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%')"
'        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
'        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'        Cmb_Rpt_Asistencia_Empleados_Supervisor.Clear
'        While Not Rs_Empleados_Supervisor.EOF
'            Cmb_Rpt_Asistencia_Empleados_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
'            Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
'            Rs_Empleados_Supervisor.MoveNext
'        Wend
'        Rs_Empleados_Supervisor.Close
'
''        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR " & _
''             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%')", True, "TODOS")
'        If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListCount > 1 Then
'            Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex = 1
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
'            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
'                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
'            End If
'        End If
'    Else
'        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
'    End If
'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%')", False, "")
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
'    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Supervisor, KeyCode)
End Sub

Private Sub Cmb_Rpt_Cursos_Hora_Hombre_Busqueda_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Cursos_Hora_Hombre_Busqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda, KeyCode)
End Sub

Private Sub Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda_Click()
Lbl_No_Tarjeta_Cursos_Horas_Hombre.Visible = False
Txt_No_Tarjeta_Cursos_Hioras_Hombre.Visible = False
     If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Caption = "Tipo de Empleado"
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Clear
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.AddItem "Todos"
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.AddItem "Sindicalizado"
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.AddItem "Confianza"
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Cursos_Hora_Hombre_Busqueda, 0, "Apellido_Paterno")
   ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 1 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Caption = "Empleado"
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Cursos_Hora_Hombre_Busqueda, 0, "Apellido_Paterno")
        Lbl_No_Tarjeta_Cursos_Horas_Hombre.Visible = True
        Txt_No_Tarjeta_Cursos_Hioras_Hombre.Visible = True
       
    ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 2 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Caption = "Curso"
        Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID, Nombre", "Cat_Cursos_Capacitaciones", Cmb_Rpt_Cursos_Hora_Hombre_Busqueda, 0, "Nombre", "", True, "")
    ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 3 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = True
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Caption = "Departamento"
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Cursos_Hora_Hombre_Busqueda, 0, "Nombre", "", True, "")
    Else
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Clear
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = False
        Lbl_Rpt_Cursos_Hora_Hombre_Busqueda.Visible = False
    End If
    If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListCount > 0 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda, KeyCode)
End Sub

Private Sub Cmb_Rpt_Cursos_Resumen_Mensual_Auditable_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Cursos_Resumen_Mensual_Auditable_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Cursos_Resumen_Mensual_Auditable, KeyCode)
End Sub

Private Sub Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleado_Accesos_Almacenes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If IsNumeric(Cmb_Rpt_Empleado_Accesos_Almacenes.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Empleado_Accesos_Almacenes, 1, "Nombre", "AND Estatus='A'")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Empleado_Accesos_Almacenes, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Empleado_Accesos_Almacenes.Text) & "%' OR " & "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Empleado_Accesos_Almacenes.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Empleado_Accesos_Almacenes.Text) & "%')", False, "")
          
        End If
        If Cmb_Rpt_Empleado_Accesos_Almacenes.ListCount > 0 Then
            Cmb_Rpt_Empleado_Accesos_Almacenes.ListIndex = 0
        Else
            Cmb_Rpt_Empleado_Accesos_Almacenes.Text = ""
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
        
    End If
End Sub
Private Sub Cmb_Rpt_Empleado_Accesos_Almacenes_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleado_Accesos_Almacenes, KeyCode)
End Sub
Private Sub Cmb_Rpt_Empleados_Alta_Departamento_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Alta_Departamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Alta_Departamento, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_Alta_Empresa_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Alta_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Alta_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_Alta_Puesto_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Alta_Puesto_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Baja_Departamento, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Departamento_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Departamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Baja_Departamento, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Empresa_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Baja_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Puesto_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_Empleados_Baja_Puesto_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Empleados_Baja_Puesto, KeyCode)
End Sub

Private Sub Cmb_Rpt_Empleados_No_Validados_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Supervisor As rdoResultset
    If KeyAscii = 13 Then
    
        'Consulta Supervisor.
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Rpt_Empleados_No_Validados_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Rpt_Empleados_No_Validados_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Rpt_Empleados_No_Validados_Supervisor.ItemData(Cmb_Rpt_Empleados_No_Validados_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Empleados_No_Validados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'AND (Nombre like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Empleados_No_Validados_Supervisor.Text) & "%')", True, "TODOS")
        If Cmb_Rpt_Empleados_No_Validados_Supervisor.ListCount > 1 Then
            Cmb_Rpt_Empleados_No_Validados_Supervisor.ListIndex = 1
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_General_Cursos_Institucion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_General_Cursos_Institucion_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_General_Cursos_Institucion, KeyCode)
End Sub

Private Sub Cmb_Rpt_General_Cursos_Instructor_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_General_Cursos_Instructor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_General_Cursos_Instructor, KeyCode)
End Sub

Private Sub Cmb_Rpt_General_Cursos_Sala_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Rpt_General_Cursos_Sala_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_General_Cursos_Sala, KeyCode)
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex > 0 Then
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, 0, True, "TODOS")
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%') AND SUpervisor_ID = '" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListCount > 1 Then
                Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex = 1
            End If
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.Text) & "%')", True, "TODOS")
            If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListCount > 1 Then
                Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex = 1
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Faltas_Retardos_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Empresa_Click()
Dim Rs_Empleados_Supervisor As rdoResultset
    If Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListIndex > -1 Then
        If Trim(Empleado_Supervisor_ID) = "" Then
            'Consulta Supervisor.
            Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
            Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
            Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
            Mi_SQL = Mi_SQL & " AND Tipo='S'"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
            Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Clear
            While Not Rs_Empleados_Supervisor.EOF
                Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
                Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
                Rs_Empleados_Supervisor.MoveNext
            Wend
            Rs_Empleados_Supervisor.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'", True, "TODOS")
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Enabled = True
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Empleado_ID='" & Empleado_Supervisor_ID & "'")
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Enabled = False
        End If
        If Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListCount > 0 Then Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Historico_Faltas_Retardos_Empresa, 1, "Nombre", True, "TODAS")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Faltas_Retardos_Empresa, KeyCode)
End Sub


Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Supervisor_Click()
     If Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex > 0 Or Trim(Empleado_Supervisor_ID) <> "" Then
       
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID = '" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
        If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListCount > 0 Then
            Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex = 0
        End If
    Else
        
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 0, "Apellido_paterno", , True, "TODOS")
        If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListCount > 0 Then
            Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex = 0
        End If
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Supervisor As rdoResultset
    If KeyAscii = 13 Then
        'Consulta Supervisor.
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Faltas_Retardos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.Text) & "%')", True, "TODOS")
        If Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListCount > 1 Then
            Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex = 1
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 0, 0, True, "TODOS")
            'If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 0 Then
            '    Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 0
            'End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Faltas_Retardos_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Faltas_Retardos_Supervisor, KeyCode)
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex > 0 Then
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, 0, True, "TODOS")
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%') AND SUpervisor_ID = '" & Format(Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 1 Then
                Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 1
            End If
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Empleado.Text) & "%')", True, "TODOS")
            If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 1 Then
                Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 1
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Permisos_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Empresa_Click()
Dim Rs_Empleados_Supervisor As rdoResultset
    If Cmb_Rpt_Historico_Permisos_Empresa.ListIndex > -1 Then
        If Trim(Empleado_Supervisor_ID) = "" Then
            'Consulta Supervisor.
            Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
            Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
            Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
            Mi_SQL = Mi_SQL & " AND Tipo='S'"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
            Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Historico_Permisos_Supervisor.Clear
            While Not Rs_Empleados_Supervisor.EOF
                Cmb_Rpt_Historico_Permisos_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
                Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
                Rs_Empleados_Supervisor.MoveNext
            Wend
            Rs_Empleados_Supervisor.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'", True, "TODOS")
            Cmb_Rpt_Historico_Permisos_Supervisor.Enabled = True
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Empleado_ID='" & Empleado_Supervisor_ID & "'")
            Cmb_Rpt_Historico_Permisos_Supervisor.Enabled = False
        End If
        If Cmb_Rpt_Historico_Permisos_Supervisor.ListCount > 0 Then Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Historico_Permisos_Empresa, 1, "Nombre", True, "TODAS")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Permisos_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Supervisor_Click()
    If Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex > 0 Or Trim(Empleado_Supervisor_ID) <> "" Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID = '" & Format(Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
        If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 0 Then
            Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 0
        End If
    Else
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Historico_Permisos_Empleado, 0, "Apellido_paterno", , True, "TODOS")
        If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 0 Then
            Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 0
        End If
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Supervisor As rdoResultset
    If KeyAscii = 13 Then
        'Consulta Supervisor.
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Rpt_Historico_Permisos_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Rpt_Historico_Permisos_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'AND (Nombre like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%' OR " & _
'             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Historico_Permisos_Supervisor.Text) & "%')", True, "TODOS")
        If Cmb_Rpt_Historico_Permisos_Supervisor.ListCount > 1 Then
            Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex = 1
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Historico_Permisos_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID = '" & Format(Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 0 Then
                Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 0
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Historico_Permisos_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Historico_Permisos_Supervisor, KeyCode)
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex > 0 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%') AND SUpervisor_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
                Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
            End If
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, 1, "Apellido_Paterno", "AND Estatus='A' AND (Nombre like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.Text) & "%')", True, "TODOS")
            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 1 Then
                Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 1
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa_Click()
Dim Rs_Empleados_Supervisor As rdoResultset

    If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex > -1 Then
        If Trim(Empleado_Supervisor_ID) = "" Then
            'Consulta Supervisor.
            Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
            Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
            Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
            Mi_SQL = Mi_SQL & " AND Tipo='S'"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
            Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Clear
            While Not Rs_Empleados_Supervisor.EOF
                Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
                Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
                Rs_Empleados_Supervisor.MoveNext
            Wend
            Rs_Empleados_Supervisor.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'", True, "TODOS")
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Enabled = True
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Empleado_ID='" & Empleado_Supervisor_ID & "'")
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Enabled = False
        End If
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListCount > 0 Then Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa, 1, "Nombre", True, "TODAS")
       If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex > 1 Then
            Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex = 1
       End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo_Click()
Dim Dias_Mes As Integer

    Select Case Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text
        Case "ACUMULADO":
            
        Case "SEMANAL":
            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value = DateAdd("d", 6, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value)
        Case "MENSUAL":
            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value = DateAdd("d", 30, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value)
    End Select
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor_Click()

    If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex > 0 Or Trim(Empleado_Supervisor_ID) <> "" Then
        
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, 1, "Nombre", "AND Supervisor_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
            Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
        End If
    Else
        
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, 0, "Apellido_paterno", , True, "TODOS")
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
            Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
        End If
    End If
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Supervisor As rdoResultset
    If KeyAscii = 13 Then
        'Consulta Supervisor.
        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
        Mi_SQL = Mi_SQL & " AND (Nombre like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Paterno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%')"
        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Clear
        While Not Rs_Empleados_Supervisor.EOF
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
            Rs_Empleados_Supervisor.MoveNext
        Wend
        Rs_Empleados_Supervisor.Close
'        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S'AND (Nombre like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%' OR " & _
'        "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.Text) & "%')", True, "TODOS")
        
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListCount > 1 Then
            Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex = 1
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado, 0, 0, True, "TODOS")
            'If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
            '    Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
            'End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor, KeyCode)
End Sub

Public Sub Cmb_Turno_Faltas_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio_Change()
    Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo_Click
End Sub

Private Sub Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio_Click()
    Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo_Click
End Sub

Private Sub Form_Load()
    Me.Width = 5180
    Me.Height = 3650
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
    Me.Top = 100
    Pic_Reportes.Visible = True
    Archivo_Reporte_Abierto = False
End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Open Ruta_Temporal & Reporte & ".txt" For Output As #1
    Open Ruta_Temporal & Reporte & "xls.txt" For Output As #2 'Reporte a xls
    Archivo_Reporte_Abierto = True
    Print #1,
    Print #2,
    Print #1, Conectar_Ayudante.Centrar_Texto(Empresa, 120)
    Print #2, "||"; Empresa
    Print #1,
    Print #2,
    Print #1, Titulo; Conectar_Ayudante.Alinea_Derecha(Format(Now, "dd MMM yyyy"), 119 - Len(Titulo))
    Print #2, "||" & Titulo; "|||||"; Format(Now, "dd MMM yyyy")
    Print #1,
    Print #2,
    If DateDiff("s", Format(Fecha_Inicial, "HH:mm:ss"), "00:00:00") <> 0 And DateDiff("s", Format(Fecha_Termino, "HH:mm:ss"), "00:00:00") <> 0 Then
        If Solo_mes Then
            Print #1, "DE "; Format(Fecha_Inicial, "MMMM yyyy")
            Print #2, "|DE|"; Format(Fecha_Inicial, "MMMM yyyy")
        Else
            Print #1, "DE "; Format(Fecha_Inicial, "dd MMMM yyyy") & " A "; Format(Fecha_Termino, "dd MMMM yyyy")
            Print #2, "|DE|"; Format(Fecha_Inicial, "dd MMMM yyyy") & "|A|"; Format(Fecha_Termino, "dd MMMM yyyy")
        End If
    End If
    Print #1,
    Print #2,
    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
    Print #2, "--------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub Encabezado_Reporte_Excel(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Archivo_Reporte_Abierto = True
    Print #2,
    Print #2, "||"; Empresa
    Print #2,
    Print #2, "||" & Titulo; "|||||"; Format(Now, "dd MMM yyyy")
    Print #2,
    If DateDiff("s", Format(Fecha_Inicial, "HH:mm:ss"), "00:00:00") <> 0 And DateDiff("s", Format(Fecha_Termino, "HH:mm:ss"), "00:00:00") <> 0 Then
        If Solo_mes Then
            Print #2, "|DE|"; Format(Fecha_Inicial, "MMMM yyyy")
        Else
            Print #2, "|DE|"; Format(Fecha_Inicial, "dd MMMM yyyy") & "|A|"; Format(Fecha_Termino, "dd MMMM yyyy")
        End If
    End If
    Print #2,
    Print #2, "--------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub Encabezado_Reporte_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Archivo_Reporte_Abierto = True
    Print #1,
    Print #1, Conectar_Ayudante.Centrar_Texto(Empresa, 120)
    Print #1,
    Print #1, Titulo; Conectar_Ayudante.Alinea_Derecha(Format(Now, "dd MMM yyyy"), 119 - Len(Titulo))
    Print #1,
    If DateDiff("s", Format(Fecha_Inicial, "HH:mm:ss"), "00:00:00") <> 0 And DateDiff("s", Format(Fecha_Termino, "HH:mm:ss"), "00:00:00") <> 0 Then
        If Solo_mes Then
            Print #1, "DE "; Format(Fecha_Inicial, "MMMM yyyy")
        Else
            Print #1, "DE "; Format(Fecha_Inicial, "dd MMMM yyyy") & " A "; Format(Fecha_Termino, "dd MMMM yyyy")
        End If
    End If
    Print #1,
    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
    
    
    Print #2,
    Print #2, Conectar_Ayudante.Centrar_Texto(Empresa, 120)
    Print #2,
    Print #2, Titulo; Conectar_Ayudante.Alinea_Derecha(Format(Now, "dd MMM yyyy"), 119 - Len(Titulo))
    Print #2,
    If DateDiff("s", Format(Fecha_Inicial, "HH:mm:ss"), "00:00:00") <> 0 And DateDiff("s", Format(Fecha_Termino, "HH:mm:ss"), "00:00:00") <> 0 Then
        If Solo_mes Then
            Print #2, "DE "; Format(Fecha_Inicial, "MMMM yyyy")
        Else
            Print #2, "DE "; Format(Fecha_Inicial, "dd MMMM yyyy") & " A "; Format(Fecha_Termino, "dd MMMM yyyy")
        End If
    End If
    Print #2,
    Print #2, "--------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub Finalizar_Reporte(Abrir As Boolean)
    Close #1, #2
    Archivo_Reporte_Abierto = False
    If Abrir Then
        Rich_Reporte.Font = "Courier New"
        'If (Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "MENSUAL" And Reporte = "Horas_Trabajadas_Empleado") Or Reporte = "Asistencias_Empleados" Then
        If (Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "MENSUAL" And Reporte = "Horas_Trabajadas_Empleado") Then
           Rich_Reporte.Font.Size = 7
        Else
            Rich_Reporte.Font.Size = 8
        End If
        Rich_Reporte.FileName = Ruta_Temporal & Reporte & ".txt"
        Rich_Reporte.Visible = True
        Pic_Reportes.ZOrder 1
        'Orientacion_Reporte = Orientacion.Vertical
        Me.Width = 13305
        Me.Height = 8340
        Me.Left = 0
        Me.Top = 0
    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Horas_Trabajadas_Empleado
'DESCRIPCION: Genera el reporte de horas trabajadas por empleado
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 28-Febrero-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Horas_Trabajadas_Empleado()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset     'Informacion de los tiempo muertos
Dim Mi_SQL As String                                                'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                                            'ID de la empresa
Dim Empleado_ID_Reporte As String                                            'ID de la empleado
Dim Horas_Empresa As Double                                         'Horas total por empresa
Dim Horas_Extra_Empresa As Double                                   'Horas exta total por empresa
Dim Horas_Total As Double                                           'Horas total
Dim Horas_Extra_Total As Double                                     'Horas extra
Dim Supervisor_ID As String                                         'Supervisor de la consulta
Dim Departamento_ID As String                                         'Departamento de la consulta
Dim Nombre_Supervisor As String                                     'Nombre del supervisor

    'Mi_SQL = "SELECT SUM(ISNULL(AA.Horas_Aprobadas,0)) as Horas, "
    Mi_SQL = "SELECT (SUM(ISNULL(AA.Horas_Aprobadas,0)) - SUM(ISNULL(AA.Horas_Extra,0))) as Horas, "
    Mi_SQL = Mi_SQL & " SUM(ISNULL(AA.Horas_Extra,0)) as Horas_Extras, "
    Mi_SQL = Mi_SQL & " SUM(ISNULL(AA.Horas_Aprobadas,0)) as Horas_Total,"
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre,"
    Mi_SQL = Mi_SQL & " CEM.Empresa_ID,CEM.Nombre as Nombre_Empresa,"
    Mi_SQL = Mi_SQL & " ISNULL(CE.Supervisor_ID,'N') as Supervisor_ID,"
    Mi_SQL = Mi_SQL & " ISNULL(CD.Nombre,'') as Departamento, CE.Departamento_ID"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA, Cat_Empleados CE, Cat_Empresas CEM, Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.EMpresa_ID = CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = CD.Departamento_ID"
    'Validacion de Empresa
    If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AA.Empleado_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex), "00000") & "'"
    End If
    'Validacion del departamento
    If Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND AA.Fecha >=" & Par_Fecha & Format(Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND AA.Fecha <=" & Par_Fecha & Format(Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CD.Nombre,CE.Departamento_ID,CE.Supervisor_ID,CEM.Nombre,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CD.Nombre,CE.Departamento_ID,CE.Supervisor_ID,CE.Apellido_Paterno"
    Empresa_ID_Reporte = ""
    Empleado_ID_Reporte = ""
    'Ejecuta la consulta
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            Horas_Total = 0
            Horas_Extra_Total = 0
            Supervisor_ID = ""
            Nombre_Supervisor = ""
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE HORAS TRABAJADAS POR EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value))
            'Agrega el encabezado para la informacion
            While Not .EOF
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Horas_Empresa = 0
                    Horas_Extra_Empresa = 0
                    Supervisor_ID = ""
                    Print #1, " "; .rdoColumns("Nombre_Empresa")
                    Print #2, .rdoColumns("Nombre_Empresa")
                    Print #1,
                    Print #2,
                    Print #1, " Empleado                                          Horas Trabajadas        Horas Extra            Total Horas"
                    Print #2, "Empleado|||Horas Trabajadas|Horas Extra|Total Horas"
                End If
                If .rdoColumns("Departamento_ID") <> Departamento_ID Then
                    Departamento_ID = .rdoColumns("Departamento_ID")
                    Print #1,
                    Print #2,
                    Print #1, "Departamento: "; .rdoColumns("Departamento")
                    Print #2, "Departamento:|"; .rdoColumns("Departamento")
                End If
                If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                    Supervisor_ID = .rdoColumns("Supervisor_ID")
                    Print #1,
                    Print #2,
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                    Print #1, "Supervisor: "; Nombre_Supervisor
                    Print #2, "Supervisor:|"; Nombre_Supervisor
                End If
                Print #1, " "; Mid(.rdoColumns("Nombre"), 1, 50); Spc(50 - Len(Mid(.rdoColumns("Nombre"), 1, 50))); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("Horas"), 16); _
                          Conectar_Ayudante.Alinea_Derecha(.rdoColumns("Horas_Extras"), 19); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("Horas_Total"), 23)
                Print #2, .rdoColumns("Nombre"); "|||"; .rdoColumns("Horas"); "|"; .rdoColumns("Horas_Extras"); "|"; .rdoColumns("Horas_Total")
                Horas_Empresa = Horas_Empresa + Val(.rdoColumns("Horas"))
                Horas_Extra_Empresa = Horas_Extra_Empresa + Val(.rdoColumns("Horas_Extras"))
                Horas_Total = Horas_Total + Val(.rdoColumns("Horas"))
                Horas_Extra_Total = Horas_Extra_Total + Val(.rdoColumns("Horas_Extras"))
                .MoveNext
                If Not .EOF Then
                    'Imprime el total por empresa
                    If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                        Print #1, Spc(38); "Total Empresa"; Conectar_Ayudante.Alinea_Derecha(Format(Horas_Empresa, "#0"), 16); Conectar_Ayudante.Alinea_Derecha(Format(Horas_Extra_Empresa, "#0"), 19); _
                                  Conectar_Ayudante.Alinea_Derecha(Format(Horas_Empresa + Horas_Extra_Empresa, "#0"), 23)
                        Print #2, "||Total|"; Format(Horas_Empresa, "#0"); "|"; Format(Horas_Extra_Empresa, "#0"); "|"; Format(Horas_Empresa + Horas_Extra_Empresa, "#0")
                        Horas_Empresa = 0
                        Horas_Extra_Empresa = 0
                    End If
                Else
                    'Imprime el total por empresa
                    Print #1, Spc(38); "Total Empresa"; Conectar_Ayudante.Alinea_Derecha(Format(Horas_Empresa, "#0"), 16); Conectar_Ayudante.Alinea_Derecha(Format(Horas_Extra_Empresa, "#0"), 19); _
                                  Conectar_Ayudante.Alinea_Derecha(Format(Horas_Empresa + Horas_Extra_Empresa, "#0"), 23)
                        Print #2, "||Total|"; Format(Horas_Empresa, "#0"); "|"; Format(Horas_Extra_Empresa, "#0"); "|"; Format(Horas_Empresa + Horas_Extra_Empresa, "#0")
                    'Imprime el total General
                    Print #1,
                    Print #2,
                    Print #1, Spc(38); "Total General"; Conectar_Ayudante.Alinea_Derecha(Format(Horas_Total, "#0"), 16); Conectar_Ayudante.Alinea_Derecha(Format(Horas_Extra_Total, "#0"), 19); _
                              Conectar_Ayudante.Alinea_Derecha(Format(Horas_Total + Horas_Extra_Total, "#0"), 23)
                    Print #2, "||Total General|"; Format(Horas_Total, "#0"); "|"; Format(Horas_Extra_Total, "#0"); "|"; Format(Horas_Total + Horas_Extra_Total, "#0")
                End If
            Wend
            .Close
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado", Me)
        Else
            MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
        End If
    End With
    Set Rs_Consulta_Adm_Asistencias = Nothing
    'Haya o no haya registros se cambia el Puntero del Mouse
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Public Sub Inicializar()
 Dim Rs_Empleados_Departamento As rdoResultset
Dim Rs_Empleados_Supervisor As rdoResultset

    Btn_Imprimir.Enabled = False
    Btn_Exportar.Enabled = False
    Btn_Regresar.Enabled = False
    Btn_Salir.Enabled = False
    
    Select Case Reporte
        Case "Horas_Trabajadas_Empleado":
            'Carga las empresas
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListCount > 0 Then
                Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex = 0
            End If
            'Departamentos
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento, 0, "Nombre", , True, "TODOS")
            
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.Text = "<-SELECCIONE->"
            
            If Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListCount > 0 Then
                Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex = 0
            End If
            If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListCount > 0 Then
                Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex = 0
            End If
            Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.ListIndex = 0
            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value = Now
            Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value = Now
            Btn_Rpt_Generar.Top = 3000
            Btn_Salir_Reporte.Top = 3000
            Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.SetFocus
                    
        Case "Historico_Faltas_Retardos":
            'Puestos
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Historico_Faltas_Retardos_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListCount > 0 Then
                Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListIndex = 0
            End If
            'Departamentos
            
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Historico_Faltas_Retardos_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Historico_Faltas_Retardos_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Historico_Faltas_Retardos_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ListCount > 0 Then
                Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ListIndex = 0
            End If
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Empresa_ID = '" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListIndex), "00000") & "'", Cmb_Rpt_Historico_Faltas_Retardos_Empleado, 0, 0, True, "TODOS")
            'If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListCount > 0 Then
            '    Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex = 0
            'End If
            Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Inicio.Value = Now
            Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Termino.Value = Now
            Cmb_Rpt_Historico_Faltas_Retardos_Empresa.SetFocus
            
        Case "Historico_Permisos":
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Historico_Permisos_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Historico_Permisos_Empresa.ListCount > 0 Then
                Cmb_Rpt_Historico_Permisos_Empresa.ListIndex = 0
            End If
            'Departamentos
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Historico_Permisos_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Historico_Permisos_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Historico_Permisos_Departamento.ItemData(Cmb_Rpt_Historico_Permisos_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Historico_Permisos_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Historico_Permisos_Departamento.ListCount > 0 Then
                Cmb_Rpt_Historico_Permisos_Departamento.ListIndex = 0
            End If
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Empresa_ID = '" & Format(Cmb_Rpt_Historico_Permisos_Empresa.ItemData(Cmb_Rpt_Historico_Permisos_Empresa.ListIndex), "00000") & "'", Cmb_Rpt_Historico_Permisos_Empleado, 0, 0, True, "TODOS")
            'If Cmb_Rpt_Historico_Permisos_Empleado.ListCount > 0 Then
            '    Cmb_Rpt_Historico_Permisos_Empleado.ListIndex = 0
            'End If
            Dtp_Rpt_Historico_Permisos_Fecha_inicio.Value = Now
            Dtp_Rpt_Historico_Permisos_Fecha_Termino.Value = Now
            Cmb_Rpt_Historico_Permisos_Empresa.SetFocus
            
        Case "Asistencias_Empleados"
            'Carga la empresa
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Asistencia_Empleados_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Asistencia_Empleados_Empresa.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex = 0
            End If
'            'Departamentos
'            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
'            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
'            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'            Mi_SQL = Mi_SQL & " ORDER BY Cat_Departamentos.Nombre"
            Mi_SQL = "SELECT * FROM Cat_Departamentos "
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Asistencia_Empleados_Departamento.Clear
            Cmb_Rpt_Asistencia_Empleados_Departamento.AddItem "TODOS"
        
            Cmb_Rpt_Asistencia_Empleados_Departamento.ItemData(Cmb_Rpt_Asistencia_Empleados_Departamento.NewIndex) = 0
        
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Asistencia_Empleados_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Asistencia_Empleados_Departamento.ItemData(Cmb_Rpt_Asistencia_Empleados_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Asistencia_Empleados_Departamento.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Departamento.ListIndex = 0
            End If
'           Supervisor
           Mi_SQL = "SELECT Empleado_ID ,(Apellido_Paterno + ' ' + Apellido_Materno + ' ' + Nombre) AS Nombre "
           Mi_SQL = Mi_SQL & " From Cat_Empleados WHERE Tipo = 'S' AND Estatus = 'A'  ORDER BY Apellido_Paterno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Asistencia_Empleados_Supervisor.Clear
            Cmb_Rpt_Asistencia_Empleados_Supervisor.AddItem "TODOS"
        
            Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.NewIndex) = 0
        
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Asistencia_Empleados_Supervisor.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex = 0
            End If
'           Empleado
           Mi_SQL = "select Empleado_ID, ISNULL(Nombre, '') + ' ' + ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') as Nombre"
           Mi_SQL = Mi_SQL & " From Cat_Empleados where Estatus = 'A' ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Asistencia_Empleados_Empleado.Clear
            Cmb_Rpt_Asistencia_Empleados_Empleado.AddItem "TODOS"
        
            Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.NewIndex) = 0
        
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Asistencia_Empleados_Empleado.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
            End If
        
            'Fechas
            Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value = Now
            Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value = Now
            Cmb_Rpt_Asistencia_Empleados_Empresa.SetFocus
        
        Case "Empleados_No_Validados"
            If Trim(Empleado_Supervisor_ID) = "" Then
                'Consulta Supervisor.
                Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
                Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
                Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
                Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
                Mi_SQL = Mi_SQL & " AND Tipo='S'"
                Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
                Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
                Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Rpt_Empleados_No_Validados_Supervisor.Clear
                While Not Rs_Empleados_Supervisor.EOF
                    Cmb_Rpt_Empleados_No_Validados_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
                    Cmb_Rpt_Empleados_No_Validados_Supervisor.ItemData(Cmb_Rpt_Empleados_No_Validados_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
                    Rs_Empleados_Supervisor.MoveNext
                Wend
                Rs_Empleados_Supervisor.Close
'                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S'", Cmb_Rpt_Empleados_No_Validados_Supervisor, 0, "Apellido_Paterno", , True, "TODOS")
                Cmb_Rpt_Empleados_No_Validados_Supervisor.Enabled = True
            Else
                'Consulta Supervisor.
                Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
                Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
                Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
                Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
                Mi_SQL = Mi_SQL & " AND Tipo='S'"
                Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
                Mi_SQL = Mi_SQL & " AND Empleado_ID='" & Empleado_Supervisor_ID & "'"
                Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
                Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Rpt_Empleados_No_Validados_Supervisor.Clear
                While Not Rs_Empleados_Supervisor.EOF
                    Cmb_Rpt_Empleados_No_Validados_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
                    Cmb_Rpt_Empleados_No_Validados_Supervisor.ItemData(Cmb_Rpt_Empleados_No_Validados_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
                    Rs_Empleados_Supervisor.MoveNext
                Wend
                Rs_Empleados_Supervisor.Close
'                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Empleado_ID='" & Empleado_Supervisor_ID & "'", Cmb_Rpt_Empleados_No_Validados_Supervisor, 0, "Apellido_Paterno")
                Cmb_Rpt_Empleados_No_Validados_Supervisor.Enabled = False
            End If
            If Cmb_Rpt_Empleados_No_Validados_Supervisor.ListCount > 0 Then
                Cmb_Rpt_Empleados_No_Validados_Supervisor.ListIndex = 0
            End If
            'Departamentos
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Empleados_No_Validados_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Empleados_No_Validados_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Empleados_No_Validados_Departamento.ItemData(Cmb_Rpt_Empleados_No_Validados_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Empleados_No_Validados_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_No_Validados_Departamento.ListCount > 0 Then
                Cmb_Rpt_Empleados_No_Validados_Departamento.ListIndex = 0
            End If
            Dtp_Rpt_Empleados_No_Validados_Fecha_Inicio.Value = Now
            Dtp_Rpt_Empleados_No_Validados_Fecha_Termino.Value = Now
        
        Case "Empleados_Baja"
            'Carga las empresas
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Empleados_Baja_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Empleados_Baja_Empresa.ListCount > 0 Then
                Cmb_Rpt_Empleados_Baja_Empresa.ListIndex = 0
            End If
            'Departamentos
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Empleados_Baja_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Empleados_Baja_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Empleados_Baja_Departamento.ItemData(Cmb_Rpt_Empleados_Baja_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Empleados_Baja_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Baja_Departamento.ListCount > 0 Then
                Cmb_Rpt_Empleados_Baja_Departamento.ListIndex = 0
            End If
            'Puestos
            Call Conectar_Ayudante.Llena_Combo_Item("Puesto_ID, Nombre", "Cat_Puestos", Cmb_Rpt_Empleados_Baja_Puesto, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Baja_Puesto.ListCount > 0 Then
                Cmb_Rpt_Empleados_Baja_Puesto.ListIndex = 0
            End If
            Dtp_Rpt_Empleados_Baja_Fecha_Inicio.Value = Now
            Dtp_Rpt_Empleados_Baja_Fecha_Termino.Value = Now
            
        Case "Empleados_Alta"
            'Carga las empresas
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Empleados_Alta_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Empleados_Alta_Empresa.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Empresa.ListIndex = 0
            End If
            'Departamentos
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Empleados_Alta_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Empleados_Alta_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Empleados_Alta_Departamento.ItemData(Cmb_Rpt_Empleados_Alta_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Empleados_Alta_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Alta_Departamento.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Departamento.ListIndex = 0
            End If
            'Puestos
            Call Conectar_Ayudante.Llena_Combo_Item("Puesto_ID, Nombre", "Cat_Puestos", Cmb_Rpt_Empleados_Alta_Puesto, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Alta_Puesto.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Puesto.ListIndex = 0
            End If
            Dtp_Rpt_Empleados_Alta_Fecha_Inicio.Value = Now
            Dtp_Rpt_Empleados_Alta_Fecha_Termino.Value = Now
            
        Case "Empleados_Faltas":
            Dtp_Fecha_Faltas_Empleados.Value = Now
            Dtp_Fecha_Faltas_Empleados_Fin.Value = Now
        
        Case "Empleados_Faltas_Validadas":
            Dtp_Fecha_Faltas_Empleados.Value = Now
            Dtp_Fecha_Faltas_Empleados_Fin.Visible = False
            Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Turno_Faltas, 0, "Nombre", "", True, "<-SELECCIONE->")
            If Cmb_Turno_Faltas.ListCount > 0 Then Cmb_Turno_Faltas.ListIndex = 0 Else Cmb_Turno_Faltas.Text = ""
        
        Case "Reporte_Comedor"
            Lbl_Curso.Visible = False
            Cmb_Curso.Visible = False
            Dtp_Fecha_Inicio_Curso.Value = Now
            Dtp_Fecha_Fin_Curso.Value = Now
            Cmb_Estatus.ListIndex = 0
        
        Case "Empleados_Huella_Comedor"
            'Carga las empresas
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Empleados_Alta_Empresa, 0, "Nombre", , True, "TODAS")
            If Cmb_Rpt_Empleados_Alta_Empresa.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Empresa.ListIndex = 0
            End If
            'Departamentos
            'Consulta Departamento.
            Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
            Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Empleados_Alta_Departamento.Clear
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Empleados_Alta_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Empleados_Alta_Departamento.ItemData(Cmb_Rpt_Empleados_Alta_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
'            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Rpt_Empleados_Alta_Departamento, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Alta_Departamento.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Departamento.ListIndex = 0
            End If
            'Puestos
            Call Conectar_Ayudante.Llena_Combo_Item("Puesto_ID, Nombre", "Cat_Puestos", Cmb_Rpt_Empleados_Alta_Puesto, 0, "Nombre", , True, "TODOS")
            If Cmb_Rpt_Empleados_Alta_Puesto.ListCount > 0 Then
                Cmb_Rpt_Empleados_Alta_Puesto.ListIndex = 0
            End If
            Dtp_Rpt_Empleados_Alta_Fecha_Inicio.Value = Now
            Dtp_Rpt_Empleados_Alta_Fecha_Termino.Value = Now
            
        Case "Cursos_Por_Empleado"
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, No_Tarjeta, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado, 1, "Apellido_Paterno", "AND Estatus='A'", True, "TODOS")
            Call Llena_Combo_Consulta(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado, "SELECT Empleado_ID ,No_Tarjeta ,(Apellido_Paterno + ' ' + Apellido_Materno + ' ' + Nombre) AS Nombre FROM Cat_Empleados WHERE ((Apellido_Paterno LIKE '%%' or Apellido_Materno LIKE '%%' OR Nombre like '%%') or ((Apellido_Paterno + ' ' + Apellido_Materno + ' ' + Nombre) like '%%') or No_Tarjeta like '%%') AND Estatus = 'A' ORDER BY  Apellido_Paterno, Apellido_Materno, Nombre  ")
            
            If Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListCount > 0 Then
                Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex = 0
            End If
            Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Fin.Value = Now
            Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Inicio.Value = Now
            
        Case "Cursos_Hora_Hombre"
            Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0
            Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value = Now
            Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value = Now
            
        Case "Cursos_Indice_Asistencia"
            Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 0
            Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value = Now
            Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value = Now
        
        Case "Cursos_Resumen_Mensual"
            Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Curso_Id, Nombre", "Cat_Tipos_Cursos", Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso, 0, "Nombre", , True)
            If Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListCount > 0 Then
                Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListIndex = 0
            End If
            Cmb_Rpt_Cursos_Resumen_Mensual_Auditable.ListIndex = 2
            Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value = Now
            Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value = Now
            
        Case "Reporte_General_Cursos"
            Label41.Caption = "Tipo"
'            Call Conectar_Ayudante.Llena_Combo_Item("Instructor_Id, (Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno)", "Cat_Instructores", Cmb_Rpt_General_Cursos_Instructor, 0, "Nombre", , True)
'            Call Conectar_Ayudante.Llena_Combo_Item("Institucion_Id, Nombre", "Cat_Instituciones", Cmb_Rpt_General_Cursos_Institucion, 0, "Nombre", , True)
'            Call Conectar_Ayudante.Llena_Combo_Item("Sala_Id, Nombre", "Cat_Salas", Cmb_Rpt_General_Cursos_Sala, 0, "Nombre", , True)
            Cmb_Rpt_General_Cursos_Instructor.AddItem ("TODOS")
            Cmb_Rpt_General_Cursos_Instructor.AddItem ("AUDITABLE")
            Cmb_Rpt_General_Cursos_Instructor.AddItem ("NO AUDITALE")
            If Cmb_Rpt_General_Cursos_Instructor.ListCount > 0 Then
                Cmb_Rpt_General_Cursos_Instructor.ListIndex = 0
            End If
            Cmb_Rpt_General_Cursos_Institucion.Visible = False
            Cmb_Rpt_General_Cursos_Sala.Visible = False
            Label42.Visible = False
            Label43.Visible = False
'            If Cmb_Rpt_General_Cursos_Institucion.ListCount > 0 Then
'                Cmb_Rpt_General_Cursos_Institucion.ListIndex = 0
'            End If
'            If Cmb_Rpt_General_Cursos_Sala.ListCount > 0 Then
'                Cmb_Rpt_General_Cursos_Sala.ListIndex = 0
'            End If
            Dtp_Rpt_Genera_Cursosl_Fecha_Inicio.Value = Now
            Dtp_Rpt_General_Cursos_Fecha_Fin.Value = Now
            
         'Almacenes Accesos
         Case "Accesos_Almacenes":
'            'Empleados
'            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, Nombre", "Cat_Empleados", Cmb_Rpt_Empleado_Accesos_Almacenes, 0, "Nombre", , True, "TODOS")
'            If Cmb_Rpt_Empleado_Accesos_Almacenes.ListCount > 0 Then
'                Cmb_Rpt_Empleado_Accesos_Almacenes.ListIndex = 0
'            End If
            Dtp_Rpt_Accesos_Almacenes_Fecha_Inicio.Value = Now
            Dtp_Rpt_Accesos_Almacenes_Fecha_Termino.Value = Now
    End Select
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Historico_Faltas_Retardos
'DESCRIPCION: Genera el reporte de las faltas y retardos que el empleado ha tenido
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 25-Mayo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Historico_Faltas_Retardos()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset         'Información de inasistencias o faltas
Dim Rs_Consulta_Adm_Inasistencias As rdoResultset       'Información de detalle de inasistencias o faltas
Dim Mi_SQL As String                                    'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                        'ID de la empresa
Dim Empleado_ID_Reporte As String                       'ID del empleado
Dim Concepto As String                                  'Registra la informacion de Falta, Falta Registrada, Retardo
Dim Observaciones As String                              'Detalle de la falta
Dim Supervisor_ID  As String
Dim Departamento_ID_Reporte As String
Dim Nombre_Supervisor As String
Dim Descripcion As String
Dim Usuario_Creo As String

    'Consulta los registros de faltas y retardos
    Mi_SQL = "SELECT AA.Simbologia, AA.SubSimbologia, AA.Fecha, ISNULL(AA.Tiempo_Retardo,0) as Tiempo_Retardo,"
    Mi_SQL = Mi_SQL & " ISNULL(CE.Retardos,0) as Retardos, CE.Fecha_Retardo, AA.Empleado_ID"
    Mi_SQL = Mi_SQL & " ,CE.No_Tarjeta,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre,"
    Mi_SQL = Mi_SQL & " CEM.Empresa_ID,CEM.Nombre as Nombre_Empresa,"
    Mi_SQL = Mi_SQL & " ISNULL(CE.Supervisor_ID,'N') as Supervisor_ID,AA.Usuario_Creo,"
    Mi_SQL = Mi_SQL & " ISNULL(CD.Nombre,'') as Departamento, CE.Departamento_ID"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA, Cat_Empleados CE, Cat_Empresas CEM, Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = CD.Departamento_ID"
    'Mi_SQL = Mi_SQL & " AND (AA.Simbologia='F' OR AA.Simbologia='RE' OR AA.Simbologia='FI' OR AA.Simbologia='FS')"
    Mi_SQL = Mi_SQL & " AND AA.Simbologia IN ('F','RE')"
    'Validacion de Empresa
    If Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CEM.Empresa_ID='" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AA.Empleado_ID='" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Empleado.ListIndex), "00000") & "'"
    End If
    'Validacion del departamento
    If Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ItemData(Cmb_Rpt_Historico_Faltas_Retardos_Departamento.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND AA.Fecha>='" & Format(Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND AA.Fecha<='" & Format(Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    'Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CE.Supervisor_ID,CEM.Nombre,AA.Movimiento,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AA.Empleado_ID,AA.Fecha,AA.Tiempo_Retardo"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CE.Supervisor_ID,AA.Empleado_ID,AA.Fecha,CE.Apellido_Paterno,AA.Simbologia"
    Empresa_ID_Reporte = ""
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Asistencias.EOF Then
        With Rs_Consulta_Adm_Asistencias
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE HISTORICO FALTAS INJUSTIFICADAS Y RETARDOS", DateAdd("s", 1, Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Historico_Faltas_Retardos_Fecha_Termino.Value))
            Print #2, "No.Empleado|Empleado|Departamento|Supervisor|Fecha|Descripcion|Concepto|Observaciones|Creo"
            While Not .EOF
                Debug.Print .rdoColumns("Empleado_ID")
                Usuario_Creo = .rdoColumns("Usuario_Creo")
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Print #1, " "; .rdoColumns("Nombre_Empresa")
                End If
                If .rdoColumns("Departamento_ID") <> Departamento_ID_Reporte Then
                    Departamento_ID_Reporte = .rdoColumns("Departamento_ID")
                    Print #1, " "; .rdoColumns("Departamento")
                End If
                If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                    Supervisor_ID = .rdoColumns("Supervisor_ID")
                    Print #1,
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                    Print #1, "Supervisor: "; Nombre_Supervisor
                End If
                If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
                    Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
                    Print #1,
                    Print #1, " "; .rdoColumns("No_Tarjeta"); "   "; .rdoColumns("Nombre")
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "Fecha      Descripcion               Concepto          Observaciones                            Creo"
                End If
                Concepto = ""
                Observaciones = ""
                Descripcion = ""
                Select Case Trim(.rdoColumns("Simbologia"))
                    Case "F":
                        Concepto = "FALTA"
                        Descripcion = "INASISTENCIA DEL DIA"
                    Case "FJ":
                        Concepto = "FALTA JUSTIFICADA"
                        Mi_SQL = "SELECT AI.Observaciones,ISNULL(AI.Motivo,'') as Motivo,AI.Usuario_Creo "
                        Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias AI"
                        Mi_SQL = Mi_SQL & " WHERE AI.Empleado_ID = '" & Empleado_ID_Reporte & "'"
                        Mi_SQL = Mi_SQL & " AND " & Par_Fecha & Format(.rdoColumns("Fecha"), "MM/dd/yyyy") & Par_Fecha
                        Mi_SQL = Mi_SQL & " BETWEEN AI.Fecha_Inicio AND AI.Fecha_Termino"
                        Set Rs_Consulta_Adm_Inasistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Adm_Inasistencias.EOF Then
                            Usuario_Creo = Rs_Consulta_Adm_Inasistencias.rdoColumns("Usuario_Creo")
                            Descripcion = Rs_Consulta_Adm_Inasistencias.rdoColumns("Motivo")
                            Observaciones = Rs_Consulta_Adm_Inasistencias.rdoColumns("Observaciones")
                            Rs_Consulta_Adm_Inasistencias.Close
                        End If
                        Set Rs_Consulta_Adm_Inasistencias = Nothing
                    Case "RE":
                        Concepto = "RETARDO"
                        Descripcion = ""
                        Observaciones = "T. RET: " & .rdoColumns("Tiempo_Retardo") '& "FECHA Y RET: " & .rdoColumns("Fecha_Retardo") & "," & .rdoColumns("Retardos")
                End Select
                Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); Spc(1); _
                          Mid(Descripcion, 1, 26); Spc(26 - Len(Mid(Descripcion, 1, 26))); _
                          Mid(Concepto, 1, 18); Spc(18 - Len(Mid(Concepto, 1, 18))); _
                          Mid(Observaciones, 1, 40); Spc(41 - Len(Mid(Observaciones, 1, 40))); Usuario_Creo
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Departamento"); _
                    "|"; Nombre_Supervisor; _
                    "|"; Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    "|"; Descripcion; _
                    "|"; Concepto; _
                    "|"; Observaciones; _
                    "|"; Usuario_Creo
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Adm_Asistencias.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Faltas_Empleados
'DESCRIPCION: Genera el reporte de los empleados que no checaron del día seleccionado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 27-Junio-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Faltas_Empleados()
Dim crxApplication As New CRAXDRT.Application
Dim crxReport As CRAXDRT.Report
Dim crxDatabase As CRAXDRT.Database
Dim crxDatabaseTables As CRAXDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crxSubreport As CRAXDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Ruta_Aplicacion As String
Dim Rs_Consulta_Empleado As rdoResultset
Dim Rs_Inserta_Registro As rdoResultset

On Error GoTo HANDLER
    Ruta_Aplicacion = App.Path
    If Mid(Ruta_Aplicacion, Len(Ruta_Aplicacion), 1) = "\" Then
        Ruta_Aplicacion = Mid(Ruta_Aplicacion, 1, Len(Ruta_Aplicacion) - 1)
    End If
    Set crxReport = crxApplication.OpenReport(Ruta_Aplicacion & "\Reportes\Rpt_Empleados_Sin_Checada.rpt")
    
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo Database, Database, User_Conexion, User_Password
            End Select
        Next
    End With
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Ruta_Imagen"
                crParamDef.AddCurrentValue App.Path & "\Perfil"
            Case "Fecha_Retardo"
                crParamDef.AddCurrentValue Format(Dtp_Fecha_Faltas_Empleados.Value, "yyyy-MM-dd")
        End Select
    Next
    
    'Borra los registros
    Mi_SQL = "DELETE FROM Tmp_Empleados_Checadas"
    Conexion_Base.Execute Mi_SQL
    Set Rs_Inserta_Registro = Conectar_Ayudante.Recordset_Agregar("Tmp_Empleados_Checadas")
    'Consulta los registros
    Mi_SQL = "SELECT Cat_Empleados.Apellido_Paterno,Cat_Empleados.Apellido_Materno,Cat_Empleados.Nombre,Cat_Empresas.Nombre AS Empresa,Cat_Departamentos.Nombre AS Departamento,Cat_Puestos.Nombre AS Puesto,Cat_Empleados.No_Tarjeta,Cat_Turnos.Nombre AS Turno,Cat_Empleados.Imagen_Perfil,Adm_Asistencias_Registro_Checadores.Fecha"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados LEFT JOIN Adm_Asistencias_Registro_Checadores ON Cat_Empleados.No_Tarjeta=Adm_Asistencias_Registro_Checadores.No_Tarjeta AND Adm_Asistencias_Registro_Checadores.Fecha='" & Format(Dtp_Fecha_Faltas_Empleados.Value, "yyyy-MM-dd") & "'"
    Mi_SQL = Mi_SQL & " ,Cat_Empresas,Cat_Puestos,Cat_Departamentos,Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empresas.Empresa_ID=Cat_Empleados.Empresa_ID AND Cat_Puestos.Puesto_ID=Cat_Empleados.Puesto_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID AND Cat_Turnos.Turno_ID=Cat_Empleados.Turno_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.FECHA IS NULL"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus='A'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Empleado.EOF
        With Rs_Inserta_Registro
            .AddNew
                .rdoColumns("Nombre") = Rs_Consulta_Empleado.rdoColumns("Apellido_Paterno") & " " & Rs_Consulta_Empleado.rdoColumns("Apellido_Materno") & " " & Rs_Consulta_Empleado.rdoColumns("Nombre")
                .rdoColumns("Empresa") = Rs_Consulta_Empleado.rdoColumns("Empresa")
                .rdoColumns("Departamento") = Rs_Consulta_Empleado.rdoColumns("Departamento")
                .rdoColumns("Puesto") = Rs_Consulta_Empleado.rdoColumns("Puesto")
                .rdoColumns("No_Tarjeta") = Rs_Consulta_Empleado.rdoColumns("No_Tarjeta")
                .rdoColumns("Turno") = Rs_Consulta_Empleado.rdoColumns("Turno")
                .rdoColumns("Imagen_Perfil") = Rs_Consulta_Empleado.rdoColumns("Imagen_Perfil")
            .Update
        End With
        Rs_Consulta_Empleado.MoveNext
    Wend
    Rs_Consulta_Empleado.Close
    Rs_Inserta_Registro.Close
    
    Frm_Ver_Reportes.Crv_Reporte.DisplayBorder = False
    Frm_Ver_Reportes.Crv_Reporte.DisplayTabs = False
    Frm_Ver_Reportes.Crv_Reporte.EnableDrillDown = False
    Frm_Ver_Reportes.Crv_Reporte.EnableRefreshButton = False
    Frm_Ver_Reportes.Crv_Reporte.ReportSource = crxReport
    Frm_Ver_Reportes.Crv_Reporte.ViewReport
    Frm_Ver_Reportes.Crv_Reporte.Zoom 100
    Set crxReport = Nothing
Exit Sub
HANDLER:
    MsgBox Err.Description
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Faltas_Empleados_Validadas
'DESCRIPCION: Genera el reporte de las faltas del día seleccionado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 18-Junio-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Faltas_Empleados_Validadas()
Dim crxApplication As New CRAXDDRT.Application
Dim crxReport As CRAXDDRT.Report
Dim crxDatabase As CRAXDDRT.Database
Dim crxDatabaseTables As CRAXDDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
Dim crxSections As CRAXDDRT.Sections
Dim crxSection As CRAXDDRT.Section
Dim crxSubreport As CRAXDDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Nombre_Reporte As String
Dim Rs_Inserta_Registro As rdoResultset
Dim Rs_Consulta_Empleado As rdoResultset
Dim Rs_Consulta_Supervisor As rdoResultset
Dim Nombre_Supervisor As String

On Error GoTo HANDLER
    'Limpia la tabla de faltas
    Mi_SQL = "DELETE FROM Tmp_Empleados_Faltas"
    Conexion_Base.Execute Mi_SQL
    
    'Busca los registros de empleados para el reporte
    Set Rs_Inserta_Registro = Conectar_Ayudante.Recordset_Agregar("Tmp_Empleados_Faltas")
    'Consulta los registros
    Mi_SQL = "SELECT Adm_Asistencias.Fecha,Cat_Turnos.Nombre AS Turno,Adm_Asistencias.No_Tarjeta,Cat_Empleados.Imagen_Perfil,Cat_Departamentos.Nombre AS Departamento,Cat_Puestos.Nombre AS Puesto"
    Mi_SQL = Mi_SQL & " ,Cat_Empleados.Apellido_Paterno,Cat_Empleados.Apellido_Materno,Cat_Empleados.Nombre,ISNULL(Cat_Gaps.Nombre,'') AS Tripulacion,Cat_Empleados.Supervisor_ID,Cat_Empleados.Fecha_Ingreso"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias INNER JOIN Cat_Empleados ON Adm_Asistencias.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " INNER JOIN Cat_Turnos ON Adm_Asistencias.Turno_ID=Cat_Turnos.Turno_ID"
    Mi_SQL = Mi_SQL & " INNER JOIN Cat_Departamentos ON Cat_Empleados.Departamento_ID=Cat_Departamentos.Departamento_ID"
    Mi_SQL = Mi_SQL & " INNER JOIN Cat_Puestos ON Cat_Empleados.Puesto_ID=Cat_Puestos.Puesto_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Gaps ON Cat_Empleados.Gap_ID=Cat_Gaps.Gap_ID"
    Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias.Simbologia='F'"         'Simbología de la falta
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Fecha='" & Format(Dtp_Fecha_Faltas_Empleados.Value, "MM/dd/yyyy") & "'"
    If Cmb_Turno_Faltas.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Turno_ID='" & Format(Cmb_Turno_Faltas.ItemData(Cmb_Turno_Faltas.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias.No_Tarjeta"
    Set Rs_Consulta_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Empleado.EOF
        'Consulta el nombre del supervisor
        If Not IsNull(Rs_Consulta_Empleado.rdoColumns("Supervisor_ID")) Then
            Mi_SQL = "SELECT (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Supervisor"
            Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Rs_Consulta_Empleado.rdoColumns("Supervisor_ID") & "'"
            Set Rs_Consulta_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Supervisor.EOF Then
                Nombre_Supervisor = Rs_Consulta_Supervisor.rdoColumns("Supervisor")
            Else
                Nombre_Supervisor = ""
            End If
            Rs_Consulta_Supervisor.Close
        Else
            Nombre_Supervisor = ""
        End If
        'Almacena el registro para el reporte
        With Rs_Inserta_Registro
            .AddNew
                .rdoColumns("Fecha") = UCase(Format(Rs_Consulta_Empleado.rdoColumns("Fecha"), "dd MMMM yyyy"))
                .rdoColumns("Turno") = Rs_Consulta_Empleado.rdoColumns("Turno")
                .rdoColumns("No_Tarjeta") = Rs_Consulta_Empleado.rdoColumns("No_Tarjeta")
                .rdoColumns("Ruta_Imagen") = PG_Ruta_Fotos & "\" & Rs_Consulta_Empleado.rdoColumns("Imagen_Perfil")
                .rdoColumns("Departamento") = Rs_Consulta_Empleado.rdoColumns("Departamento")
                .rdoColumns("Clase") = Rs_Consulta_Empleado.rdoColumns("Puesto")
                .rdoColumns("Nombre") = Rs_Consulta_Empleado.rdoColumns("Apellido_Paterno") & " " & Rs_Consulta_Empleado.rdoColumns("Apellido_Materno") & " " & Rs_Consulta_Empleado.rdoColumns("Nombre")
                .rdoColumns("Gerencia") = ""
                .rdoColumns("Area") = Rs_Consulta_Empleado.rdoColumns("Tripulacion")
                .rdoColumns("Supervisor") = Nombre_Supervisor
                .rdoColumns("Antiguedad") = Calcula_Edad(Rs_Consulta_Empleado.rdoColumns("Fecha_Ingreso"))
            .Update
        End With
        Rs_Consulta_Empleado.MoveNext
    Wend
    Rs_Consulta_Empleado.Close
    Rs_Inserta_Registro.Close
    
    'Asigna el formato de la factura a la variable
    Set crxReport = crxApplication.OpenReport(App.Path & "\Reportes\Rpt_Faltas_Empleados.Rpt")
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo Database, Database, User_Conexion, User_Password
            End Select
        Next
    End With
    
'    'Ver el previo del reporte
'    Frm_Ver_Reportes.Crv_Reporte.DisplayBorder = False
'    Frm_Ver_Reportes.Crv_Reporte.DisplayTabs = False
'    Frm_Ver_Reportes.Crv_Reporte.EnableDrillDown = False
'    Frm_Ver_Reportes.Crv_Reporte.EnableRefreshButton = False
'    Frm_Ver_Reportes.Crv_Reporte.ReportSource = crxReport
'    Frm_Ver_Reportes.Crv_Reporte.ViewReport
'    Frm_Ver_Reportes.Crv_Reporte.Zoom 100

    'Asigna el nombre del reporte
    Nombre_Reporte = App.Path & "\Faltas\" & Format(Now, "yyyyMMddHHmm") & ".pdf"
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
    crxReport.ExportOptions.DiskFileName = Nombre_Reporte

    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat
    crxReport.ExportOptions.PDFExportAllPages = True
    'Oculta el progreso de la exportacion
    crxReport.DisplayProgressDialog = False
    'Genera la exportación del documento
    crxReport.Export (False)
    'Destruye el documento
    Set crxReport = Nothing
     ShellExecute Me.hwnd, "open", Nombre_Reporte, "", "", 4
    Exit Sub
HANDLER:
    Printer.EndDoc
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Historico_Permisos
'DESCRIPCION: Genera el reporte de los permisos de los empleados
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 28 Febrero 2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Historico_Permisos()
Dim Rs_Consulta_Adm_Permisos As rdoResultset               'Información de inasistencias o faltas
Dim Mi_SQL As String                                       'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                                   'ID de la empresa
Dim Empleado_ID_Reporte As String                          'ID de la empresa
Dim Supervisor_ID  As String
Dim Departamento_ID As String
Dim Nombre_Supervisor As String

    Mi_SQL = "SELECT AM.Simbologia,ISNULL(AM.SubSimbologia,'') AS SubSimbologia,AM.Fecha_Solicitud,ISNULL(AM.Motivo,'') AS Motivo"
    Mi_SQL = Mi_SQL & " ,AM.Observaciones,AM.Empleado_ID,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,CEM.Empresa_ID,CEM.Nombre AS Nombre_Empresa,ISNULL(CE.Supervisor_ID,'N') AS Supervisor_ID"
    Mi_SQL = Mi_SQL & " ,AM.Usuario_Creo,ISNULL(CD.Nombre,'') AS Departamento,CE.Departamento_ID"
    Mi_SQL = Mi_SQL & " ,CE.No_Tarjeta,AM.Fecha_Inicio,AM.Fecha_Termino"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Empresas CEM,Adm_Movimientos_Asistencias AM,Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AM.Empleado_ID=CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.EMpresa_ID=CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID=CD.Departamento_ID "
    'Validacion de Empresa
    If Cmb_Rpt_Historico_Permisos_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CEM.Empresa_ID='" & Format(Cmb_Rpt_Historico_Permisos_Empresa.ItemData(Cmb_Rpt_Historico_Permisos_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Supervisores
    If Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Rpt_Historico_Permisos_Supervisor.ItemData(Cmb_Rpt_Historico_Permisos_Supervisor.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Historico_Permisos_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AM.Empleado_ID='" & Format(Cmb_Rpt_Historico_Permisos_Empleado.ItemData(Cmb_Rpt_Historico_Permisos_Empleado.ListIndex), "00000") & "'"
    End If
    'Validacion del departamento
    If Cmb_Rpt_Historico_Permisos_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Rpt_Historico_Permisos_Departamento.ItemData(Cmb_Rpt_Historico_Permisos_Departamento.ListIndex), "00000") & "'"
    End If
    ''Rango de Fechas
    Mi_SQL = Mi_SQL & " AND ((AM.Fecha_Inicio BETWEEN '" & Format(Dtp_Rpt_Historico_Permisos_Fecha_inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Historico_Permisos_Fecha_Termino.Value, "MM/dd/yyyy") & "')"
    Mi_SQL = Mi_SQL & " OR (AM.Fecha_Termino BETWEEN '" & Format(Dtp_Rpt_Historico_Permisos_Fecha_inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Historico_Permisos_Fecha_Termino.Value, "MM/dd/yyyy") & "'))"
    'Mi_SQL = Mi_SQL & " AND AM.Fecha >=" & Par_Fecha & Format(Dtp_Rpt_Historico_Permisos_Fecha_inicio.Value, "MM/dd/yyyy") & Par_Fecha
    'Mi_SQL = Mi_SQL & " AND AM.Fecha <=" & Par_Fecha & Format(Dtp_Rpt_Historico_Permisos_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CD.Nombre,CE.Departamento_ID,CE.Supervisor_ID,AM.Fecha_Solicitud,AM.Empleado_ID"
    Mi_SQL = Mi_SQL & " ,CEM.Nombre,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AM.Simbologia,AM.SubSimbologia"
    Mi_SQL = Mi_SQL & " ,AM.Motivo,AM.Observaciones,AM.Usuario_Creo,CE.No_Tarjeta,AM.Fecha_Inicio,AM.Fecha_Termino"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CE.Departamento_ID, CE.Supervisor_ID,CE.Apellido_Paterno,AM.Fecha_Solicitud"
    Empresa_ID_Reporte = ""
    Set Rs_Consulta_Adm_Permisos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Permisos.EOF Then
        With Rs_Consulta_Adm_Permisos
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE HISTORICO DE PERMISOS", DateAdd("s", 1, Dtp_Rpt_Historico_Permisos_Fecha_inicio.Value), DateAdd("s", 1, Dtp_Rpt_Historico_Permisos_Fecha_Termino.Value))
            Print #2, "Empresa|Departamento|Supervisor|No.Empleado|Nombre|Solicitud|Fecha Inicio|Fecha Termino|Tipo|Motivo-Horas|Observaciones"
            Supervisor_ID = ""
            Nombre_Supervisor = ""
            While Not .EOF
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Print #1, " "; .rdoColumns("Nombre_Empresa")
                    Supervisor_ID = ""
                End If
                If .rdoColumns("Departamento_ID") <> Departamento_ID Then
                    Supervisor_ID = .rdoColumns("Departamento_ID")
                    Print #1,
                    Print #1, "Departamento: "; .rdoColumns("Departamento")
                End If
                If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                    Supervisor_ID = .rdoColumns("Supervisor_ID")
                    Print #1,
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                    Print #1, "Supervisor: "; Nombre_Supervisor
                End If
                If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
                    Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
                    Print #1,
                    Print #1, " "; .rdoColumns("Nombre")
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "Solicito      F.Incio      F.Termino   Tipo      Motivo-Horas          Observaciones                                      "
                End If
                Print #1, Format(.rdoColumns("Fecha_Solicitud"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Termino"), "dd/MMM/yyyy"); _
                    Spc(2); .rdoColumns("Simbologia"); _
                    Spc(8 - Len(.rdoColumns("Simbologia"))); Mid(.rdoColumns("Motivo"), 1, 20); _
                    Spc(22 - Len(Mid(.rdoColumns("Motivo"), 1, 20))); Mid(.rdoColumns("Observaciones"), 1, 50)
                Print #2, .rdoColumns("Nombre_Empresa"); _
                    "|"; .rdoColumns("Departamento"); _
                    "|"; Nombre_Supervisor; _
                    "|"; .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; Format(.rdoColumns("Fecha_Solicitud"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Fecha_Termino"), "dd/MMM/yyyy"); _
                    "|"; .rdoColumns("Simbologia"); _
                    "|"; .rdoColumns("Motivo"); _
                    "|"; Conectar_Ayudante.Quitar_Caracter(.rdoColumns("Observaciones"), Chr(13))
                .MoveNext
            Wend
            .Close
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Set Rs_Consulta_Adm_Permisos = Nothing
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Generar_Reporte_Horas_Trabajadas_Empleado
    'DESCRIPCIÓN:          Genera el reporte de horas trabajadas por empleado en un periodo determinado
    'PARÁMETROS:
    'CREO:                 Yañez Rodriguez Diego Neftali
    'FECHA_CREO:           28 Febrero 2009
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Generar_Reporte_Horas_Trabajadas_Empleado_Periodo()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset     'Informacion de los tiempo muertos
Dim Mi_SQL As String                                                'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                                            'ID de la empresa
Dim Empleado_ID_Reporte As String                                            'ID de la empleado
Dim Horas_Empresa As Double                                         'Horas total por empresa
Dim Horas_Extra_Empresa As Double                                   'Horas exta total por empresa
Dim Horas_Total As Double                                           'Horas total
Dim Horas_Extra_Total As Double                                     'Horas extra
Dim Cadena_Grid As String                                           'Define la informacion del encabezado
Dim Cadena_Grid_2 As String                                           'Define la informacion del encabezado
Dim Cont_Fila As Integer                                            'recorre los dias
Dim Encontrado As Boolean
Dim Fila_Encontrado As Integer
Dim Cont_Col As Integer
Dim Total_Horas_Empleado As Double
Dim Fecha_Encontrada As Boolean
Dim Supervisor_ID As String                                         'Identificador del supervisor
Dim Departamento_ID As String                                       'Identificador del departamento
Dim Nombre_Supervisor As String                                     'Nombre del supervisor
Mi_SQL = "SELECT ISNULL(AA.Horas_Aprobadas,0) as Horas, "
Mi_SQL = Mi_SQL & " ISNULL(AA.Horas_Extra,0) as Horas_Extras, "
Mi_SQL = Mi_SQL & " ISNULL(AA.Horas_Aprobadas,0) + ISNULL(AA.Horas_Extra,0) as Horas_Total,"
Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre,"
Mi_SQL = Mi_SQL & " CEM.Empresa_ID,CEM.Nombre as Nombre_Empresa,"
Mi_SQL = Mi_SQL & " AA.Hora_Entrada, AA.Hora_Salida, AA.Fecha, CE.Empleado_ID,"
Mi_SQL = Mi_SQL & " ISNULL(CE.Supervisor_ID,'N') as Supervisor_ID,"
Mi_SQL = Mi_SQL & " ISNULL(CD.Nombre,'') as Departamento, CD.Departamento_ID"
Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA, Cat_Empleados CE, Cat_Empresas CEM, Cat_Departamentos CD"""
Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = CE.Empleado_ID"
Mi_SQL = Mi_SQL & " AND CE.EMpresa_ID = CEM.Empresa_ID"
Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = CD.Departamento_ID"
'Validacion de Empresa
If Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Empresa.ListIndex), "00000") & "'"
End If
'Validacion de Empleado
If Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Supervisor.ListIndex), "00000") & "'"
End If
'Validacion de Empleado
If Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " AND AA.Empleado_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Empleado.ListIndex), "00000") & "'"
End If
'Validacion del departamento
If Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = '" & Format(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ItemData(Cmb_Rpt_Horas_Trabajadas_Empleado_Departamento.ListIndex), "00000") & "'"
End If

'Rango de Fechas
Mi_SQL = Mi_SQL & " AND AA.Fecha >=" & Par_Fecha & Format(Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
Mi_SQL = Mi_SQL & " AND AA.Fecha <=" & Par_Fecha & Format(Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
'Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CEM.Nombre,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre"
Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CE.Supervisor_ID,CE.Apellido_Paterno,AA.Fecha"
Empresa_ID_Reporte = ""
Empleado_ID_Reporte = ""
Supervisor_ID = ""

''Ejecuta la consulta
Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Adm_Asistencias
    If Not .EOF Then
        MDIFrm_Apl_Principal.MousePointer = 11
        'Prepara el grid para agregar la información
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "SEMANAL" Then
            Grid_Rpt_Informacion.Rows = 0
            Grid_Rpt_Informacion.Cols = DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value) + 6
            Cadena_Grid = "Empresa_ID" & Chr(9) & "Empleado_ID" & Chr(9) & "Supervisor_ID" & Chr(9) & "Nombre"
            For Cont_Fila = 0 To DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value)
                Cadena_Grid = Cadena_Grid & Chr(9) & Format(DateAdd("d", Cont_Fila, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value), "MM/dd/yyyy")
            Next
            Grid_Rpt_Informacion.AddItem Cadena_Grid
        End If
        If Cmb_Rpt_Horas_Trabajadas_Empleado_Periodo.Text = "MENSUAL" Then
            Grid_Rpt_Informacion.Rows = 0
            Grid_Rpt_Informacion.Cols = DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value) + 6
            Cadena_Grid = "Empresa_ID" & Chr(9) & "Empleado_ID" & Chr(9) & "Supervisor_ID" & Chr(9) & "Nombre"
            For Cont_Fila = 0 To DateDiff("d", Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value)
                Cadena_Grid = Cadena_Grid & Chr(9) & Format(DateAdd("d", Cont_Fila, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value), "MM/dd/yyyy")
            Next
            Grid_Rpt_Informacion.AddItem Cadena_Grid
        End If
        While Not .EOF
            Debug.Print Empleado_ID_Reporte
            Encontrado = False
            Fila_Encontrado = 0
            If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
                Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
                Supervisor_ID = ""
            End If
            If .rdoColumns("Departamento_ID") <> Departamento_ID Then
                Departamento_ID = .rdoColumns("Departamento_ID")
            End If
            If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                Supervisor_ID = .rdoColumns("Supervisor_ID")
            End If
            'busca si el registro ya se ha agregado
            For Cont_Fila = 0 To Grid_Rpt_Informacion.Rows - 1
                If Grid_Rpt_Informacion.TextMatrix(Cont_Fila, 2) = Empleado_ID_Reporte Then
                    Encontrado = True
                    Fila_Encontrado = Cont_Fila
                    Exit For
                End If
            Next
            If Encontrado = False Then
                Grid_Rpt_Informacion.AddItem .rdoColumns("Nombre_Empresa") & Chr(9) & .rdoColumns("Departamento") & Chr(9) & .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Supervisor_ID") & Chr(9) & .rdoColumns("Nombre")
                Fila_Encontrado = Grid_Rpt_Informacion.Rows - 1
            End If
            Total_Horas_Empleado = 0
            'Busca la columna donde insertara la fecha
            Fecha_Encontrada = False
            For Cont_Col = 0 To Grid_Rpt_Informacion.Cols - 2
                If Cont_Col >= 5 Then
                    If DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Grid_Rpt_Informacion.TextMatrix(0, Cont_Col), "MM/dd/yyyy")) = 0 And Fecha_Encontrada = False Then
                        Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col) = Val(.rdoColumns("Horas")) 'Format(.rdoColumns("Hora_Entrada"), "HH:mm") & "-" & Format(.rdoColumns("Hora_Salida"), "HH:mm")
                        Fecha_Encontrada = True
                    End If
                End If
                If Cont_Col >= 5 Then
                    Total_Horas_Empleado = Total_Horas_Empleado + Val(Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col))
                End If
            Next
            Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Grid_Rpt_Informacion.Cols - 1) = Total_Horas_Empleado
            .MoveNext
        Wend
       'Agrega el encabezado al reporte
        Call Encabezado_Reporte("REPORTE DE HORAS TRABAJADAS POR EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Horas_Trabajadas_Empleado_Fecha_Termino.Value))
        With Grid_Rpt_Informacion
            Cadena_Grid = Conectar_Ayudante.Agregar_Espacios("Nombre", 25)
            Cadena_Grid_2 = "Nombre||"
            For Cont_Col = 5 To .Cols - 2
                Cadena_Grid = Cadena_Grid & Conectar_Ayudante.Alinea_Derecha(Format(.TextMatrix(0, Cont_Col), "dd"), 5)
                Cadena_Grid_2 = Cadena_Grid_2 & "|" & Format(.TextMatrix(0, Cont_Col), "dd")
            Next
            Cadena_Grid = Cadena_Grid & Chr(9) & "Total"
            Cadena_Grid_2 = Cadena_Grid_2 & "|Total"
            Print #1, Cadena_Grid
            Print #2, Cadena_Grid_2
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "--------------------------------------------------------------------------------------------------------------------------"
            Cadena_Grid = ""
            Supervisor_ID = ""
            Departamento_ID = ""
            Empresa_ID_Reporte = ""
            For Cont_Fila = 1 To .Rows - 1
                If Empresa_ID_Reporte <> Trim(.TextMatrix(Cont_Fila, 0)) Then
                    Empresa_ID_Reporte = Trim(.TextMatrix(Cont_Fila, 0))
                    Print #1,
                    Print #2,
                    Print #1, Empresa_ID_Reporte
                    Print #2, Empresa_ID_Reporte
                    Supervisor_ID = ""
                End If
                If Departamento_ID <> Trim(.TextMatrix(Cont_Fila, 1)) Then
                    Departamento_ID = Trim(.TextMatrix(Cont_Fila, 1))
                    Print #1,
                    Print #2,
                    Print #1, "Departamento: " & Departamento_ID
                    Print #2, "Departamento:|" & Departamento_ID
                End If
                If Supervisor_ID <> Trim(.TextMatrix(Cont_Fila, 3)) Then
                    Supervisor_ID = Trim(.TextMatrix(Cont_Fila, 3))
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                    If Supervisor_ID = "N" Then
                        Nombre_Supervisor = "Supervisor no asignado"
                    End If
                    Print #1,
                    Print #2,
                    Print #1, "Supervisor: " & Nombre_Supervisor
                    Print #2, "Supervisor:|" & Nombre_Supervisor
                End If
                Cadena_Grid = Trim(Mid(.TextMatrix(Cont_Fila, 3), 1, 25))
                Cadena_Grid = Conectar_Ayudante.Agregar_Espacios(Cadena_Grid, 25)
                Cadena_Grid_2 = .TextMatrix(Cont_Fila, 3) & "||"
                For Cont_Col = 4 To .Cols - 2
                    Cadena_Grid = Cadena_Grid & Conectar_Ayudante.Alinea_Derecha(.TextMatrix(Cont_Fila, Cont_Col), 5)
                    Cadena_Grid_2 = Cadena_Grid_2 & "|" & .TextMatrix(Cont_Fila, Cont_Col)
                Next
                Cadena_Grid = Cadena_Grid + Conectar_Ayudante.Alinea_Derecha(.TextMatrix(Cont_Fila, Cont_Col), 7)
                Cadena_Grid_2 = Cadena_Grid_2 & "|" & .TextMatrix(Cont_Fila, Cont_Col)
                Print #1, Cadena_Grid
                Print #2, Cadena_Grid_2
            Next
        End With
        .Close
        Call Finalizar_Reporte(True)
        Btn_Imprimir.Enabled = True
        Btn_Exportar.Enabled = True
        Btn_Regresar.Enabled = True
        Btn_Salir.Enabled = True
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado", Me)
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
End With
Set Rs_Consulta_Adm_Asistencias = Nothing
'Haya o no haya registros se cambia el Puntero del Mouse
 MDIFrm_Apl_Principal.MousePointer = 0

End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Asistencia_Empleados
'DESCRIPCION: Genera el reporte de asistencia de empleados
'PARAMETROS :
'CREO       : Diego Neftali Yañez Rodriguez
'FECHA_CREO : 28-Febrero-2009
'MODIFICO   : Sergio Ulises Durán Hernández
'FECHA_MODIFICO: 02-Julio-2014
'CAUSA_MODIFICO: Adecuaciones para SRG
'*******************************************************************************
Private Sub Generar_Reporte_Asistencia_Empleados()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Informacion de los tiempo muertos

    'Genera la consulta
    Mi_SQL = "SELECT ISNULL(AA.Horas_Aprobadas,0) as Horas, "
    Mi_SQL = Mi_SQL & " (aa.Horas_Aprobadas % 1) as Diferencia, FLOOR(aa.Horas_Aprobadas) as Enteros,"
    'Mi_SQL = Mi_SQL & " ((datediff(n,ISNULL(AA.Hora_Entrada,0),ISNULL(AA.Hora_Salida,0))/60) -"
    'Mi_SQL = Mi_SQL & " (datediff(n,ISNULL(AA.Hora_Salida_Comida,0),ISNULL(AA.Hora_Entrada_Comida,0))/60)) as Horas_Reales,"
    Mi_SQL = Mi_SQL & " (cast((casT(datediff(n,ISNULL(AA.Hora_Entrada,0),ISNULL(AA.Hora_Salida,0)) as Decimal(18,2))/60)as decimal(18,2)) - "
    Mi_SQL = Mi_SQL & " cast((casT(datediff(n,ISNULL(AA.Hora_Entrada_Comida,0),ISNULL(AA.Hora_Salida_Comida,0)) as Decimal(18,2))/60)as decimal(18,2))) as Horas_Reales, "
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre,"
    Mi_SQL = Mi_SQL & " CEM.Empresa_ID,CEM.Nombre as Nombre_Empresa,"
    Mi_SQL = Mi_SQL & " AA.Fecha, CE.No_Tarjeta, CE.Empleado_ID, AA.Simbologia, AA.Referencia,"
    Mi_SQL = Mi_SQL & " ISNULL(CE.Supervisor_ID,'N') as Supervisor_ID,"
    Mi_SQL = Mi_SQL & " ISNULL(CD.Nombre,'') as Departamento, CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA, Cat_Empleados CE, Cat_Empresas CEM, Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.EMpresa_ID = CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = CD.Departamento_ID"
    'Validacion de Empresa
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Supervisor
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AA.Empleado_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    'Validacion del departamento
    If Cmb_Rpt_Asistencia_Empleados_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Departamento.ItemData(Cmb_Rpt_Asistencia_Empleados_Departamento.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND AA.Fecha >=" & Par_Fecha & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND AA.Fecha <=" & Par_Fecha & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
'    Mi_SQL = Mi_SQL & " ORDER BY CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " ORDER BY AA.No_Tarjeta, AA.Fecha ASC "
    'Ejecuta la consulta
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            'Prepara el grid para agregar la información
            Open Ruta_Temporal & Reporte & ".txt" For Output As #1
            Open Ruta_Temporal & Reporte & "xls.txt" For Output As #2 'Reporte a xls
    
            Call Reporte_Asistencias_Empleados_Reporte(0, 2, Mi_SQL, False)
'            Call Reporte_Asistencias_Empleados_Excel(Mi_SQL)
            Call Finalizar_Reporte(True)
        Else
            MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
            Call Finalizar_Reporte(False)
        End If
    End With
    Btn_Imprimir.Enabled = True
    Btn_Exportar.Enabled = True
    Btn_Regresar.Enabled = True
    Btn_Salir.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias", Me)
    'Haya o no haya registros se cambia el Puntero del Mouse
    Set Rs_Consulta_Adm_Asistencias = Nothing
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Reporte_Asistencias_Empleados_Excel(Cadena_Consulta As String)
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Informacion de los tiempo muertos
Dim Mi_SQL As String                            'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                'ID de la empresa
Dim Empleado_ID_Reporte As String               'ID de la empleado
Dim Supervisor_ID As String                     'ID del supervisor
Dim Nombre_Supervisor As String                 'Nombre del supervisor
Dim Cadena_Grid As String                       'Define la informacion del encabezado
Dim Cadena_Grid_2 As String                     'Define la informacion del encabezado
Dim Cadena_Grid_Excel As String                 'Define la informacion para la exportacion a excel
Dim Cadena_Grid_Excel_2 As String               'Define la informacion para la exportacion a excel
Dim Cadena_Firmas_Excel As String               'Define la informacion para la cadena de firmas
Dim Cadena_Firmas_Excel_2 As String             'Define la informacion para la cadena de firmas
Dim Cont_Fila As Integer                        'Recorre los dias
Dim Encontrado As Boolean                       'Define si el registro se encontro
Dim Fila_Encontrado As Integer                  'Define la fila donde se encontro el registro
Dim Cont_Col As Integer                         'Contador de columnas del grid
Dim Fecha_Encontrada As Boolean                 'Define si la fecha se encontro en el grid
Dim Fecha As Date                               'Indica la fecha del registro
Dim Observaciones As String                     'Mantiene las observaciones de las incidencias
Dim Horas As String                             'Define las horas aprobadas del empleado
Dim Departamento_ID As String

    Empresa_ID_Reporte = ""
    Empleado_ID_Reporte = ""
    Supervisor_ID = ""
    Departamento_ID = ""
    
    'Ejecuta la consulta
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Cadena_Consulta)
     Grid_Rpt_Informacion.Rows = 0
     Grid_Rpt_Informacion.Cols = (DateDiff("d", Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value) * 3) + 8
     Cadena_Grid = "Empresa_ID" & Chr(9) & "Departamento" & Chr(9) & "Empleado_ID" & Chr(9) & "Supervisor_ID" & Chr(9) & "Nombre"
     For Cont_Fila = 0 To DateDiff("d", Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value)
         Fecha = Format(DateAdd("d", Cont_Fila, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), "MM/dd/yyyy")
         Cadena_Grid = Cadena_Grid & Chr(9) & Fecha & Chr(9) & Fecha & Chr(9) & Fecha
     Next
     Grid_Rpt_Informacion.AddItem Cadena_Grid
     With Rs_Consulta_Adm_Asistencias
         While Not .EOF
             Debug.Print Empleado_ID_Reporte
             Encontrado = False
             Fila_Encontrado = 0
             If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
                 Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
                 Supervisor_ID = ""
             End If
             If .rdoColumns("Departamento_ID") <> Departamento_ID Then
                 Departamento_ID = .rdoColumns("Departamento_ID")
             End If
             If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                 Supervisor_ID = .rdoColumns("Supervisor_ID")
             End If
             'busca si el registro ya se ha agregado
             For Cont_Fila = 0 To Grid_Rpt_Informacion.Rows - 1
                 If Grid_Rpt_Informacion.TextMatrix(Cont_Fila, 2) = Empleado_ID_Reporte Then
                     Encontrado = True
                     Fila_Encontrado = Cont_Fila
                     Exit For
                 End If
             Next
             If Encontrado = False Then
                 Grid_Rpt_Informacion.AddItem .rdoColumns("Nombre_Empresa") & Chr(9) & .rdoColumns("Departamento") & Chr(9) & .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Supervisor_ID") & Chr(9) & .rdoColumns("Nombre")
                 Fila_Encontrado = Grid_Rpt_Informacion.Rows - 1
             End If
             'Busca la columna donde insertara la fecha
             Fecha_Encontrada = False
             For Cont_Col = 0 To Grid_Rpt_Informacion.Cols - 2
                 If Cont_Col >= 5 Then
                     If DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Grid_Rpt_Informacion.TextMatrix(0, Cont_Col), "MM/dd/yyyy")) = 0 And Fecha_Encontrada = False Then
                         Observaciones = ""
                         Horas = ""
                         If Not IsNull(.rdoColumns("Referencia")) Then
                             Mi_SQL = "SELECT No_Movimiento, Observaciones FROM Adm_Movimientos_Asistencias"
                             Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(.rdoColumns("Referencia")) & "'"
                             Observaciones = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Observaciones")
                         End If
                         Horas = CStr(.rdoColumns("Horas"))
                         If Val(.rdoColumns("Horas_Reales")) <> Val(.rdoColumns("Horas")) Then
                             Horas = Horas & "/"
                         End If
                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col) = .rdoColumns("Simbologia")
                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col + 1) = Horas
                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col + 2) = Observaciones
                         Fecha_Encontrada = True
                         Exit For
                     End If
                 End If
             Next
             .MoveNext
         Wend
        'Agrega el encabezado al reporte
         Call Encabezado_Reporte_Excel("REPORTE DE ASISTENCIAS DE EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value))
         With Grid_Rpt_Informacion
             Cadena_Grid_Excel = "Nombre"
             Cadena_Grid_Excel_2 = "|"
            For Cont_Col = 5 To .Cols - 2 Step 3
                 Cadena_Grid_Excel = Cadena_Grid_Excel & "||" & Format(.TextMatrix(0, Cont_Col), "dd") & "|"
                 Cadena_Grid_Excel_2 = Cadena_Grid_Excel_2 & "Asist.|Hrs.|Observaciones" & "|"
                 Cadena_Firmas_Excel = Cadena_Firmas_Excel & "||_______________________|"
                 Cadena_Firmas_Excel_2 = Cadena_Firmas_Excel_2 & "||AUTORIZADO POR GRTE ADM|"
             Next
             Print #2, Cadena_Grid_Excel
             Print #2, Cadena_Grid_Excel_2
             'Print #2, "--------------------------------------------------------------------------------------------------------------------------"
             Cadena_Grid = ""
             Supervisor_ID = ""
             Departamento_ID = ""
             Empresa_ID_Reporte = ""
             For Cont_Fila = 1 To .Rows - 1
                 If Empresa_ID_Reporte <> Trim(.TextMatrix(Cont_Fila, 0)) Then
                     Empresa_ID_Reporte = Trim(.TextMatrix(Cont_Fila, 0))
                     Print #2,
                     Print #2, Empresa_ID_Reporte
                     Supervisor_ID = ""
                 End If
                 If Departamento_ID <> Trim(.TextMatrix(Cont_Fila, 1)) Then
                     Departamento_ID = Trim(.TextMatrix(Cont_Fila, 1))
                     Supervisor_ID = ""
                     'Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                     'If Supervisor_ID = "N" Then
                         'Nombre_Supervisor = "Supervisor no asignado"
                     'End If
                     Print #2,
                     Print #2, "Departamento: " & Departamento_ID
                 End If
                 If Supervisor_ID <> Trim(.TextMatrix(Cont_Fila, 3)) Then
                     Supervisor_ID = Trim(.TextMatrix(Cont_Fila, 3))
                     Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                     If Supervisor_ID = "N" Then
                         Nombre_Supervisor = "Supervisor no asignado"
                     End If
                     Print #2,
                     Print #2, "Supervisor: " & Nombre_Supervisor
                 End If
                 Cadena_Grid_Excel = .TextMatrix(Cont_Fila, 4)
                 For Cont_Col = 5 To .Cols - 2 Step 3
                     'Cadena_Grid_2 = Cadena_Grid_2 & "Asist." & "   " & "Hrs." & "   " & "Observaciones    "
                     Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & .TextMatrix(Cont_Fila, Cont_Col) & "|" & .TextMatrix(Cont_Fila, Cont_Col + 1) & "|" & .TextMatrix(Cont_Fila, Cont_Col + 2)
                     'Cadena_Grid_Excel_2 = Cadena_Grid_Excel_2 & "|Asist.|Hrs.|Observaciones|"
                 Next
                 Print #2, Cadena_Grid_Excel
             Next
         End With
         Print #2,
         Print #2,
         Print #2, Cadena_Firmas_Excel
         Print #2, Cadena_Firmas_Excel_2
         .Close
    End With
End Sub

Private Sub Reporte_Asistencias_Empleados_Reporte(Cant_Inicio As Integer, Cant_Fin As Integer, Cadena_Consulta As String, Bandera As Boolean)
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Informacion de los tiempo muertos
Dim Mi_SQL As String                            'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                'ID de la empresa
Dim Empleado_ID_Reporte As String               'ID de la empleado
Dim Supervisor_ID As String                     'ID del supervisor
Dim Nombre_Supervisor As String                 'Nombre del supervisor
Dim Cadena_Grid As String                       'Define la informacion del encabezado
Dim Cadena_Grid_2 As String                     'Define la informacion del encabezado
Dim Cadena_Grid_Excel As String                 'Define la informacion para la exportacion a excel
Dim Cadena_Grid_Excel_2 As String               'Define la informacion para la exportacion a excel
Dim Cadena_Firmas As String                     'Define la informacion para la cadena de firmas
Dim Cadena_Firmas_2 As String                   'Define la informacion para la cadena de firmas
Dim Cadena_Firmas_Excel As String               'Define la informacion para la cadena de firmas
Dim Cadena_Firmas_Excel_2 As String             'Define la informacion para la cadena de firmas
Dim Cont_Fila As Integer                        'Recorre los dias
Dim Encontrado As Boolean                       'Define si el registro se encontro
Dim Fila_Encontrado As Integer                  'Define la fila donde se encontro el registro
Dim Cont_Col As Integer                         'Contador de columnas del grid
Dim Fecha_Encontrada As Boolean                 'Define si la fecha se encontro en el grid
Dim Fecha As Date                               'Indica la fecha del registro
Dim Observaciones As String                     'Mantiene las observaciones de las incidencias
Dim Horas As String                             'Define las horas aprobadas del empleado
Dim Departamento_ID As String
Dim Empleado_Anterior As String
    Empresa_ID_Reporte = ""
    Empleado_ID_Reporte = ""
    Supervisor_ID = ""
    Departamento_ID = ""
    
    'Ejecuta la consulta
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Cadena_Consulta)
     Grid_Rpt_Informacion.Rows = 0
     Grid_Rpt_Informacion.Cols = 5
     Cadena_Grid = ""
     Empleado_Anterior = ""
With Rs_Consulta_Adm_Asistencias
    If Not .EOF Then

        Call Encabezado_Reporte_Reporte("REPORTE DE ASISTENCIAS DE EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value))
        Print #1, "                                                         Fecha     Horas Aprobadas     Simbologia     Observaciones                                                                                      "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #2, "                                                         ||||Fecha |Horas Aprobadas|Simbologia|Observaciones |                                                                                     "
        Print #2, "-------------------------------------------------------------------------------------------------------------------------"
        While Not .EOF
        '       Validación si tarda 20 minutos más de la hora, solo almacena las horas enteras
                Dim Horas_Visualizar As Double
'                Dim Horas_Visualizar As String
                Horas_Visualizar = .rdoColumns("Horas")
                If (.rdoColumns("Diferencia")) <= 0.33 Then
'                    Horas_Visualizar = .rdoColumns("Enteros")
                    Horas_Visualizar = .rdoColumns("Enteros")
                Else
'                    Horas_Visualizar = ((.rdoColumns("Diferencia") - 0.33) * 60)
                    Horas_Visualizar = .rdoColumns("Enteros") + (.rdoColumns("Diferencia") - 0.33)
                End If
        
            If (Empleado_Anterior <> .rdoColumns("Empleado_Id")) Then
                If .rdoColumns("Supervisor_ID") = "N" Then
                    Nombre_Supervisor = "No Asignado"
                Else
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Trim(.rdoColumns("Supervisor_ID")) & "'", "Nombre_Supervisor")
                End If
                Print #1, "Empresa: " & .rdoColumns("Nombre_Empresa") & Chr(9) & "Departamento: " & .rdoColumns("Departamento") & Chr(9)
                Print #1, "Empleado: " & .rdoColumns("No_Tarjeta") & Chr(9) & " " & .rdoColumns("Nombre")
               Print #1, Spc(55); .rdoColumns("Fecha") & Chr(9); Spc(4); Mid(Horas_Visualizar, 1, 5); Spc(21 - Mid(Len(Horas_Visualizar), 1, 5)); Chr(9); .rdoColumns("Simbologia")
                Print #2, "Empresa: " & .rdoColumns("Nombre_Empresa") & Chr(9) & "Departamento: " & .rdoColumns("Departamento") & Chr(9) & "|"
                Print #2, "Empleado: " & .rdoColumns("No_Tarjeta") & Chr(9) & " " & .rdoColumns("Nombre") & "|"
                Print #2, "|"; "|"; "|"; "|"; .rdoColumns("Fecha") & "|"; _
                Horas_Visualizar & " |"; _
                .rdoColumns("Simbologia") & "|"
            
            Else
               Print #1, Spc(55); .rdoColumns("Fecha") & Chr(9); Spc(2); Mid(Horas_Visualizar, 1, 5); Spc(21 - Mid(Horas_Visualizar, 1, 5)); Chr(9); .rdoColumns("Simbologia")
               Print #2, "|"; "|"; "|"; "|"; .rdoColumns("Fecha") & "|"; _
                    Horas_Visualizar & " |"; _
                    .rdoColumns("Simbologia") & "|"
            End If
            Empleado_Anterior = .rdoColumns("Empleado_Id")
            .MoveNext
        Wend
    End If
End With

'             Encontrado = False
'             Fila_Encontrado = 0
'             If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
'                 Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
'                 Supervisor_ID = ""
'             End If
'             If .rdoColumns("Departamento_ID") <> Departamento_ID Then
'                 Departamento_ID = .rdoColumns("Departamento_ID")
'             End If
'             If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
'                 Supervisor_ID = .rdoColumns("Supervisor_ID")
'             End If
'             'busca si el registro ya se ha agregado
'             For Cont_Fila = 0 To Grid_Rpt_Informacion.Rows - 1
'                 If Grid_Rpt_Informacion.TextMatrix(Cont_Fila, 2) = Empleado_ID_Reporte Then
'                     Encontrado = True
'                     Fila_Encontrado = Cont_Fila
'                     Exit For
'                 End If
'             Next
'             If Encontrado = False Then
'                 Grid_Rpt_Informacion.AddItem .rdoColumns("Nombre_Empresa") & Chr(9) & .rdoColumns("Departamento") & Chr(9) & .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Supervisor_ID") & Chr(9) & .rdoColumns("Nombre")
'                 Fila_Encontrado = Grid_Rpt_Informacion.Rows - 1
'             End If
'             'Busca la columna donde insertara la fecha
'             Fecha_Encontrada = False
'             For Cont_Col = 0 To Grid_Rpt_Informacion.Cols - 2
'                 If Cont_Col >= 5 Then
'                     If DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Grid_Rpt_Informacion.TextMatrix(0, Cont_Col), "MM/dd/yyyy")) = 0 And Fecha_Encontrada = False Then
'                         Observaciones = ""
'                         Horas = ""
'                         If Not IsNull(.rdoColumns("Referencia")) Then
'                             Mi_SQL = "SELECT No_Movimiento, Observaciones FROM Adm_Movimientos_Asistencias"
'                             Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(.rdoColumns("Referencia")) & "'"
'                             Observaciones = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Observaciones")
'                         End If
'                         Horas = CStr(.rdoColumns("Horas"))
'                         If Val(.rdoColumns("Horas_Reales")) <> Val(.rdoColumns("Horas")) Then
'                             Horas = Horas & "/"
'                         End If
'                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col) = .rdoColumns("Simbologia")
'                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col + 1) = Horas
'                         Grid_Rpt_Informacion.TextMatrix(Fila_Encontrado, Cont_Col + 2) = Observaciones
'                         Fecha_Encontrada = True
'                         Exit For
'                     End If
'                 End If
'             Next
'             .MoveNext
'         Wend
'        'Agrega el encabezado al reporte
'         Call Encabezado_Reporte_Reporte("REPORTE DE ASISTENCIAS DE EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value))
'         With Grid_Rpt_Informacion
'             Cadena_Grid = Conectar_Ayudante.Agregar_Espacios("Nombre", 25)
'             Cadena_Grid_2 = Conectar_Ayudante.Agregar_Espacios("", 25)
'             Cadena_Firmas = Conectar_Ayudante.Agregar_Espacios("", 25)
'             Cadena_Firmas_2 = Conectar_Ayudante.Agregar_Espacios("", 25)
'             For Cont_Col = 5 To .Cols - 2 Step 3
'                 Cadena_Grid = Cadena_Grid & "          " & Format(.TextMatrix(0, Cont_Col), "dd") & "                   "
'                 Cadena_Grid_2 = Cadena_Grid_2 & "Asist." & "   " & "Hrs." & "   " & "Observaciones  "
'                 Cadena_Firmas = Cadena_Firmas & "      " & "_______________________  "
'                 Cadena_Firmas_2 = Cadena_Firmas_2 & "      " & "AUTORIZADO POR GRTE ADM  "
'             Next
'             Print #1, Cadena_Grid
'             Print #1, Cadena_Grid_2
'             Print #1, "--------------------------------------------------------------------------------------------------------------------------"
'             Cadena_Grid = ""
'             Supervisor_ID = ""
'             Departamento_ID = ""
'             Empresa_ID_Reporte = ""
'             For Cont_Fila = 1 To .Rows - 1
'                 If Empresa_ID_Reporte <> Trim(.TextMatrix(Cont_Fila, 0)) Then
'                     Empresa_ID_Reporte = Trim(.TextMatrix(Cont_Fila, 0))
'                     Print #1,
'                     Print #1, Empresa_ID_Reporte
'                     Supervisor_ID = ""
'                 End If
'                 If Departamento_ID <> Trim(.TextMatrix(Cont_Fila, 1)) Then
'                     Departamento_ID = Trim(.TextMatrix(Cont_Fila, 1))
'                     Supervisor_ID = ""
'                     'Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
'                     'If Supervisor_ID = "N" Then
'                         'Nombre_Supervisor = "Supervisor no asignado"
'                     'End If
'                     Print #1,
'                     Print #1, "Departamento: " & Departamento_ID
'                 End If
'                 If Supervisor_ID <> Trim(.TextMatrix(Cont_Fila, 3)) Then
'                     Supervisor_ID = Trim(.TextMatrix(Cont_Fila, 3))
'                     Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
'                     If Supervisor_ID = "N" Then
'                         Nombre_Supervisor = "Supervisor no asignado"
'                     End If
'                     Print #1,
'                     Print #1, "Supervisor: " & Nombre_Supervisor
'                 End If
'                 Cadena_Grid = Trim(Mid(.TextMatrix(Cont_Fila, 4), 1, 25))
'                 Cadena_Grid = Conectar_Ayudante.Agregar_Espacios(Cadena_Grid, 25) & "  "
'                 For Cont_Col = 5 To .Cols - 2 Step 3
'                     Cadena_Grid = Cadena_Grid & .TextMatrix(Cont_Fila, Cont_Col) _
'                         & Conectar_Ayudante.Alinea_Derecha("", 5 - Len(.TextMatrix(Cont_Fila, Cont_Col))) & .TextMatrix(Cont_Fila, Cont_Col + 1) _
'                         & Conectar_Ayudante.Alinea_Derecha("", 7 - Len(.TextMatrix(Cont_Fila, Cont_Col + 1))) & Left(.TextMatrix(Cont_Fila, Cont_Col + 2), 14) & Conectar_Ayudante.Alinea_Derecha("", 17 - Len(Left(.TextMatrix(Cont_Fila, Cont_Col + 2), 14)))
'                 Next
'                 Print #1, Cadena_Grid
'             Next
'         End With
'         Print #1,
'         Print #1,
'         Print #1, Cadena_Firmas
'         Print #1, Cadena_Firmas_2
'         Print #1,
'         Print #1, "."
'         .Close
'    End With
'
'    If DateAdd("d", Cant_Fin + 3, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value) <= Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value Then
'        Call Reporte_Asistencias_Empleados_Reporte(Cant_Inicio + 3, Cant_Fin + 3, Cadena_Consulta, False)
'    ElseIf DateDiff("d", DateAdd("d", Cant_Inicio + 2, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value) < 3 And DateDiff("d", DateAdd("d", Cant_Inicio + 2, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value) > 0 And Bandera = False Then
'        Call Reporte_Asistencias_Empleados_Reporte(Cant_Inicio + 3, Cant_Fin + DateDiff("d", DateAdd("d", Cant_Inicio + 2, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value), Cadena_Consulta, True)
'    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleados_No_Validados
'DESCRIPCION: Genera el reporte de los empleados que no se han validado sus horas
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 28-Febrero-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleados_No_Validados()
Dim Rs_Consulta_Adm_Asistencias_Detalles As rdoResultset               'Información de inasistencias o faltas
Dim Mi_SQL As String                                       'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String
Dim Empleado_ID_Reporte As String                          'ID de la empresa
Dim Supervisor_ID  As String
Dim Nombre_Supervisor As String

    Mi_SQL = "SELECT DIA.Fecha, DIA.Empleado_ID,"
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre,"
    Mi_SQL = Mi_SQL & " CEM.Empresa_ID,CEM.Nombre as Nombre_Empresa,"
    Mi_SQL = Mi_SQL & " ISNULL(CE.Supervisor_ID,'N') as Supervisor_ID"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Detalles DIA, Cat_Empleados CE, Cat_Empresas CEM"
    Mi_SQL = Mi_SQL & " WHERE DIA.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND DIA.Validada='N'"
    'Validacion de Supervisores
    If Cmb_Rpt_Empleados_No_Validados_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Rpt_Empleados_No_Validados_Supervisor.ItemData(Cmb_Rpt_Empleados_No_Validados_Supervisor.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND DIA.Fecha >=" & Par_Fecha & Format(Dtp_Rpt_Empleados_No_Validados_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND DIA.Fecha <=" & Par_Fecha & Format(Dtp_Rpt_Empleados_No_Validados_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CEM.Nombre,CE.Supervisor_ID,DIA.Fecha,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,DIA.Empleado_ID"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CE.Supervisor_ID,DIA.Fecha,CE.Apellido_Paterno"
    Empresa_ID_Reporte = ""
    'Ejecuta la consulta
    Set Rs_Consulta_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias_Detalles
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE EMPLEADOS NO VALIDADOS", DateAdd("s", 1, Dtp_Rpt_Empleados_No_Validados_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Empleados_No_Validados_Fecha_Termino.Value))
            Supervisor_ID = ""
            Nombre_Supervisor = ""
            While Not .EOF
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Print #1,
                    Print #2,
                    Print #1, "Empresa: "; .rdoColumns("Nombre_Empresa")
                    Print #2, "Empresa:|"; .rdoColumns("Nombre_Empresa")
                    Supervisor_ID = ""
                End If
                If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                    Nombre_Supervisor = ""
                    Supervisor_ID = .rdoColumns("Supervisor_ID")
                    Print #1,
                    Print #2,
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                    Print #1, "Supervisor: "; Nombre_Supervisor
                    Print #2, "Supervisor:|"; Nombre_Supervisor
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "Fecha      Nombre"
                    Print #2, "Fecha|Nombre"
                End If
                Print #1, Format(.rdoColumns("Fecha"), "MM/dd/yyyy"); Spc(1); .rdoColumns("Nombre")
                Print #2, Format(.rdoColumns("Fecha"), "MM/dd/yyyy"); "|"; .rdoColumns("Nombre")
                .MoveNext
            Wend
            .Close
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados", Me)
        Else
            MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
        End If
    End With
    Set Rs_Consulta_Adm_Asistencias_Detalles = Nothing
    'Haya o no haya registros se cambia el Puntero del Mouse
     MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleados_Baja
'DESCRIPCION: Genera el reporte de los empleados que estan en baja
'PARAMETROS :
'CREO:      : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleados_Baja()
Dim Rs_Consulta_Cat_Empleados_Baja As rdoResultset               'Información de inasistencias o faltas
Dim Mi_SQL As String                                       'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String
Dim Empleado_ID_Reporte As String                          'ID de la empresa
Dim Departamento As String
    
    Mi_SQL = "SELECT CE.No_Tarjeta,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,CEM.Empresa_ID,CEM.Nombre AS Nombre_Empresa,CMB.Nombre AS Baja"
    Mi_SQL = Mi_SQL & " ,CE.Comentarios_Baja,CE.Fecha_Ingreso,CE.Fecha_Baja"
    Mi_SQL = Mi_SQL & " ,CP.Nombre AS Puesto,CD.Nombre AS Departamento, CE.Imagen_Perfil"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Empresas CEM,Cat_Puestos CP"
    Mi_SQL = Mi_SQL & " ,Cat_Departamentos CD,Cat_Motivos_Baja CMB"
    Mi_SQL = Mi_SQL & " WHERE CE.Empresa_ID=CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Motivo_Baja_ID=CMB.Motivo_Baja_ID"
    Mi_SQL = Mi_SQL & " AND CE.Puesto_ID=CP.Puesto_ID"
    Mi_SQL = Mi_SQL & " AND CE.Departamento_ID=CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND CE.Estatus='I'"
    'Validacion de Empresa
    If Cmb_Rpt_Empleados_Baja_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Rpt_Empleados_Baja_Empresa.ItemData(Cmb_Rpt_Empleados_Baja_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Departamentos
    If Cmb_Rpt_Empleados_Baja_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Rpt_Empleados_Baja_Departamento.ItemData(Cmb_Rpt_Empleados_Baja_Departamento.ListIndex), "00000") & "'"
    End If
    'Validacion de Puesto
    If Cmb_Rpt_Empleados_Baja_Puesto.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Puesto_ID='" & Format(Cmb_Rpt_Empleados_Baja_Puesto.ItemData(Cmb_Rpt_Empleados_Baja_Puesto.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND CE.Fecha_Baja BETWEEN '" & Format(Dtp_Rpt_Empleados_Baja_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Empleados_Baja_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " GROUP BY CEM.EMpresa_ID,CEM.Nombre,CD.Nombre,CP.Nombre,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,CMB.Nombre,CE.Comentarios_Baja,CE.Fecha_Ingreso,CE.Fecha_Baja, CE.Imagen_Perfil"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.EMpresa_ID,CD.Nombre,CP.Nombre,CE.Fecha_Baja,CE.Apellido_Paterno"
    Empresa_ID_Reporte = ""
    Set Rs_Consulta_Cat_Empleados_Baja = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Empleados_Baja
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE EMPLEADOS EN BAJA", DateAdd("s", 1, Dtp_Rpt_Empleados_Baja_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Empleados_Baja_Fecha_Termino.Value))
            While Not .EOF
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Print #1,
                    Print #1, "Empresa: "; .rdoColumns("Nombre_Empresa")
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #2,
                    Print #2, "Empresa:|"; .rdoColumns("Nombre_Empresa")
                    Print #2,
                    Print #2, "Nomina|Departamento|Puesto|Nombre|Motivo de Baja|F. Ingreso|F. Baja|Comentario|Foto"
                End If
                If .rdoColumns("Departamento") <> Departamento Then
                    Departamento = .rdoColumns("Departamento")
                    Print #1,
                    Print #1, " Departamento: "; Departamento
                    Print #1, " -------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " Nomina Puesto                           Nombre                           Motivo de Baja           F. Ingreso       F.Baja"
                End If
                Print #1, _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Tarjeta"), 6); _
                    Spc(1); Mid(.rdoColumns("Puesto"), 1, 25); _
                    Spc(27 - Len(Mid(.rdoColumns("Puesto"), 1, 25))); _
                    Spc(2); Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); Mid(.rdoColumns("Baja"), 1, 16); _
                    Spc(18 - Len(Mid(.rdoColumns("Baja"), 1, 16))); Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Baja"), "dd/MMM/yyyy")
                Print #1, "Comentario: "; .rdoColumns("Comentarios_Baja")
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; Departamento; _
                    "|"; .rdoColumns("Puesto"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Baja"); _
                    "|"; Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Fecha_Baja"), "dd/MMM/yyyy"); _
                    "|"; Replace(Replace(.rdoColumns("Comentarios_Baja"), Chr(10), " "), Chr(13), " "); _
                    "|"; "Perfil\"; .rdoColumns("Imagen_Perfil")

                .MoveNext
            Wend
            .Close
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja", Me)
        Else
            MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
        End If
    End With
    Set Rs_Consulta_Cat_Empleados_Baja = Nothing
    'Haya o no haya registros se cambia el Puntero del Mouse
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleados_Alta
'DESCRIPCION: Genera el reporte de los empleados que estan en Alta
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 05-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleados_Alta()
Dim Rs_Consulta_Cat_Empleados As rdoResultset               'Información de inasistencias o faltas
Dim Mi_SQL As String                                       'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String
Dim Empleado_ID_Reporte As String                          'ID de la empresa
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT CE.No_Tarjeta,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,CEM.Empresa_ID,CEM.Nombre AS Nombre_Empresa,CE.Fecha_Ingreso"
    Mi_SQL = Mi_SQL & " ,CP.Nombre AS Puesto,CD.Nombre AS Departamento,CE.Estatus"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Empresas CEM,Cat_Puestos CP,Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE CE.Empresa_ID=CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Puesto_ID=CP.Puesto_ID"
    Mi_SQL = Mi_SQL & " AND CD.Departamento_ID=CE.Departamento_ID"
    'Mi_SQL = Mi_SQL & " AND CE.Estatus='A'"
    'Validacion de Empresa
    If Cmb_Rpt_Empleados_Alta_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Rpt_Empleados_Alta_Empresa.ItemData(Cmb_Rpt_Empleados_Alta_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Departamentos
    If Cmb_Rpt_Empleados_Alta_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Rpt_Empleados_Alta_Departamento.ItemData(Cmb_Rpt_Empleados_Alta_Departamento.ListIndex), "00000") & "'"
    End If
    'Validacion de Puesto
    If Cmb_Rpt_Empleados_Alta_Puesto.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Puesto_ID='" & Format(Cmb_Rpt_Empleados_Alta_Puesto.ItemData(Cmb_Rpt_Empleados_Alta_Puesto.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND CE.Fecha_Ingreso BETWEEN '" & Format(Dtp_Rpt_Empleados_Alta_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Empleados_Alta_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY CEM.Empresa_ID,CE.Fecha_Ingreso,CE.No_Tarjeta,CD.Nombre,CP.Nombre,CE.Apellido_Paterno"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Empresa_ID_Reporte = ""
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE EMPLEADOS", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            While Not .EOF
                If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                    Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                    Print #1,
                    Print #2,
                    Print #1, "Empresa: "; .rdoColumns("Nombre_Empresa")
                    Print #2, "Empresa:|"; .rdoColumns("Nombre_Empresa")
                    Print #1,
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " No.Nomina    Nombre                                          Departamento              Puesto            Ingreso  Estatus"
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #2, "No. Nomina|Nombre|Departamento|Puesto|F. Ingreso|Estatus"
                End If
                Print #1, Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Tarjeta"), 10); _
                    Spc(2); Mid(.rdoColumns("Nombre"), 1, 48); _
                    Spc(50 - Len(Mid(.rdoColumns("Nombre"), 1, 48))); Mid(.rdoColumns("Departamento"), 1, 20); _
                    Spc(22 - Len(Mid(.rdoColumns("Departamento"), 1, 20))); Mid(.rdoColumns("Puesto"), 1, 20); _
                    Spc(22 - Len(Mid(.rdoColumns("Puesto"), 1, 20))); Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                    Spc(2); .rdoColumns("Estatus")
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Departamento"); _
                    "|"; .rdoColumns("Puesto"); _
                    "|"; Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                    "|"; .rdoColumns("Estatus")
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleados_Huella_Comedor
'DESCRIPCION: Genera el reporte de los empleados que ya le registraron la huella de comedor
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 29-Mayo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleados_Huella_Comedor()
Dim Rs_Consulta_Cat_Empleados As rdoResultset               'Información de inasistencias o faltas
Dim Mi_SQL As String                                       'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String
Dim Empleado_ID_Reporte As String                          'ID de la empresa
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT CE.No_Tarjeta,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,CEM.Empresa_ID,CEM.Nombre AS Nombre_Empresa,CE.Fecha_Ingreso"
    Mi_SQL = Mi_SQL & " ,CP.Nombre AS Puesto,CD.Nombre AS Departamento,CE.Estatus,Cat_Empleados_Huellas.Huella_Ruta"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Empleados_Huellas,Cat_Empresas CEM,Cat_Puestos CP,Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE CE.Empleado_ID=Cat_Empleados_Huellas.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID=CEM.Empresa_ID"
    Mi_SQL = Mi_SQL & " AND CE.Puesto_ID=CP.Puesto_ID"
    Mi_SQL = Mi_SQL & " AND CD.Departamento_ID=CE.Departamento_ID"
    'Validacion de Empresa
    If Cmb_Rpt_Empleados_Alta_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Rpt_Empleados_Alta_Empresa.ItemData(Cmb_Rpt_Empleados_Alta_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Departamentos
    If Cmb_Rpt_Empleados_Alta_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Rpt_Empleados_Alta_Departamento.ItemData(Cmb_Rpt_Empleados_Alta_Departamento.ListIndex), "00000") & "'"
    End If
    'Validacion de Puesto
    If Cmb_Rpt_Empleados_Alta_Puesto.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Puesto_ID='" & Format(Cmb_Rpt_Empleados_Alta_Puesto.ItemData(Cmb_Rpt_Empleados_Alta_Puesto.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND CE.Fecha_Ingreso BETWEEN '" & Format(Dtp_Rpt_Empleados_Alta_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Empleados_Alta_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY CE.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Empresa_ID_Reporte = ""
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE EMPLEADOS CON HUELLA REGISTRADA", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            While Not .EOF
                'Valida exista el archivo físico
                If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Huellas\" & .rdoColumns("Huella_Ruta"), "ARCHIVO") = True Then
                    If .rdoColumns("Empresa_ID") <> Empresa_ID_Reporte Then
                        Empresa_ID_Reporte = .rdoColumns("Empresa_ID")
                        Print #1,
                        Print #2,
                        Print #1, "Empresa: "; .rdoColumns("Nombre_Empresa")
                        Print #2, "Empresa:|"; .rdoColumns("Nombre_Empresa")
                        Print #1,
                        Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                        Print #1, " No.Nomina    Nombre                                          Departamento              Puesto            Ingreso  Estatus"
                        Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                        Print #2, "No. Nomina|Nombre|Departamento|Puesto|F. Ingreso|Estatus"
                    End If
                    Print #1, Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Tarjeta"), 10); _
                        Spc(2); Mid(.rdoColumns("Nombre"), 1, 48); _
                        Spc(50 - Len(Mid(.rdoColumns("Nombre"), 1, 48))); Mid(.rdoColumns("Departamento"), 1, 20); _
                        Spc(22 - Len(Mid(.rdoColumns("Departamento"), 1, 20))); Mid(.rdoColumns("Puesto"), 1, 20); _
                        Spc(22 - Len(Mid(.rdoColumns("Puesto"), 1, 20))); Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                        Spc(2); .rdoColumns("Estatus")
                    Print #2, .rdoColumns("No_Tarjeta"); _
                        "|"; .rdoColumns("Nombre"); _
                        "|"; .rdoColumns("Departamento"); _
                        "|"; .rdoColumns("Puesto"); _
                        "|"; Format(.rdoColumns("Fecha_Ingreso"), "dd/MMM/yyyy"); _
                        "|"; .rdoColumns("Estatus")
                End If
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleado_Cursos
'DESCRIPCION: Genera el reporte de los empleados que tomaron el curso seleccionado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 10-Mayo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleado_Cursos()
Dim Rs_Consulta_Cat_Empleados As rdoResultset
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT Cat_Empleados.No_Tarjeta,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,Cat_Cursos.Nombre AS Curso,Cat_Cursos.Tipo,Cat_Cursos.Horas"
    Mi_SQL = Mi_SQL & " ,Cat_Cursos_Detalles.Comentarios AS Instructor,Cat_Cursos_Detalles.Estatus,Cat_Cursos_Detalles.Fecha_Inicio,Cat_Cursos_Detalles.Fecha_Fin"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados,Cat_Cursos,Cat_Cursos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Empleado_ID=Cat_Cursos_Detalles.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Cursos.Curso_ID=Cat_Cursos_Detalles.Curso_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus='A'"
    'Validacion de Curso
    If Cmb_Curso.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos.Curso_ID='" & Format(Cmb_Curso.ItemData(Cmb_Curso.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Empleado_Curso.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Empleado_Curso.ItemData(Cmb_Empleado_Curso.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Fechas_Curso.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Detalles.Fecha_Inicio BETWEEN '" & Format(Dtp_Fecha_Inicio_Curso.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Curso.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Cursos.Nombre,Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE EMPLEADOS POR CURSO", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            Print #1,
            Print #1, "Empleado: "; .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre")
            Print #1,
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "     Curso                               Tipo      Horas   Instructor                  Inicio       Fin        Estatus    "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "No. Nomina|Nombre|Curso|Tipo|Horas|Instructor|Inicio|Fin|Estatus"
            While Not .EOF
                Print #1, Mid(.rdoColumns("Curso"), 1, 35); _
                    Spc(37 - Len(Mid(.rdoColumns("Curso"), 1, 35))); Mid(.rdoColumns("Tipo"), 1, 10); _
                    Spc(12 - Len(Mid(.rdoColumns("Tipo"), 1, 10))); Conectar_Ayudante.Alinea_Derecha(.rdoColumns("Horas"), 6); _
                    Spc(2); Mid(.rdoColumns("Instructor"), 1, 25); _
                    Spc(27 - Len(Mid(.rdoColumns("Instructor"), 1, 25))); Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy"); _
                    Spc(2); .rdoColumns("Estatus")
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Curso"); _
                    "|"; .rdoColumns("Tipo"); _
                    "|"; .rdoColumns("Horas"); _
                    "|"; .rdoColumns("Instructor"); _
                    "|"; Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy"); _
                    "|"; .rdoColumns("Estatus")
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Empleado_Comidas
'DESCRIPCION: Genera el reporte de comidas por empleado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 04-Abril-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Empleado_Comidas()
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Total_Comidas As Long
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT Cat_Empleados.No_Tarjeta,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,Adm_Entradas_Comedor.Fecha,Adm_Entradas_Comedor.Hora"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados,Adm_Entradas_Comedor"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Empleado_ID=Adm_Entradas_Comedor.Empleado_ID"
    'Validacion de Empleado
    If Cmb_Empleado_Curso.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Empleado_Curso.ItemData(Cmb_Empleado_Curso.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Fechas_Curso.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Entradas_Comedor.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Curso.Value, "MM/dd/yyyy 00:00:00") & "' AND '" & Format(Dtp_Fecha_Fin_Curso.Value, "MM/dd/yyyy") & " 23:59:59'"
    End If
'    If Cmb_Estatus.ListIndex = 1 Then
'     Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus = 'A' "
'    End If
'    If Cmb_Estatus.ListIndex = 2 Then
'    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus = 'I' "
'    End If
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Entradas_Comedor.Fecha,Adm_Entradas_Comedor.Hora,Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Fechas_Curso.Value = 1 Then
                Call Encabezado_Reporte("REPORTE DE COMIDAS POR EMPLEADO DETALLADO", Dtp_Fecha_Inicio_Curso.Value, Dtp_Fecha_Fin_Curso.Value, False)
            Else
                Call Encabezado_Reporte("REPORTE DE COMIDAS POR EMPLEADO DETALLADO (TODAS)", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1,
            Print #1, "Empleado: "; .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre")
            Print #1,
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "     Fecha       Hora      "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "No.Nomina|Nombre|Fecha|Hora"
            While Not .EOF
                Print #1, Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Hora"), "HH:mm:ss")
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Hora"), "dd/MMM/yyyy")
                Total_Comidas = Total_Comidas + 1
                .MoveNext
            Wend
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, " TOTALES" & Conectar_Ayudante.Alinea_Derecha(Format(Total_Comidas, "#,##0"), 12)
            Print #2, "||Totales|" & Format(Total_Comidas, "#,##0")
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Entradas_Comedor", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Curso_Empleados
'DESCRIPCION: Genera el reporte del os cursos que fueron tomaron el curso seleccionado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 10-Mayo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Curso_Empleados()
Dim Rs_Consulta_Cat_Empleados As rdoResultset
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT Cat_Empleados.No_Tarjeta,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,Cat_Cursos.Nombre AS Curso,Cat_Cursos.Tipo,Cat_Cursos.Horas"
    Mi_SQL = Mi_SQL & " ,Cat_Cursos_Detalles.Comentarios AS Instructor,Cat_Cursos_Detalles.Estatus,Cat_Cursos_Detalles.Fecha_Inicio,Cat_Cursos_Detalles.Fecha_Fin"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados,Cat_Cursos,Cat_Cursos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Empleado_ID=Cat_Cursos_Detalles.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Cursos.Curso_ID=Cat_Cursos_Detalles.Curso_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus='A'"
    'Validacion de Curso
    If Cmb_Curso.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos.Curso_ID='" & Format(Cmb_Curso.ItemData(Cmb_Curso.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Empleado_Curso.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Empleado_Curso.ItemData(Cmb_Empleado_Curso.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Fechas_Curso.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Detalles.Fecha_Inicio BETWEEN '" & Format(Dtp_Fecha_Inicio_Curso.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Curso.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Cursos.Nombre,Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE EMPLEADOS POR CURSO", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            Print #1,
            Print #1, "Curso: "; .rdoColumns("Curso")
            Print #1, "Tipo : "; .rdoColumns("Tipo"); "     Horas : "; .rdoColumns("Horas")
            Print #1,
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "     Empleado                                              Instructor                  Inicio       Fin        Estatus    "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "No. Nomina|Nombre|Curso|Tipo|Horas|Instructor|Inicio|Fin|Estatus"
            While Not .EOF
                Print #1, Alinea_Derecha(.rdoColumns("No_Tarjeta"), 8); _
                    Spc(2); Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Tipo"), 1, 40))); Mid(.rdoColumns("Instructor"), 1, 25); _
                    Spc(27 - Len(Mid(.rdoColumns("Instructor"), 1, 25))); Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    Spc(2); Format(.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy"); _
                    Spc(2); .rdoColumns("Estatus")
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Curso"); _
                    "|"; .rdoColumns("Tipo"); _
                    "|"; .rdoColumns("Horas"); _
                    "|"; .rdoColumns("Instructor"); _
                    "|"; Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                    "|"; Format(.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy"); _
                    "|"; .rdoColumns("Estatus")
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Comidas_Empleados
'DESCRIPCION: Genera el reporte de comidas de empleados
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 04-Abril-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Comidas_Empleados()
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Total_Comidas As Long
Dim Total_Empresa As Double
Dim Total_Empleado As Double
    
    'Consulta los empleados del sistema
    Mi_SQL = "SELECT Cat_Empleados.No_Tarjeta,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Nombre"
    Mi_SQL = Mi_SQL & " ,ISNULL(COUNT(Adm_Entradas_Comedor.Fecha),0) AS Entradas,Cat_Departamentos.Clave,Cat_Departamentos.Nombre AS Departamento"
    Mi_SQL = Mi_SQL & " FROM Adm_Entradas_Comedor,Cat_Empleados,Cat_Departamentos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Entradas_Comedor.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Departamento_ID=Cat_Departamentos.Departamento_ID"
    'Rango de Fechas
    If Chk_Fechas_Curso.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Entradas_Comedor.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Curso.Value, "MM/dd/yyyy 00:00:00") & "' AND '" & Format(Dtp_Fecha_Fin_Curso.Value, "MM/dd/yyyy") & " 23:59:59'"
    End If
    If Cmb_Estatus.ListIndex = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus = 'A'"
    End If
    If Cmb_Estatus.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus = 'I'"
    End If
    Mi_SQL = Mi_SQL & " GROUP BY Cat_Empleados.No_Tarjeta,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre),Cat_Departamentos.Clave,Cat_Departamentos.Nombre"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Fechas_Curso.Value = 1 Then
                Call Encabezado_Reporte("REPORTE DE COMIDAS POR EMPLEADO", Dtp_Fecha_Inicio_Curso.Value, Dtp_Fecha_Fin_Curso.Value, False)
            Else
                Call Encabezado_Reporte("REPORTE DE COMIDAS POR EMPLEADO (TODAS)", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1,
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No.      Empleado                                        Centro Costo                 Cantidad  C.Empresa   C.Empleado  "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "No.Nomina|Empleado|Clave|CentroCosto|Cantidad|CostoEmpresa|CostoEmpleado"
            While Not .EOF
                Print #1, Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Tarjeta"), 8); _
                    Spc(2); Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(43 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); Mid(.rdoColumns("Clave"), 1, 5); _
                    Spc(7 - Len(Mid(.rdoColumns("Clave"), 1, 5))); Mid(.rdoColumns("Departamento"), 1, 20); _
                    Spc(22 - Len(Mid(.rdoColumns("Departamento"), 1, 20))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Entradas"), "#,##0"), 10); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Entradas") * PG_Costo_Comida_Empresa, "#,##0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Entradas") * PG_Costo_Comida_Empleado, "#,##0.00"), 12)
                Print #2, .rdoColumns("No_Tarjeta"); _
                    "|"; .rdoColumns("Nombre"); _
                    "|"; .rdoColumns("Clave"); _
                    "|"; .rdoColumns("Departamento"); _
                    "|"; Format(.rdoColumns("Entradas"), "#,##0"); _
                    "|"; Format(.rdoColumns("Entradas") * PG_Costo_Comida_Empresa, "#,##0.00"); _
                    "|"; Format(.rdoColumns("Entradas") * PG_Costo_Comida_Empleado, "#,##0.00")
                Total_Comidas = Total_Comidas + Val(.rdoColumns("Entradas"))
                Total_Empresa = Total_Empresa + Val(.rdoColumns("Entradas") * PG_Costo_Comida_Empresa)
                Total_Empleado = Total_Empleado + Val(.rdoColumns("Entradas") * PG_Costo_Comida_Empleado)
                .MoveNext
            Wend
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(70); " TOTALES  " & Conectar_Ayudante.Alinea_Derecha(Format(Total_Comidas, "#,##0"), 12); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Empresa, "#,##0.00"), 12); Spc(2); Conectar_Ayudante.Alinea_Derecha(Format(Total_Empleado, "#,##0.00"), 12)
            Print #2, "|||TOTALES|" & Format(Total_Comidas, "#,##0") & "|" & Format(Total_Empresa, "#,##0.00"); "|" & Format(Total_Empleado, "#,##0.00")
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Entradas_Comedor", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Cat_Empleados.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Generar_Reporte_Cursos_Por_Empleado()
Dim Consulta_Cursos_Por_Empleado As rdoResultset
Dim Guardar_Nombre As String
Guardar_Nombre = ""

    'Consulta los empleados del sistema
    Mi_SQL = "SELECT (ce.Nombre + ' ' + ce.Apellido_Paterno + ' ' + ce.Apellido_Materno) AS Nombre_Empleado"
    Mi_SQL = Mi_SQL & " ,Opc.No_Programa_Curso,ccc.Clave,ccc.Nombre AS Nombre_Curso,CONVERT(DATE, opc.Fecha_Inicio) AS Fecha_Inicio"
    Mi_SQL = Mi_SQL & " ,CONVERT(DATE, opc.Fecha_Fin) AS Fecha_Fin,ccc.Horas,cs.Nombre AS Lugar,'SI' AS Invitado"
    Mi_SQL = Mi_SQL & " ,ISNUll((SELECT TOP 1 Archivo FROM Ope_Evaluaciones_Empleados pee Where pee.No_Programa_Curso = opc.No_Programa_Curso"
    Mi_SQL = Mi_SQL & " AND pee.Empleado_Id = ola.Empleado_Id), '') AS Archivo"
    Mi_SQL = Mi_SQL & " from Ope_Programacion_Cursos opc, Ope_Lista_Asistencia ola, Cat_Empleados ce,"
    Mi_SQL = Mi_SQL & " Cat_Cursos_Capacitaciones ccc, Cat_Tipos_Cursos ctc, Cat_Salas cs"
    Mi_SQL = Mi_SQL & " Where opc.No_Programa_Curso = ola.No_Programa_Curso and ola.Empleado_Id = ce.Empleado_ID and ccc.Curso_ID = opc.Curso_Id"
    Mi_SQL = Mi_SQL & " and ccc.Tipo_Curso_Id = ctc.Tipo_Curso_Id and cs.Sala_Id = opc.Sala_Id"
    'Validacion de Empleado
    If Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND ce.Empleado_ID='" & Format(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ItemData(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Rpt_Cursos_Tomados_Por_Empleado_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND opc.Fecha_Inicio >= '" & Format(Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Inicio.Value, "yyyy/MM/dd") & "' AND opc.Fecha_Fin <= '" & Format(Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Fin.Value, "yyyy/MM/dd") & "'"
    End If
    Mi_SQL = Mi_SQL & " GROUP BY ccc.Clave,ccc.Nombre,ce.Nombre,ce.Apellido_Paterno,ce.Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,opc.Fecha_Inicio,opc.Fecha_Fin,ccc.Horas,cs.Nombre,ctc.Nombre,opc.No_Programa_Curso"
    Mi_SQL = Mi_SQL & " ,ola.Empleado_ID,opc.Hora_Inicio,opc.Hora_Fin,ola.No_Programa_Curso"
    Mi_SQL = Mi_SQL & " ORDER BY ce.Nombre,ce.Apellido_Paterno,ce.Apellido_Materno,ccc.Nombre,opc.Fecha_Inicio"
    Set Consulta_Cursos_Por_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Consulta_Cursos_Por_Empleado.EOF Then
        With Consulta_Cursos_Por_Empleado
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Rpt_Cursos_Tomados_Por_Empleado_Fechas.Value = 1 Then
                Call Encabezado_Reporte("CURSOS TOMADOS POR EMPLEADO", Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Fin.Value, False)
            Else
                Call Encabezado_Reporte("CURSOS TOMADOS POR EMPLEADO", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1, "Registro Patronal: B4744175109"
            Print #2, "Registro Patronal: B4744175109"
            Print #1, "RFC: SMG100824LY0"
            Print #2, "RFC: SMG100824LY0"
            Print #1,
            Print #2,
            Print #1, "Programación   Curso          Nombre del Curso           Inicio       Fin      Horas    Lugar  Asist.   Calificación "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "Programación|Curso|Nombre del Curso|Inicio|Fin|Horas|Lugar|Asistencia|Calificación"
            Print #2, "--------------------------------------------------------------------------------------------------------------------------"
        
        While Not .EOF
            If Guardar_Nombre <> .rdoColumns("Nombre_Empleado") Then
                Print #1,
                Print #2,
                Print #1, .rdoColumns("Nombre_Empleado")
                Print #2, .rdoColumns("Nombre_Empleado")
                Guardar_Nombre = .rdoColumns("Nombre_Empleado")
            End If
            Print #1, Mid(Format(.rdoColumns("No_Programa_Curso"), "0000000000"), 1, 10); Spc(3); _
                Mid(.rdoColumns("Clave"), 1, 8); Spc(10 - Len(Mid(.rdoColumns("Clave"), 1, 8))); _
                Mid(.rdoColumns("Nombre_Curso"), 1, 30); Spc(32 - Len(Mid(.rdoColumns("Nombre_Curso"), 1, 30))); _
                Mid(Format(.rdoColumns("Fecha_Inicio"), "dd/MM/yyyy"), 1, 10); Spc(2); _
                Mid(Format(.rdoColumns("Fecha_Fin"), "dd/MM/yyyy"), 1, 10); Spc(2); _
                Mid(Format(.rdoColumns("Horas"), "00.00"), 1, 6); Spc(2); _
                Mid(.rdoColumns("Lugar"), 1, 8); Spc(10 - Len(Mid(.rdoColumns("Lugar"), 1, 8))); _
                Mid(.rdoColumns("Invitado"), 1, 4); Spc(6 - Len(Mid(.rdoColumns("Invitado"), 1, 4))); _
                Mid(.rdoColumns("Archivo"), 1, 8)
            Print #2, .rdoColumns("No_Programa_Curso"); _
                "|"; .rdoColumns("Clave"); _
                "|"; .rdoColumns("Nombre_Curso"); _
                "|"; Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy"); _
                "|"; Format(.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy"); _
                "|"; Format(.rdoColumns("Horas"), "00.00"); _
                "|"; .rdoColumns("Lugar"); _
                "|"; .rdoColumns("Invitado"); _
                "|"; .rdoColumns("Archivo")
            .MoveNext
        Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Por_Empleado", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Consulta_Cursos_Por_Empleado.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Generar_Reporte_Cursos_Hora_Hombre()
Dim Consulta_Cursos_Hora_Hombre As rdoResultset
    
     If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = -1 Then
        Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0
    End If
    'Validación para cuando elige algun elemento del combo tipo de busqueda
    'Busqueda por Empleado
    If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 1 Or Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0 Then
        Mi_SQL = "SELECT DISTINCT (Ope_Programacion_Cursos.No_Programa_Curso) AS Programacion"
        Mi_SQL = Mi_SQL & " ,(cast(Cat_Empleados.No_Tarjeta AS VARCHAR) + ' ' + Cat_Empleados.Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno) AS Empleado"
        Mi_SQL = Mi_SQL & " ,Tipo_Empleado,Cat_Cursos_Capacitaciones.Clave,Cat_Cursos_Capacitaciones.Nombre AS Nombre_Curso"
        Mi_SQL = Mi_SQL & " ,CONVERT(DATE, Ope_Lista_Asistencia.Fecha_Hora_Registro) AS Fecha_Hora_Registro"
        Mi_SQL = Mi_SQL & " ,Cat_Salas.Nombre AS Sala,'1' AS Asistentes,CAST(((DATEDIFF(MI, Ope_Programacion_Cursos.Hora_Inicio, "
        Mi_SQL = Mi_SQL & " Ope_Programacion_Cursos.Hora_Fin) * COUNT(*) * 1.00) / 60.0) AS DECIMAL(5, 2)) AS Horas_Hombre"
        Mi_SQL = Mi_SQL & " from Cat_Empleados,Ope_Programacion_Cursos,Ope_Lista_Asistencia,Cat_Cursos_Capacitaciones,Cat_Salas"
        Mi_SQL = Mi_SQL & " Where Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso"
        Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Empleado_Id = Cat_Empleados.Empleado_ID"
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id"
        Mi_SQL = Mi_SQL & " AND Cat_Salas.Sala_Id = Ope_Programacion_Cursos.Sala_Id"
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus = 'A'"
        If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID = " & Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex) & ""
        ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0 Then
            If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex > 0 Then
                Mi_SQL = Mi_SQL & " AND Cat_Empleados.Tipo_Empleado like '%" & Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Text & "%' "
            End If
        End If
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND (Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, "MM/dd/yyyy") & "')"
        End If
        Mi_SQL = Mi_SQL & " group by Ope_Programacion_Cursos.No_Programa_Curso,Cat_Empleados.No_Tarjeta,Cat_Empleados.Nombre"
        Mi_SQL = Mi_SQL & " ,Apellido_Paterno,Apellido_Materno,Tipo_Empleado,Cat_Cursos_Capacitaciones.Clave,Cat_Cursos_Capacitaciones.Nombre"
        Mi_SQL = Mi_SQL & " ,Fecha_Hora_Registro,Cat_Salas.Nombre,Hora_Inicio,Hora_Fin"
        Mi_SQL = Mi_SQL & " order by cast(Cat_Empleados.No_Tarjeta AS VARCHAR) + ' ' + Cat_Empleados.Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno,"
        Mi_SQL = Mi_SQL & " Tipo_Empleado,Ope_Programacion_Cursos.No_Programa_Curso,Fecha_Hora_Registro"
    'Busqueda por curso
    ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 2 Then
        Mi_SQL = "select ccc.Nombre as Nombre_Curso, ctc.Nombre as Tipo_Curso, opc.Fecha_Inicio,"
        Mi_SQL = Mi_SQL & " opc.Fecha_Fin, (select COUNT(DISTINCT(ce.Empleado_ID)) from Cat_Empleados ce, "
        Mi_SQL = Mi_SQL & " Ope_Lista_Asistencia olas Where ce.Empleado_ID = olas.Empleado_ID"
        Mi_SQL = Mi_SQL & " and olas.No_Programa_Curso = opc.No_Programa_Curso"
        Mi_SQL = Mi_SQL & " and ce.Tipo_Empleado = 'sindicalizado') as Sind, (select COUNT(DISTINCT(ce.Empleado_ID))"
        Mi_SQL = Mi_SQL & " from Cat_Empleados ce, Ope_Lista_Asistencia olas"
        Mi_SQL = Mi_SQL & " Where ce.Empleado_ID = olas.Empleado_ID and olas.No_Programa_Curso = opc.No_Programa_Curso"
        Mi_SQL = Mi_SQL & " and ce.Tipo_Empleado != 'sindicalizado') as No_Sind,"
        Mi_SQL = Mi_SQL & " COUNT(DISTINCT(ola.Empleado_Id)) as Total_Asist, ccc.Horas,"
        Mi_SQL = Mi_SQL & " CAST(((DATEDIFF(MI, opc.Hora_Inicio, opc.Hora_Fin) * COUNT(*) * 1.00) / 60.0) as decimal(5,2)) as Horas_Hombre"
        Mi_SQL = Mi_SQL & " from Ope_Programacion_Cursos opc, Cat_Cursos_Capacitaciones ccc, Cat_Tipos_Cursos ctc,"
        Mi_SQL = Mi_SQL & " Ope_Lista_Asistencia ola, Cat_Empleados cee Where opc.Curso_ID = ccc.Curso_ID"
        Mi_SQL = Mi_SQL & " and ccc.Tipo_Curso_Id = ctc.Tipo_Curso_Id"
        Mi_SQL = Mi_SQL & " and opc.No_Programa_Curso = ola.No_Programa_Curso"
        Mi_SQL = Mi_SQL & " and cee.Empleado_ID = ola.Empleado_Id"
        If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND ccc.Curso_ID = '" & Format(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex), "00000") & "'"
        End If
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND opc.Fecha_Inicio>='" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND opc.Fecha_Fin<='" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
        End If
        Mi_SQL = Mi_SQL & " group by opc.No_Programa_Curso, ccc.Nombre, ctc.Nombre, opc.Fecha_Inicio, "
        Mi_SQL = Mi_SQL & " opc.Fecha_Fin, ccc.Horas , opc.Hora_Inicio, opc.Hora_Fin"
        Mi_SQL = Mi_SQL & " order by ccc.Nombre, opc.Fecha_Inicio"
    'Busqueda por departamento
    ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 3 Then
        Mi_SQL = "select Cat_Departamentos.Clave AS Clave_Departamento, Cat_Departamentos.Nombre as Nombre_Departamento, "
        Mi_SQL = Mi_SQL & " Cat_Cursos_Capacitaciones.Clave as Clave_Curso, Cat_Cursos_Capacitaciones.Nombre as Nombre_Curso,"
        Mi_SQL = Mi_SQL & " CONVERT(DATE,Ope_Programacion_Cursos.Fecha_Inicio) AS Inicio, CONVERT(DATE,Ope_Programacion_Cursos.Fecha_Fin) AS Fin,"
        Mi_SQL = Mi_SQL & " CAST(((DATEDIFF(MI, Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin) * COUNT(*) * 1.00) / 60.0) as decimal(5,2)) as Duracion,"
        Mi_SQL = Mi_SQL & " CAST(Cat_Empleados.No_Tarjeta AS Varchar) as No_Tarjeta, Cat_Empleados.Nombre+' '  + Cat_Empleados.Apellido_Paterno+' ' + Cat_Empleados.Apellido_Materno as Empleado,"
        Mi_SQL = Mi_SQL & " Cat_Instructores.Nombre + ' ' + Cat_Instructores.Apellido_Paterno+ ' ' + Cat_Instructores.Apellido_Materno as Instructor"
        Mi_SQL = Mi_SQL & " from Cat_Departamentos, Cat_Cursos_Capacitaciones, Cat_Empleados, Ope_Lista_Asistencia, "
        Mi_SQL = Mi_SQL & " Ope_Programacion_Cursos , Cat_Instructores"
        Mi_SQL = Mi_SQL & " where Cat_Departamentos.Departamento_ID = Cat_Empleados.Departamento_ID"
        Mi_SQL = Mi_SQL & " and Cat_Empleados.Empleado_ID = Ope_Lista_Asistencia.Empleado_Id"
        Mi_SQL = Mi_SQL & " and Cat_Instructores.Instructor_Id = Ope_Programacion_Cursos.Instructor_Id"
        Mi_SQL = Mi_SQL & " and Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso"
        Mi_SQL = Mi_SQL & " and Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id"
        If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " and Cat_Departamentos.Departamento_ID = '" & Format(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex), "00000") & "'"
        End If
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND (Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, "MM/dd/yyyy") & "')"
        End If
        Mi_SQL = Mi_SQL & " group BY Cat_Departamentos.Clave, Cat_Departamentos.Nombre, Cat_Cursos_Capacitaciones.Clave,"
        Mi_SQL = Mi_SQL & " Cat_Cursos_Capacitaciones.Nombre,Ope_Programacion_Cursos.Fecha_Inicio, Ope_Programacion_Cursos.Fecha_Fin,"
        Mi_SQL = Mi_SQL & " Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin, Cat_Empleados.No_Tarjeta,"
        Mi_SQL = Mi_SQL & " Cat_Empleados.Nombre, Cat_Empleados.Apellido_Paterno, Cat_Empleados.Apellido_Materno,"
        Mi_SQL = Mi_SQL & " Cat_Instructores.Nombre , Cat_Instructores.Apellido_Paterno, Cat_Instructores.Apellido_Materno"
        Mi_SQL = Mi_SQL & " order By Cat_Departamentos.Clave, Cat_Cursos_Capacitaciones.Clave"
    End If

    Set Consulta_Cursos_Hora_Hombre = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Not Consulta_Cursos_Hora_Hombre.EOF Then
        If Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 1 Or Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 0 Then
            Call Formato_Reporte_Cursos_Hora_Hombre_Por_Empleado(Consulta_Cursos_Hora_Hombre)
        ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 2 Then
            Call Formato_Reporte_Cursos_Hora_Hombre_Por_Curso(Consulta_Cursos_Hora_Hombre)
        ElseIf Cmb_Rpt_Cursos_Hora_Hombre_Tipo_De_Busqueda.ListIndex = 3 Then
            Call Formato_Reporte_Cursos_Hora_Hombre_Por_Departamento(Consulta_Cursos_Hora_Hombre)
        End If
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Consulta_Cursos_Hora_Hombre.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Formato_Reporte_Cursos_Hora_Hombre_Por_Empleado(Consulta_Cursos_Hora_Hombre As rdoResultset)
Dim Guardar_Nombre As String
Dim Suma_Horas As Double

    With Consulta_Cursos_Hora_Hombre
        MDIFrm_Apl_Principal.MousePointer = 11
        'Agrega el encabezado al reporte
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, False)
        Else
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
        End If
        Print #1, "Registro Patronal: B4744175109"
        Print #2, "Registro Patronal: B4744175109"
        Print #1, "RFC: SMG100824LY0"
        Print #2, "RFC: SMG100824LY0"
        Print #1,
        Print #2,
        Print #1, "Programación      Curso                 Nombre del Curso                          Fecha      Sala    Asist.  Horas/Hombre "
        Print #1, "--------------------------------------------------------------------------------------------------------------------------"
        Print #2, "Programación|Curso|Nombre del Curso|Fecha|Asistentes|Horas/Hombre"
        Print #2, "--------------------------------------------------------------------------------------------------------------------------"
        
        While Not .EOF
            
            If Guardar_Nombre <> .rdoColumns("Empleado") And (Guardar_Nombre <> "") Then
                Print #1, "                                                                                                         _________________"
                Print #1, "                                                                                                                " & Format(Suma_Horas, "00.00")
                Print #2, "|||||||____________________"
                Print #2, "|||||||          " & Suma_Horas
                Suma_Horas = 0
            End If
            
            If Guardar_Nombre <> .rdoColumns("Empleado") Then
                Print #1, .rdoColumns("Empleado"); Spc(6); .rdoColumns("Tipo_Empleado")
                Print #2, .rdoColumns("Empleado"); "|"; .rdoColumns("Tipo_Empleado")
                Guardar_Nombre = .rdoColumns("Empleado")
            End If
        
            Print #1, Mid(Format(.rdoColumns("Programacion"), "0000000000"), 1, 10); Spc(4); _
                Mid(.rdoColumns("Clave"), 1, 12); Spc(14 - Len(Mid(.rdoColumns("Clave"), 1, 12))); _
                Mid(.rdoColumns("Nombre_Curso"), 1, 49); Spc(51 - Len(Mid(.rdoColumns("Nombre_Curso"), 1, 49))); _
                Mid(Format(.rdoColumns("Fecha_Hora_Registro"), "dd/MM/yyyy"), 1, 10); Spc(2); _
                Mid(.rdoColumns("Sala"), 1, 10); Spc(12 - Len(Mid(.rdoColumns("Sala"), 1, 10))); _
                Mid(.rdoColumns("Asistentes"), 1, 7); Spc(9 - Len(Mid(.rdoColumns("Asistentes"), 1, 7))); _
                Mid(Format(.rdoColumns("Horas_Hombre"), "00.00"), 1, 6)
            Print #2, Format(.rdoColumns("Programacion"), "0000000000"); "|"; _
                .rdoColumns("Clave"); "|"; _
                .rdoColumns("Nombre_Curso"); "|"; _
                Format(.rdoColumns("Fecha_Hora_Registro"), "dd/MM/yyyy"); "|"; _
                .rdoColumns("Sala"); "|"; _
                .rdoColumns("Asistentes"); "|"; _
                .rdoColumns("Horas_Hombre")
            Suma_Horas = Suma_Horas + .rdoColumns("Horas_Hombre")
            
            .MoveNext
        Wend
                
        If Suma_Horas <> 0 Then
            Print #1, "                                                                                                         _________________"
            Print #1, "                                                                                                                " & Format(Suma_Horas, "00.00")
            Print #2, "|||||||____________________"
            Print #2, "|||||||          " & Suma_Horas
            Suma_Horas = 0
        End If
        
        Call Finalizar_Reporte(True)
        Btn_Imprimir.Enabled = True
        Btn_Exportar.Enabled = True
        Btn_Regresar.Enabled = True
        Btn_Salir.Enabled = True
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Hora_Hombre", Me)
    End With
End Sub

Private Sub Formato_Reporte_Cursos_Hora_Hombre_Por_Curso(Consulta_Cursos_Hora_Hombre As rdoResultset)
    With Consulta_Cursos_Hora_Hombre
        MDIFrm_Apl_Principal.MousePointer = 11
        'Agrega el encabezado al reporte
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, False)
        Else
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
        End If
        Print #1, "Registro Patronal: B4744175109"
        Print #2, "Registro Patronal: B4744175109"
        Print #1, "RFC: SMG100824LY0"
        Print #2, "RFC: SMG100824LY0"
        Print #1,
        Print #2,
        Print #1, "  Nombre del Curso          Tipo de Curso         Inicio        Fin         No Asistencia        Duración     Número de   "
        Print #1, "                                                                         Sind  No Sind  Total     (Horas)    Horas/Hombre "
        Print #1, "--------------------------------------------------------------------------------------------------------------------------"
        Print #2, "Nombre del Curso|Tipo de Curso|Inicio|Fin|No Asistencia||Duración (Horas)|Número de Horas/Hombre"
        Print #2, "||||Sind|No Sind|Total||"
        Print #2, "--------------------------------------------------------------------------------------------------------------------------"
        
        While Not .EOF
            Print #1, Mid(.rdoColumns("Nombre_Curso"), 1, 26); Spc(28 - Len(Mid(.rdoColumns("Nombre_Curso"), 1, 26))); _
                Mid(.rdoColumns("Tipo_Curso"), 1, 18); Spc(20 - Len(Mid(.rdoColumns("Tipo_Curso"), 1, 18))); _
                Mid(Format(.rdoColumns("Fecha_Inicio"), "dd/MM/yyyy"), 1, 12); Spc(2); _
                Mid(Format(.rdoColumns("Fecha_Fin"), "dd/MM/yyyy"), 1, 12); Spc(5); _
                Mid(.rdoColumns("Sind"), 1, 4); Spc(7 - Len(Mid(.rdoColumns("Sind"), 1, 4))); _
                Mid(.rdoColumns("No_Sind"), 1, 4); Spc(8 - Len(Mid(.rdoColumns("No_Sind"), 1, 4))); _
                Mid(.rdoColumns("Total_Asist"), 1, 4); Spc(9 - Len(Mid(.rdoColumns("Total_Asist"), 1, 4))); _
                Mid(Format(.rdoColumns("Horas"), "00.00"), 1, 6); Spc(8); _
                Mid(Format(.rdoColumns("Horas_Hombre"), "00.00"), 1, 6)
            Print #2, .rdoColumns("Nombre_Curso"); "|"; _
                .rdoColumns("Tipo_Curso"); "|"; _
                Format(.rdoColumns("Fecha_Inicio"), "dd/MM/yyyy"); "|"; _
                Format(.rdoColumns("Fecha_Fin"), "dd/MM/yyyy"); "|"; _
                .rdoColumns("Sind"); "|"; _
                .rdoColumns("No_Sind"); "|"; _
                .rdoColumns("Total_Asist"); "|"; _
                Format(.rdoColumns("Horas"), "00.00"); "|"; _
                Format(.rdoColumns("Horas_Hombre"), "00.00")
            .MoveNext
        Wend
        Call Finalizar_Reporte(True)
        Btn_Imprimir.Enabled = True
        Btn_Exportar.Enabled = True
        Btn_Regresar.Enabled = True
        Btn_Salir.Enabled = True
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Hora_Hombre", Me)
    End With
End Sub

Private Sub Formato_Reporte_Cursos_Hora_Hombre_Por_Departamento(Consulta_Cursos_Hora_Hombre As rdoResultset)
Dim Guardar_Nombre_Curso As String
Dim Guardar_Nombre_Departamento As String
Guardar_Nombre_Curso = ""
Guardar_Nombre_Departamento = ""

    With Consulta_Cursos_Hora_Hombre
        MDIFrm_Apl_Principal.MousePointer = 11
        'Agrega el encabezado al reporte
        If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, False)
        Else
            Call Encabezado_Reporte("CURSOS HORA HOMBRE", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
        End If
        Print #1, "Registro Patronal: B4744175109"
        Print #2, "Registro Patronal: B4744175109"
        Print #1, "RFC: SMG100824LY0"
        Print #2, "RFC: SMG100824LY0"
        Print #1, "--------------------------------------------------------------------------------------------------------------------------"
        Print #2, "--------------------------------------------------------------------------------------------------------------------------"
        
        While Not .EOF
            If Guardar_Nombre_Departamento <> .rdoColumns("Nombre_Departamento") Then
                Print #1,
                Print #2,
                Print #1, .rdoColumns("Clave_Departamento") & "  -  " & .rdoColumns("Nombre_Departamento")
                Print #2, .rdoColumns("Clave_Departamento") & "  -  " & .rdoColumns("Nombre_Departamento")
                Guardar_Nombre_Departamento = .rdoColumns("Nombre_Departamento")
            End If
            If Guardar_Nombre_Curso <> .rdoColumns("Nombre_Curso") Then
                Print #1,
                Print #2,
                Print #1, .rdoColumns("Clave_Curso") & "  -  " & .rdoColumns("Nombre_Curso")
                Print #2, .rdoColumns("Clave_Curso") & "  -  " & .rdoColumns("Nombre_Curso")
                Print #1, "                                             Inicio:" & .rdoColumns("Inicio") & "             Duración:" & .rdoColumns("Duracion")
                Print #2, "||Inicio:" & .rdoColumns("Inicio") & "|Duración:" & .rdoColumns("Duracion")
                Print #1, "       Código      Empleado                  Termino:" & .rdoColumns("Fin") & "             Instructor:" & .rdoColumns("Instructor")
                Print #2, "Código|Empleado|Termino:" & .rdoColumns("Fin") & "|Instructor:" & .rdoColumns("Instructor")
                Print #1, "     ----------------------------------------------------------------------------------------------------------------------"
                Print #2, "     ----------------------------------------------------------------------------------------------------------------------"
                Guardar_Nombre_Curso = .rdoColumns("Nombre_Curso")
            End If
            Print #1, Spc(8); Mid(.rdoColumns("No_Tarjeta"), 1, 8); Spc(10 - Len(Mid(.rdoColumns("No_Tarjeta"), 1, 8))); _
                .rdoColumns("Empleado")
            Print #2, .rdoColumns("No_Tarjeta"); "|"; _
                .rdoColumns("Empleado")
            .MoveNext
        Wend
        Call Finalizar_Reporte(True)
        Btn_Imprimir.Enabled = True
        Btn_Exportar.Enabled = True
        Btn_Regresar.Enabled = True
        Btn_Salir.Enabled = True
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Hora_Hombre", Me)
    End With
End Sub

Private Sub Generar_Reporte_Cursos_Indice_Asistencias()
Dim Consulta_Reporte_Indice_Asistencia As rdoResultset
Dim Guardar_No_Programa As String
Guardar_No_Programa = ""

    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = -1 Then
        Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 0
    End If
    
    Mi_SQL = " DECLARE @Empleado_Id AS CHAR (5)"
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
        Mi_SQL = Mi_SQL & " SET @Empleado_Id = '" & Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000") & "'"
    
    End If
    
    Mi_SQL = Mi_SQL & " DECLARE @Curso_Id AS CHAR (5)"
    
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " SET @Curso_Id = '" & Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " DECLARE @Departamento_Id AS CHAR (5) "
    
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
        Mi_SQL = Mi_SQL & " SET @Departamento_Id = '" & Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000") & "'"
    End If

    Mi_SQL = Mi_SQL & " SELECT Ope_Lista_Invitados.No_Programa_Curso, Ope_Lista_Invitados.Empleado_Id, Cat_Cursos_Capacitaciones.Clave, "
    Mi_SQL = Mi_SQL & "Cat_Cursos_Capacitaciones.Nombre AS Nombre_Curso, CONVERT(DATE, Ope_Programacion_Cursos.Fecha_Inicio) AS Fecha_Inicio, "
    Mi_SQL = Mi_SQL & "CONVERT(DATE, Ope_Programacion_Cursos.Fecha_Fin) AS Fecha_Final, CONVERT(TIME, Ope_Programacion_Cursos.Hora_Inicio) AS Hora_Inicial, "
    Mi_SQL = Mi_SQL & "CONVERT(TIME, Ope_Programacion_Cursos.Hora_Fin) AS Hora_Fin, Cat_Empresas.Acronimo, cast(No_Tarjeta AS VARCHAR) AS No_Tarjeta, "
    Mi_SQL = Mi_SQL & "Cat_Empleados.Nombre + ' ' + Cat_Empleados.Apellido_Paterno + ' ' + Cat_Empleados.Apellido_Materno AS Empleado, "
    Mi_SQL = Mi_SQL & "Cat_Puestos.Nombre, Cat_Empleados.Estatus, 'NO' AS Asistio, "
    Mi_SQL = Mi_SQL & "( SELECT count(Ope_Lista_Invitados.Empleado_id) From Ope_Lista_Invitados "
    Mi_SQL = Mi_SQL & " Where Ope_Lista_Invitados.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso"
    Mi_SQL = Mi_SQL & " AND Empleado_Id NOT IN ( SELECT Empleado_ID From Ope_Lista_Asistencia Where Ope_Lista_Asistencia.Empleado_ID = Ope_Lista_Invitados.Empleado_ID"
    Mi_SQL = Mi_SQL & ")) AS No_Asistieron, '0' AS Asistieron "
    Mi_SQL = Mi_SQL & " From Ope_Lista_Invitados, Cat_Cursos_Capacitaciones, Ope_Programacion_Cursos, Cat_Empleados, Cat_Empresas, Cat_Puestos "
    Mi_SQL = Mi_SQL & " Where Ope_Lista_Invitados.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id "
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Curso_ID = @Curso_Id "
    End If
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Departamento_ID = @Departamento_Id "
    End If
    Mi_SQL = Mi_SQL & " AND Ope_Lista_Invitados.Empleado_Id = Cat_Empleados.Empleado_ID AND Cat_Empresas.Empresa_ID = Cat_Empleados.Empresa_ID "
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Puesto_ID = Cat_Puestos.Puesto_ID "
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Lista_Invitados.Empleado_Id = @Empleado_Id "
    End If
    Mi_SQL = Mi_SQL & "  AND Ope_Lista_Invitados.Empleado_Id NOT IN ( SELECT Ope_Lista_Asistencia.Empleado_Id"
    Mi_SQL = Mi_SQL & "  From Ope_Lista_Asistencia "
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " where Cat_Cursos_Capacitaciones.Curso_ID = @Curso_Id "
    ElseIf Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
        Mi_SQL = Mi_SQL & " where Cat_Empleados.Departamento_ID = @Departamento_Id "
    ElseIf Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
         Mi_SQL = Mi_SQL & " where Ope_Lista_Asistencia.Empleado_Id = @Empleado_Id "
    Else
        Mi_SQL = Mi_SQL & " where Ope_Programacion_Cursos.Curso_Id LIKE '%%' "
    End If
    If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND (Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND '" & Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value, "MM/dd/yyyy") & "')"
    End If

    Mi_SQL = Mi_SQL & " ) UNION "

    Mi_SQL = Mi_SQL & " SELECT Ope_Lista_Asistencia.No_Programa_Curso, Ope_Lista_Asistencia.Empleado_Id, Cat_Cursos_Capacitaciones.Clave, "
    Mi_SQL = Mi_SQL & "Cat_Cursos_Capacitaciones.Nombre AS Nombre_Curso, CONVERT(DATE, Ope_Programacion_Cursos.Fecha_Inicio) AS Fecha_Inicio, "
    Mi_SQL = Mi_SQL & "CONVERT(DATE, Ope_Programacion_Cursos.Fecha_Fin) AS Fecha_Final, CONVERT(TIME, Ope_Programacion_Cursos.Hora_Inicio) AS Hora_Inicial, "
    Mi_SQL = Mi_SQL & "CONVERT(TIME, Ope_Programacion_Cursos.Hora_Fin) AS Hora_Fin, Cat_Empresas.Acronimo, cast(No_Tarjeta AS VARCHAR) AS No_Tarjeta, "
    Mi_SQL = Mi_SQL & "Cat_Empleados.Nombre + ' ' + Cat_Empleados.Apellido_Paterno + ' ' + Cat_Empleados.Apellido_Materno AS Empleado, "
    Mi_SQL = Mi_SQL & " Cat_Puestos.Nombre, Cat_Empleados.Estatus, 'SI' AS Asistio, '0' AS No_Asistieron, "
    Mi_SQL = Mi_SQL & " ( SELECT count(DISTINCT Ope_Lista_Asistencia.Empleado_Id) From Ope_Lista_Asistencia "
    Mi_SQL = Mi_SQL & "Where Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso ) AS Asistieron "
    Mi_SQL = Mi_SQL & "From Ope_Lista_Asistencia, Cat_Cursos_Capacitaciones, Ope_Programacion_Cursos, Cat_Empleados, Cat_Empresas, Cat_Puestos "
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " where Cat_Cursos_Capacitaciones.Curso_ID = @Curso_Id "
    ElseIf Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
        Mi_SQL = Mi_SQL & " where Cat_Empleados.Departamento_ID = @Departamento_Id "
    Else
        Mi_SQL = Mi_SQL & " where Ope_Programacion_Cursos.Curso_Id LIKE '%%' "
    End If
  
    Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id "
    Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Empleado_Id = Cat_Empleados.Empleado_ID "
    Mi_SQL = Mi_SQL & " AND Cat_Empresas.Empresa_ID = Cat_Empleados.Empresa_ID "
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Puesto_ID = Cat_Puestos.Puesto_ID "
    If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
         Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Empleado_Id = @Empleado_Id "
    End If
    If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND (Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND '" & Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value, "MM/dd/yyyy") & "')"
    End If

    Set Consulta_Reporte_Indice_Asistencia = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    If Not Consulta_Reporte_Indice_Asistencia.EOF Then
        With Consulta_Reporte_Indice_Asistencia
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
                Call Encabezado_Reporte("CURSOS INDICE DE ASISTENCIA", Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value, False)
            Else
                Call Encabezado_Reporte("CURSOS INDICE DE ASISTENCIA", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1,
            Print #2,
            Print #1, "Registro Patronal: B4744175109"
            Print #2, "Registro Patronal: B4744175109"
            Print #1, "RFC: SMG100824LY0"
            Print #2, "RFC: SMG100824LY0"
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "--------------------------------------------------------------------------------------------------------------------------"
            While Not .EOF
                If Guardar_No_Programa <> .rdoColumns("No_Programa_Curso") Then
                    Print #1,
                    Print #2,
                    Print #1, "Programación:   " & .rdoColumns("No_Programa_Curso")
                    Print #2, "Programación:" & .rdoColumns("No_Programa_Curso")
                    Print #1, "Curso:          " & .rdoColumns("Clave") & "         " & .rdoColumns("Nombre_Curso")
                    Print #2, "Curso:" & .rdoColumns("Clave") & "|" & .rdoColumns("Nombre_Curso")
                    Print #1, "Fecha Inicial:  " & .rdoColumns("Fecha_Inicio") & "         Hora Inicial:  " & Format(.rdoColumns("Hora_Inicial"), "00.00")
                    Print #2, "Fecha Inicial:" & .rdoColumns("Fecha_Inicio") & "|Hora Inicial:" & Format(.rdoColumns("Hora_Inicial"), "00.00")
                    Print #1, "Fecha Final  :  " & .rdoColumns("Fecha_Inicio") & "         Hora Final:    " & Format(.rdoColumns("Hora_Fin"), "00.00")
                    Print #2, "Fecha Final:" & .rdoColumns("Fecha_Inicio") & "|Hora Final:" & Format(.rdoColumns("Hora_Fin"), "00.00")
                    Print #1,
                    Print #2,
                    Print #1, "Empresa           Código        Nombre                            Puesto                        Asistió        Estatus"
                    Print #2, "Empresa|Código|Nombre|Puesto|Asistió|Estatus"
                    Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                    Print #2, "--------------------------------------------------------------------------------------------------------------------------"
                    Guardar_No_Programa = .rdoColumns("No_Programa_Curso")
                End If

                Print #1, Mid(.rdoColumns("Acronimo"), 1, 16); Spc(18 - Len(Mid(.rdoColumns("Acronimo"), 1, 16))); _
                    Mid(.rdoColumns("No_Tarjeta"), 1, 12); Spc(14 - Len(Mid(.rdoColumns("No_Tarjeta"), 1, 12))); _
                    Mid(.rdoColumns("Empleado"), 1, 30); Spc(32 - Len(Mid(.rdoColumns("Empleado"), 1, 30))); _
                    Mid(.rdoColumns("Nombre"), 1, 32); Spc(34 - Len(Mid(.rdoColumns("Nombre"), 1, 32))); _
                    Mid(.rdoColumns("Asistio"), 1, 2); Spc(14); _
                    Mid(.rdoColumns("Estatus"), 1, 2)
                Print #2, .rdoColumns("Acronimo"); "|"; _
                    .rdoColumns("No_Tarjeta"); "|"; _
                    .rdoColumns("Empleado"); "|"; _
                    .rdoColumns("Nombre"); "|"; _
                    .rdoColumns("Asistio"); "|"; _
                    .rdoColumns("Estatus")
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Indice_Asistencias", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Consulta_Reporte_Indice_Asistencia.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Generar_Reporte_Cursos_Resumen_Mensual()
Dim Consulta_Cursos_Resumen_Mensual As rdoResultset
    
    Mi_SQL = "SELECT Ope_Programacion_Cursos.No_Programa_Curso,Cat_Cursos_Capacitaciones.Nombre AS Curso_Programado_en_el_Mes "
    Mi_SQL = Mi_SQL & ",(SELECT COUNT(Ope_Lista_Asistencia.No_Programa_Curso) From Ope_Lista_Asistencia Where Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso) AS Realizado "
    Mi_SQL = Mi_SQL & ",Cat_Tipos_Cursos.Nombre AS Tipo_Curso ,Cat_Cursos_Capacitaciones.Auditable, Cat_Instituciones.Nombre "
    Mi_SQL = Mi_SQL & ",( SELECT COUNT(DISTINCT (ce.Empleado_ID)) FROM Cat_Empleados ce ,Ope_Lista_Asistencia olas "
    Mi_SQL = Mi_SQL & " Where ce.Empleado_ID = olas.Empleado_ID AND olas.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    Mi_SQL = Mi_SQL & " AND ce.Tipo_Empleado = 'sindicalizado' "
    If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value Then
        Mi_SQL = Mi_SQL & " AND olas.Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & ") AS Sind "
    Mi_SQL = Mi_SQL & ",( SELECT COUNT(DISTINCT (ce.Empleado_ID)) FROM Cat_Empleados ce, Ope_Lista_Asistencia olas "
    Mi_SQL = Mi_SQL & " Where ce.Empleado_ID = olas.Empleado_ID AND olas.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    Mi_SQL = Mi_SQL & " AND ce.Tipo_Empleado != 'sindicalizado' "
     If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value Then
        Mi_SQL = Mi_SQL & " AND olas.Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & ") AS No_Sind "
    Mi_SQL = Mi_SQL & ",(SELECT COUNT(DISTINCT (Ope_Lista_Asistencia.Empleado_ID)) From Cat_Empleados ,Ope_Lista_Asistencia "
    Mi_SQL = Mi_SQL & " Where Cat_Empleados.Empleado_ID = Ope_Lista_Asistencia.Empleado_ID AND Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
     If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value Then
        Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & ") AS Total_Empleados "
    Mi_SQL = Mi_SQL & ",CAST(((DATEDIFF(MI, Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin) * COUNT(*) * 1.00) / 60.0) AS DECIMAL(5, 2)) AS Horas "
    Mi_SQL = Mi_SQL & ",( CAST(((DATEDIFF(MI, Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin) * COUNT(*) * 1.00) / 60.0) AS DECIMAL(5, 2)) * ( "
    Mi_SQL = Mi_SQL & " SELECT COUNT(DISTINCT (Ope_Lista_Asistencia.Empleado_ID)) From Cat_Empleados ,Ope_Lista_Asistencia "
    Mi_SQL = Mi_SQL & " Where Cat_Empleados.Empleado_ID = Ope_Lista_Asistencia.Empleado_ID AND Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value Then
        Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " )) AS HH_Capacitacion "
    Mi_SQL = Mi_SQL & ",(SELECT COUNT(Ope_Programacion_Cursos.Curso_Id) From Ope_Programacion_Cursos ,Cat_Cursos_Capacitaciones "
    Mi_SQL = Mi_SQL & " WHERE Cat_Cursos_Capacitaciones.Auditable = 'SI' AND Ope_Programacion_Cursos.Curso_Id = Cat_Cursos_Capacitaciones.Curso_ID ) AS Total_Si "
    Mi_SQL = Mi_SQL & ",( SELECT COUNT(Ope_Programacion_Cursos.Curso_Id) From Ope_Programacion_Cursos ,Cat_Cursos_Capacitaciones WHERE Cat_Cursos_Capacitaciones.Auditable = 'NO' AND Ope_Programacion_Cursos.Curso_Id = Cat_Cursos_Capacitaciones.Curso_ID ) AS Total_No "
    Mi_SQL = Mi_SQL & " From Ope_Programacion_Cursos ,Cat_Cursos_Capacitaciones ,Cat_Tipos_Cursos ,Cat_Instructores ,Cat_Instituciones "
    Mi_SQL = Mi_SQL & " Where Cat_Cursos_Capacitaciones.Curso_Id = Ope_Programacion_Cursos.Curso_Id "
    Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso "
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Tipo_Curso_Id = Cat_Tipos_Cursos.Tipo_Curso_Id "
    Mi_SQL = Mi_SQL & " AND Cat_Instructores.Instructor_Id = Ope_Programacion_Cursos.Instructor_Id "
    Mi_SQL = Mi_SQL & " AND Cat_Instituciones.Institucion_Id = Cat_Instructores.Institucion_Id "
    If Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Tipo_Curso_Id = '" & Format(Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ItemData(Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Cursos_Resumen_Mensual_Auditable.ListIndex = 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Auditable='SI'"
    ElseIf Cmb_Rpt_Cursos_Resumen_Mensual_Auditable.ListIndex = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Auditable='NO'"
    End If
    Mi_SQL = Mi_SQL & " GROUP BY Ope_Programacion_Cursos.No_Programa_Curso ,Cat_Cursos_Capacitaciones.Nombre "
    Mi_SQL = Mi_SQL & " ,Cat_Tipos_Cursos.Nombre ,Cat_Cursos_Capacitaciones.Auditable ,Cat_Instituciones.Nombre "
    Mi_SQL = Mi_SQL & ",Cat_Cursos_Capacitaciones.Horas ,Ope_Programacion_Cursos.Hora_Inicio "
    Mi_SQL = Mi_SQL & ",Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin "


    Set Consulta_Cursos_Resumen_Mensual = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Consulta_Cursos_Resumen_Mensual.EOF Then
        With Consulta_Cursos_Resumen_Mensual
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value = 1 Then
                Call Encabezado_Reporte("CURSOS RESUMEN MENSUAL", Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, False)
            Else
                Call Encabezado_Reporte("CURSOS RESUMEN MENSUAL", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1, "  Cursos programados      Realizados     Tipo de Curso      S.T.P.S      No.  Participantes     Duración    Hrs/Hombre   "
            Print #1, "      en el mes            Si   No                          SI  NO       Sind No_Sind Total     (Horas)     Capacitación "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #2, "Cursos Programados|Realizados||Tipo de curso|S.T.P.S||No. Participantes|||Duración|Hrs/Hombre"
            Print #2, "en el mes|Si|No||SI|NO|Sind|No Sind|Total|(Horas)|Capacitación|"
            Print #2,
             While Not .EOF
             Dim Realizado_Si As String
             Dim Realizado_No As String
             Dim STPS_Si As String
             Dim STPS_No As String
             If Val(.rdoColumns("Realizado")) > 0 Then
             Realizado_Si = "*"
             Realizado_No = ""
             Else
             Realizado_Si = ""
             Realizado_No = "*"
             End If
             If Trim(.rdoColumns("Auditable")) = "SI" Then
             STPS_Si = "*"
             STPS_No = ""
             Else
             STPS_Si = ""
             STPS_No = "*"
             End If
                Print #1, Mid(.rdoColumns("Curso_Programado_en_el_Mes"), 1, 26); Spc(28 - Len(Mid(.rdoColumns("Curso_Programado_en_el_Mes"), 1, 26))); _
                    Mid(Realizado_Si, 1, 5); Spc(5); _
                    Mid(Realizado_No, 1, 5); Spc(5); _
                    Mid(.rdoColumns("Tipo_Curso"), 1, 20); Spc(22 - Len(Mid(.rdoColumns("Tipo_Curso"), 1, 20))); _
                    Mid(STPS_Si, 1, 5); Spc(3); _
                    Mid(STPS_No, 1, 5); Spc(9); _
                    Mid(.rdoColumns("Sind"), 1, 4); Spc(6 - Len(Mid(.rdoColumns("Sind"), 1, 4))); _
                    Mid(.rdoColumns("No_Sind"), 1, 4); Spc(6 - Len(Mid(.rdoColumns("No_Sind"), 1, 4))); _
                    Mid(.rdoColumns("Total_Empleados"), 1, 4); Spc(12 - Len(Mid(.rdoColumns("Total_Empleados"), 1, 4))); _
                    Mid(Format(.rdoColumns("Horas"), "00.00"), 1, 6); Spc(6); _
                    Mid(Format(.rdoColumns("HH_Capacitacion"), "00.00"), 1, 6); Spc(5)
                Print #2, .rdoColumns("Curso_Programado_en_el_Mes"); "|"; _
                    Realizado_Si; "|"; _
                    Realizado_No; "|"; _
                    .rdoColumns("Tipo_Curso"); "|"; _
                    STPS_Si; "|"; _
                    STPS_No; "|"; _
                    .rdoColumns("Sind"); "|"; _
                    .rdoColumns("No_Sind"); "|"; _
                    .rdoColumns("Total_Empleados"); "|"; _
                    Format(.rdoColumns("Horas"), "00.00"); "|"; _
                    Format(.rdoColumns("HH_Capacitacion"), "00.00"); "|"
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Resumen_Mensual", Me)
         End With
      Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Consulta_Cursos_Resumen_Mensual.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Generar_Reporte_General_Cursos()
Dim Consulta_Reporte_General_Cursos As rdoResultset

    Mi_SQL = "SELECT No_Lista_Asistencia,Cat_Cursos_Capacitaciones.Nombre,Cast(No_Tarjeta as Varchar) as No_Tarjeta"
    Mi_SQL = Mi_SQL & " ,Cat_Empleados.Nombre + ' ' + Cat_Empleados.Apellido_Paterno + ' ' + Cat_Empleados.Apellido_Materno AS Empleado"
    Mi_SQL = Mi_SQL & " ,CAST(((DATEDIFF(MI, Ope_Programacion_Cursos.Hora_Inicio, Ope_Programacion_Cursos.Hora_Fin) * COUNT(*) * 1.00) / 60.0) AS DECIMAL(5, 2)) AS Horas"
    Mi_SQL = Mi_SQL & " ,convert(Date, Ope_Programacion_Cursos.Fecha_Inicio) AS Inicio,convert(Date, Ope_Programacion_Cursos.Fecha_Fin) AS Termino"
    Mi_SQL = Mi_SQL & " ,Cat_Instructores.Nombre + ' ' + Cat_Instructores.Apellido_Paterno + ' ' + Cat_Instructores.Apellido_Materno AS Instructor"
    Mi_SQL = Mi_SQL & " ,Auditable FROM Ope_Lista_Asistencia,Cat_Empleados,Cat_Cursos_Capacitaciones"
    Mi_SQL = Mi_SQL & " ,Ope_Programacion_Cursos,Cat_Instructores Where Cat_Empleados.Empleado_ID = Ope_Lista_Asistencia.Empleado_Id"
    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Capacitaciones.Curso_ID = Ope_Programacion_Cursos.Curso_Id"
    Mi_SQL = Mi_SQL & " AND Ope_Lista_Asistencia.No_Programa_Curso = Ope_Programacion_Cursos.No_Programa_Curso"
    Mi_SQL = Mi_SQL & " AND Cat_Instructores.Instructor_Id = Ope_Programacion_Cursos.Instructor_Id"
    
    If Cmb_Rpt_General_Cursos_Instructor.ListIndex = 1 Then
        Mi_SQL = Mi_SQL & " and Auditable = 'SI'"
    ElseIf Cmb_Rpt_General_Cursos_Instructor.ListIndex = 2 Then
        Mi_SQL = Mi_SQL & " and Auditable = 'NO'"
    End If
    
    If Chk_Rpt_General_Cursos_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " and (Fecha_Hora_Registro BETWEEN '" & Format(Dtp_Rpt_Genera_Cursosl_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " and '" & Format(Dtp_Rpt_General_Cursos_Fecha_Fin.Value, "MM/dd/yyyy") & "')"
    End If
    Mi_SQL = Mi_SQL & " GROUP BY  No_Lista_Asistencia,Cat_Cursos_Capacitaciones.Nombre,No_Tarjeta,Cat_Empleados.Nombre"
    Mi_SQL = Mi_SQL & " ,Cat_Empleados.Apellido_Paterno,Cat_Empleados.Apellido_Materno,Ope_Programacion_Cursos.Hora_Inicio"
    Mi_SQL = Mi_SQL & " ,Ope_Programacion_Cursos.Hora_Fin,Ope_Programacion_Cursos.Fecha_Inicio,Ope_Programacion_Cursos.Fecha_Fin"
    Mi_SQL = Mi_SQL & " ,Cat_Instructores.Nombre,Cat_Instructores.Apellido_Paterno,Cat_Instructores.Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,Auditable order by Cat_Cursos_Capacitaciones.Nombre"
    
    Set Consulta_Reporte_General_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Consulta_Reporte_General_Cursos.EOF Then
        With Consulta_Reporte_General_Cursos
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            If Chk_Rpt_General_Cursos_Fechas.Value = 1 Then
                Call Encabezado_Reporte("REPORTE GENERAL DE CURSOS", Dtp_Rpt_Genera_Cursosl_Fecha_Inicio.Value, Dtp_Rpt_General_Cursos_Fecha_Fin.Value, False)
            Else
                Call Encabezado_Reporte("REPORTE GENERAL DE CURSOS", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            End If
            Print #1, "Registro Patronal: B4744175109"
            Print #1, "RFC: SMG100824LY0"
            Print #1,
            Print #1, "        Nombre           Codigo         Empleado             Horas    Inicio      Termino      Instructor        Auditable"
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "Nombre|Codigo|Empleado|Horas|Inicio|Termino|Instructor|Auditable"
            Print #2,
            While Not .EOF
                Print #1, Mid(.rdoColumns("Nombre"), 1, 25); Spc(27 - Len(Mid(.rdoColumns("Nombre"), 1, 25))); _
                    Mid(.rdoColumns("No_Tarjeta"), 1, 6); Spc(8 - Len(Mid(.rdoColumns("No_Tarjeta"), 1, 6))); _
                    Mid(.rdoColumns("Empleado"), 1, 24); Spc(26 - Len(Mid(.rdoColumns("Empleado"), 1, 24))); _
                    Mid(Format(.rdoColumns("Horas"), "00.00"), 1, 5); Spc(2); _
                    Mid(Format(.rdoColumns("Inicio"), "dd/MM/yyyy"), 1, 12); Spc(2); _
                    Mid(Format(.rdoColumns("Termino"), "dd/MM/yyyy"), 1, 12); Spc(2); _
                    Mid(.rdoColumns("Instructor"), 1, 22); Spc(25 - Len(Mid(.rdoColumns("Instructor"), 1, 22))); _
                    Mid(.rdoColumns("Auditable"), 1, 7)
                Print #2, .rdoColumns("Nombre"); "|"; _
                    .rdoColumns("No_Tarjeta"); "|"; _
                    .rdoColumns("Empleado"); "|"; _
                    Format(.rdoColumns("Horas"), "00.00"); "|"; _
                    Format(.rdoColumns("Inicio"), "dd/MM/yyyy"); "|"; _
                    Format(.rdoColumns("Termino"), "dd/MM/yyyy"); "|"; _
                    .rdoColumns("Instructor"); "|"; _
                    .rdoColumns("Auditable")
                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_General_Cursos", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Consulta_Reporte_General_Cursos.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Public Sub Crea_PDF(Reporte_Rpt As String, Nombre As String)
Dim crxApplication As New CRAXDDRT.Application
Dim crxReport As CRAXDDRT.Report
Dim crxDatabase As CRAXDDRT.Database
Dim crxDatabaseTables As CRAXDDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
Dim crxSections As CRAXDDRT.Sections
Dim crxSection As CRAXDDRT.Section
Dim crxSubreport As CRAXDDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Ruta_RPT As String
Dim Ruta_Salida
    
On Error GoTo HANDLER
    'Asigna el formato de la factura a la variable
    Ruta_RPT = App.Path & "\Reportes\" & Reporte_Rpt & ".rpt"
    Ruta_Salida = App.Path & "\Reportes_Cursos_Capacitaciones\" & Nombre & ".pdf"

     Set crxReport = crxApplication.OpenReport(Ruta_RPT)
           
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo "SRG_Recursos_Humanos", Conectar_Ayudante.Base, Conectar_Ayudante.Usuario_Conexion, Conectar_Ayudante.Password
            End Select
        Next
    End With
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
    Dim Fecha As Date
    Dim parametro As String
        Select Case crParamDef.ParameterFieldName
        'Cursos_Tomados_Por_Empleado
            Case "Empleado_Id_Cursos_Por_Empleado"
                If Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex = 0 Then
                    parametro = "aaaaa"
                Else
                    parametro = Format(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ItemData(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex), "00000")
                End If
                 crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Fecha_Inicio_Cursos_Por_Empleado"
                If Chk_Rpt_Cursos_Tomados_Por_Empleado_Fechas.Value = 1 Then
                   Fecha = Format(Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Inicio.Value, "MM/dd/yyyy") & " 00:00:00"
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
                End If
                crParamDef.AddCurrentValue (Fecha)
            
            Case "Fecha_Fin_Cursos_Por_Empleado"
                If Chk_Rpt_Cursos_Tomados_Por_Empleado_Fechas.Value = 1 Then
                   Fecha = Format(Dtp_Rpt_Cursos_Tomados_Por_Empleado_Fecha_Fin.Value, "MM/dd/yyyy") & " 23:59:59"
                Else
                   Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
                End If
                crParamDef.AddCurrentValue (Fecha)
                
            Case "Curso_Id_Cursos_Programas_Hombre_Curso"
                parametro = Format(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex), "00000")
                 crParamDef.AddCurrentValue ("'" & parametro & "'")
                 
            Case "Tipo_Empleado"
                 parametro = "SINDICALIZADO-CONFIANZA"
                 If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex = 1 Then
                 parametro = "sindicalizado"
                 End If
                 If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex = 2 Then
                 parametro = "confianza"
                 End If
                 
                 crParamDef.AddCurrentValue ("'" & parametro & "'")
                   
            Case "Fecha_Inicio_Programas_Hombre_Curso"
                If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
                    Fecha = Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Inicio.Value, " MM/dd/yyyy") & " 00:00:00"
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
                End If
                    crParamDef.AddCurrentValue (Fecha)
                        
            Case "Fecha_Fin_Programas_Hombre_Curso"
                If Chk_Rpt_Cursos_Hora_Hombre_Fechas.Value = 1 Then
                    Fecha = Format(Dtp_Rpt_Cursos_Hora_Hombre_Fecha_Termino.Value, "MM/dd/yyyy") & " 23:59:59"
                Else
                    Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
                End If
                    crParamDef.AddCurrentValue (Fecha)

            Case "Empleado_Id_Cursos_Programas_Hombre_Empleado"
                parametro = Format(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex), "00000")
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Departamento_Id_Cursos_Programas_Hombre_Departamento"
                parametro = Format(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex), "00000")
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Auditable_Cursos_Resumen_Mensual"
                parametro = "aa"
                    If Cmb_Rpt_Cursos_Resumen_Mensual_Auditable.ListIndex = 0 Then
                        parametro = "SI"
                    ElseIf Cmb_Rpt_Cursos_Resumen_Mensual_Auditable.ListIndex = 1 Then
                        parametro = "NO"
                    End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Tipo_Curso_Id_Cursos_Resumen_Mensual"
                parametro = "aaaaa"
                    If Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListIndex > 0 Then
                        parametro = Format(Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ItemData(Cmb_Rpt_Cursos_Resumen_Mesual_Tipo_Curso.ListIndex), "00000")
                    End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Fecha_Inicio_Cursos_Resumen_Mensual"
                If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value = 1 Then
                   Fecha = Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Inicio.Value, "MM/dd/yyyy") & " 00:00:00"
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
                End If
                crParamDef.AddCurrentValue (Fecha)
            
            Case "Fecha_Fin_Cursos_Resumen_Mensual"
                If Chk_Rpt_Cursos_Resumen_Mensual_Fechas.Value = 1 Then
                    Fecha = Format(Dtp_Rpt_Cursos_Resumen_Mensual_Fecha_Fin.Value, "MM/dd/yyyy") & " 23:59:59"
                Else
                   Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
                End If
                crParamDef.AddCurrentValue (Fecha)
                
            Case "Tipo_Curso_Cursos_Reporte_General"
                If Cmb_Rpt_General_Cursos_Instructor.ListIndex > 0 Then
                    If Cmb_Rpt_General_Cursos_Instructor.ListIndex = 1 Then
                        parametro = "SI"
                    Else
                        parametro = "NO"
                    End If
                Else
                    parametro = "am"
                End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")

'            Case "Instructor_Id_Cursos_Reporte_General"
'                If Cmb_Rpt_General_Cursos_Instructor.ListIndex > 0 Then
'                    parametro = Format(Cmb_Rpt_General_Cursos_Instructor.ItemData(Cmb_Rpt_General_Cursos_Instructor.ListIndex), "00000")
'                Else
'                    parametro = "aaaaa"
'                End If
'                crParamDef.AddCurrentValue ("'" & parametro & "'")
'
'            Case "Institucion_Id_Cursos_Reporte_General"
'                If Cmb_Rpt_General_Cursos_Institucion.ListIndex > 0 Then
'                    parametro = Format(Cmb_Rpt_General_Cursos_Institucion.ItemData(Cmb_Rpt_General_Cursos_Institucion.ListIndex), "00000")
'                Else
'                    parametro = "aaaaa"
'                End If
'                crParamDef.AddCurrentValue ("'" & parametro & "'")
'
'            Case "Sala_Id_Cursos_Reporte_General"
'                If Cmb_Rpt_General_Cursos_Sala.ListIndex > 0 Then
'                    parametro = Format(Cmb_Rpt_General_Cursos_Sala.ItemData(Cmb_Rpt_General_Cursos_Sala.ListIndex), "00000")
'                Else
'                    parametro = "aaaaa"
'                End If
'                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Fecha_Inicio_Cursos_Reporte_General"
                If Chk_Rpt_General_Cursos_Fechas.Value = 1 Then
                    Fecha = Format(Dtp_Rpt_Genera_Cursosl_Fecha_Inicio.Value, "MM/dd/yyyy") & " 00:00:00"
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
                End If
                    crParamDef.AddCurrentValue (Fecha)
            
            Case "Fecha_Fin_Cursos_Reporte_General"
                If Chk_Rpt_General_Cursos_Fechas.Value = 1 Then
                   Fecha = Format(Dtp_Rpt_General_Cursos_Fecha_Fin.Value, "MM/dd/yyyy") & " 23:59:59"
                Else
                   Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
                End If
                crParamDef.AddCurrentValue (Fecha)
            'AQUIIIII
            Case "Empleado_Id_Cursos_Indices_Asistencia"
            If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 1 Then
                    parametro = Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000")
     
               Else
                    parametro = "aaaaa"
                End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Curso_Id_Cursos_Indices_Asistencia"
                If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 2 Then
                    parametro = Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000")
                Else
                    parametro = "aaaaa"
                End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            Case "Departamento_Id_Cursos_Indices_Asistencia"
                If Cmb_Rpt_Cursos_Indice_Asistencias_Tipo_Busqueda.ListIndex = 3 Then
                    parametro = Format(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex), "00000")
                Else
                    parametro = "aaaaa"
                End If
                crParamDef.AddCurrentValue ("'" & parametro & "'")
            Case "Fecha_Inicio_Cursos_Indices_Asistencia"
                If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
                    Fecha = Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value, "MM/dd/yyyy") & " 00:00:00"
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
                End If
                    crParamDef.AddCurrentValue (Fecha)
            
            Case "Fecha_Fin_Cursos_Indices_Asistencia"
                If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
                   Fecha = Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value, "MM/dd/yyyy") & " 23:59:59"
                Else
                   Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
                End If
                crParamDef.AddCurrentValue (Fecha)
            
        End Select
    Next
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
   crxReport.ExportOptions.DiskFileName = Ruta_Salida

   

    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat
'crxReport.ExportOptions.FormatType = crEFTExcel97
    crxReport.ExportOptions.PDFExportAllPages = True
    'Oculta el progreso de la exportacion
    crxReport.DisplayProgressDialog = False
    'Genera la exportación del documento
    crxReport.Export (False)
    'Destruye el documento
    Set crxReport = Nothing
    ShellExecute Me.hwnd, "open", Ruta_Salida, "", "", 4
    Exit Sub
HANDLER:
    Printer.EndDoc
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


''*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo_Consulta
    'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
    'PARÁMETROS:
    '             1. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             2. Consulta: SQL para llenar el combo
    'CREO: Ana Laura Huichapa Ramírez
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Llena_Combo_Consulta(Combo_Control As ComboBox, Consulta As String)
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = Consulta
'        .SQL = .SQL & " FROM " & Tabla
'        If Tipo = 1 Then
'            .SQL = .SQL & " WHERE " & Campo_con & " LIKE '%" & Combo_Control.Text & "%' " & Condicion_Adicional
'        End If
'        .SQL = .SQL & " ORDER BY " & Campo_con
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    Combo_Control.Clear
            Combo_Control.AddItem "TODOS"
            
        Combo_Control.ItemData(Combo_Control.NewIndex) = 0
  
    If Not Rs_Combo.EOF Then
        While Not Rs_Combo.EOF
            Combo_Control.AddItem Rs_Combo(2)
            Combo_Control.ItemData(Combo_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Sub

Private Sub Txt_No_Tarjeta_Cursos_Hioras_Hombre_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Departamento As rdoResultset
Dim No_Tarjeta As String
If Txt_No_Tarjeta_Cursos_Hioras_Hombre.Visible Then
 If KeyAscii = 13 Then
        No_Tarjeta = Format(Txt_No_Tarjeta_Cursos_Hioras_Hombre.Text, "00000")
           Mi_SQL = "select Empleado_ID, ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') + ' ' "
           Mi_SQL = Mi_SQL & "+ ISNULL(Nombre, '') as Nombre From Cat_Empleados "
           Mi_SQL = Mi_SQL & "WHERE Estatus = 'A' "
           If Trim(No_Tarjeta) = "" Then
           Mi_SQL = Mi_SQL & "and No_Tarjeta like '%%' "
           Else
            Mi_SQL = Mi_SQL & " and No_Tarjeta = " & No_Tarjeta & " "
           End If
           
           Mi_SQL = Mi_SQL & "ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.Clear
            If Trim(No_Tarjeta) = "" Then
'            Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.AddItem "TODOS"
'
'            Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.NewIndex) = 0
        End If
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ItemData(Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListCount > 0 Then
                Cmb_Rpt_Cursos_Hora_Hombre_Busqueda.ListIndex = 0
            End If
        
           
End If
End If
End Sub

Private Sub Txt_No_Tarjeta_Cursos_Por_Empleado_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Departamento As rdoResultset
Dim No_Tarjeta As String
 If KeyAscii = 13 Then
        No_Tarjeta = Format(Txt_No_Tarjeta_Cursos_Por_Empleado.Text, "00000")
           Mi_SQL = "select Empleado_ID, ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') + ' ' "
           Mi_SQL = Mi_SQL & "+ ISNULL(Nombre, '') as Nombre From Cat_Empleados "
           Mi_SQL = Mi_SQL & "WHERE Estatus = 'A' "
           If Trim(No_Tarjeta) = "" Then
           Mi_SQL = Mi_SQL & "and No_Tarjeta like '%%' "
           Else
            Mi_SQL = Mi_SQL & " and No_Tarjeta = " & No_Tarjeta & " "
           End If
           
           Mi_SQL = Mi_SQL & "ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.Clear
'            If Trim(No_Tarjeta) = "" Then
            Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.AddItem "TODOS"

            Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ItemData(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.NewIndex) = 0
'        End If
            While Not Rs_Empleados_Departamento.EOF
                If Trim(Rs_Empleados_Departamento.rdoColumns("Nombre")) <> "" Then
                Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ItemData(Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
                End If
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListCount > 0 Then
            
            If No_Tarjeta = "" Then
             Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex = 0
            Else
            If Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListCount = 2 Then
             Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex = 1
             Else
             Cmb_Rpt_Cursos_Tomados_Por_Empleado_Empleado.ListIndex = 0
             End If
            End If
               
            End If
        
           
End If
End Sub

Private Sub Txt_No_Tarjeta_Indices_Asistencia_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Departamento As rdoResultset
Dim No_Tarjeta As String
If Txt_No_Tarjeta_Indices_Asistencia.Visible Then
 If KeyAscii = 13 Then
        No_Tarjeta = Format(Txt_No_Tarjeta_Indices_Asistencia.Text, "00000")
           Mi_SQL = "select Empleado_ID, ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') + ' ' "
           Mi_SQL = Mi_SQL & "+ ISNULL(Nombre, '') as Nombre From Cat_Empleados "
           Mi_SQL = Mi_SQL & "WHERE Estatus = 'A' "
           If Trim(No_Tarjeta) = "" Then
           Mi_SQL = Mi_SQL & "and No_Tarjeta like '%%' "
           Else
            Mi_SQL = Mi_SQL & " and No_Tarjeta = " & No_Tarjeta & " "
           End If
           
           Mi_SQL = Mi_SQL & "ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.Clear
            If Trim(No_Tarjeta) = "" Then
'            Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.AddItem "TODOS"
'
'            Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.NewIndex) = 0
        End If
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ItemData(Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListCount > 0 Then
                Cmb_Rpt_Cursos_Indice_Asistencias_Busqueda.ListIndex = 0
            End If
        
           

End If
End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Accesos_Almacenes
'DESCRIPCION: Genera el reporte de los accesos al almacen
'PARAMETROS :
'CREO       : Flores Ramirez Yazmin
'FECHA_CREO : 13-Diciembre-2016
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Accesos_Almacenes()
Dim Rs_Consulta_Checadores_Almacen As rdoResultset
    
    'Consulta los registros de checadas
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores_Almacenes AARC"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"

    'Validacion de Empleado
    If Cmb_Rpt_Empleado_Accesos_Almacenes.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID='" & Format(Cmb_Rpt_Empleado_Accesos_Almacenes.ItemData(Cmb_Rpt_Empleado_Accesos_Almacenes.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    If Chk_Rpt_Accesos_Almacen_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Rpt_Accesos_Almacenes_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Accesos_Almacenes_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AARC.Hora"
    Set Rs_Consulta_Checadores_Almacen = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Checadores_Almacen.EOF Then
        With Rs_Consulta_Checadores_Almacen
            MDIFrm_Apl_Principal.MousePointer = 11
            'Agrega el encabezado al reporte
            Call Encabezado_Reporte("REPORTE DE IMPORTACION ACCESOS ALMACENES", Format(Now, "dd MMMM yyyy HH:mm:ss"), , True)
            Print #1,
'            Print #1, "Curso: "; .rdoColumns("Curso")
'            Print #1, "Tipo : "; .rdoColumns("Tipo"); "     Horas : "; .rdoColumns("Horas")
            Print #1,
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "No. Nomina        Empleado                  Fecha           Hora "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #2, "No. Nomina|Empleado|||Fecha|Hora"
            While Not .EOF
                        
                Print #1, .rdoColumns("No_Tarjeta"); _
                    Spc(16); Mid(.rdoColumns("Nombre"), 1, 25); _
                    Spc(15); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(5); Format(.rdoColumns("Hora"), "HH:mm:ss")

                Print #2, .rdoColumns("No_Tarjeta"); "|"; .rdoColumns("Nombre"); "|||"; _
                        Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); "|"; _
                        Format(.rdoColumns("Hora"), "HH:mm:ss")
                        

                .MoveNext
            Wend
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_Rpt_Accesos_Almacenes", Me)
        End With
    Else
        MsgBox "No hay registros que mostrar", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Checadores_Almacen.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

