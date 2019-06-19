VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Frm_Adm_Bolsa_Horas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Bolsa_Horas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9045
      Left            =   0
      ScaleHeight     =   9045
      ScaleWidth      =   7740
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      Begin VB.Frame Fra_Bolsa_Horas_General 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos Generales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   45
         TabIndex        =   1
         Top             =   360
         Width           =   7500
         Begin VB.CommandButton Btn_Bolsa_Horas_Agregar 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   6060
            TabIndex        =   31
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox Txt_Hrs_Pagado 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   5640
            MaxLength       =   5
            TabIndex        =   26
            Top             =   2280
            Width           =   1770
         End
         Begin VB.TextBox Txt_Departmento 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1560
            Width           =   5850
         End
         Begin VB.TextBox Txt_Supervisor 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1200
            Width           =   5850
         End
         Begin VB.ComboBox Cmb_Bolsa_Horas_Empleado 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Frm_Adm_Bolsa_Horas.frx":0000
            Left            =   1545
            List            =   "Frm_Adm_Bolsa_Horas.frx":0002
            TabIndex        =   21
            Top             =   840
            Width           =   5850
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Empleado_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5645
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_No_Asistencias 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Hrs_Debe 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   5640
            MaxLength       =   5
            TabIndex        =   6
            Top             =   1920
            Width           =   1770
         End
         Begin VB.TextBox Txt_Bolsa_Horas_ID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   405
            Width           =   1770
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Referencia 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   3465
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Referencia_Siguiente 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   3780
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.TextBox Txt_Adm_Asistencias_Movimiento 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   4095
            Visible         =   0   'False
            Width           =   1750
         End
         Begin MSComCtl2.DTPicker Dtp_Bolsa_Horas_Fecha 
            Height          =   315
            Left            =   5640
            TabIndex        =   9
            Top             =   420
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   135659523
            CurrentDate     =   39986
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Debe 
            Height          =   315
            Left            =   1560
            TabIndex        =   24
            Top             =   1920
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy"
            Format          =   135659523
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Pagado 
            Height          =   315
            Left            =   1545
            TabIndex        =   27
            Top             =   2280
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy"
            Format          =   135659523
            CurrentDate     =   39940
         End
         Begin VB.ComboBox Cmb_Supervisor 
            Height          =   315
            Left            =   1800
            TabIndex        =   34
            Top             =   1200
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.ComboBox Cmb_Departamento 
            Height          =   315
            Left            =   1560
            TabIndex        =   35
            Top             =   1560
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dia que Debe"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   1995
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado ID"
            Height          =   195
            Left            =   4680
            TabIndex        =   19
            Top             =   3525
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No_Asistencia"
            Height          =   195
            Left            =   45
            TabIndex        =   18
            Top             =   3525
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Lbl_Horas_Trabajadaas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hrs."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            TabIndex        =   17
            Top             =   2355
            Width           =   315
         End
         Begin VB.Label Lbl_Horas_Extra 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hrs."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            TabIndex        =   16
            Top             =   1920
            Width           =   315
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dia Pagado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   2355
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            TabIndex        =   14
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Lbl_Bolsa_Horas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Bolsa Horas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   11
            Top             =   1245
            Width           =   855
         End
         Begin VB.Label Lbl_Turno 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Area"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   10
            Top             =   1635
            Width           =   375
         End
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   645
         Left            =   45
         Picture         =   "Frm_Adm_Bolsa_Horas.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   58
         Tag             =   "A"
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   1290
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   645
         Left            =   2040
         Picture         =   "Frm_Adm_Bolsa_Horas.frx":058E
         Style           =   1  'Graphical
         TabIndex        =   57
         Tag             =   "M"
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   1290
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   645
         Left            =   6180
         Picture         =   "Frm_Adm_Bolsa_Horas.frx":0B18
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   1290
      End
      Begin VB.CommandButton Btn_Consultar 
         Caption         =   "Consultar"
         Height          =   645
         Left            =   4140
         Picture         =   "Frm_Adm_Bolsa_Horas.frx":10A2
         Style           =   1  'Graphical
         TabIndex        =   55
         Tag             =   "C"
         Top             =   8280
         Width           =   1290
      End
      Begin VB.Frame Fra_Control_Bolsa_Horas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   45
         TabIndex        =   30
         Top             =   3480
         Width           =   7500
         Begin VB.CommandButton Btn_Bolsa_Horas_Eliminar 
            Caption         =   "Eliminar"
            Height          =   315
            Left            =   6060
            TabIndex        =   32
            Top             =   3280
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Bolsa_Horas 
            Height          =   2985
            Left            =   75
            TabIndex        =   33
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5265
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Fra_Observaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   45
         TabIndex        =   28
         Top             =   7200
         Width           =   7500
         Begin VB.TextBox Txt_Observaciones 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   90
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   225
            Width           =   7335
         End
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTROL TIEMPO POR TIEMPO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2055
         TabIndex        =   20
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.PictureBox Pic_Bolsa_Horas_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8850
      Left            =   0
      ScaleHeight     =   8850
      ScaleWidth      =   7605
      TabIndex        =   36
      Top             =   360
      Width           =   7605
      Begin VB.Frame Fra_Permisos_Consulta_Resultados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5805
         Left            =   45
         TabIndex        =   53
         Top             =   2760
         Width           =   7520
         Begin MSFlexGridLib.MSFlexGrid Grid_Consulta_Resultados 
            Height          =   5430
            Left            =   90
            TabIndex        =   54
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   9578
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Fra_Permisos_Consulta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filtros Busqueda"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   45
         TabIndex        =   37
         Top             =   75
         Width           =   7520
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   510
            Left            =   6240
            Picture         =   "Frm_Adm_Bolsa_Horas.frx":162C
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "C"
            Top             =   225
            Width           =   1110
         End
         Begin VB.CheckBox Chk_Consulta_Empleado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   46
            Top             =   1125
            Width           =   1050
         End
         Begin VB.CheckBox Chk_Consulta_Periodo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   45
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Regresar 
            Cancel          =   -1  'True
            Caption         =   "Regresar"
            Height          =   510
            Left            =   6240
            Picture         =   "Frm_Adm_Bolsa_Horas.frx":1BB6
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1965
            UseMaskColor    =   -1  'True
            Width           =   1110
         End
         Begin VB.CheckBox Chk_Consulta_Supervisor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   750
            Width           =   1050
         End
         Begin VB.CheckBox Chk_Consulta_Departamentos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   360
            Width           =   1410
         End
         Begin VB.ComboBox Cmb_Consulta_Departamento 
            Height          =   315
            Left            =   1665
            TabIndex        =   41
            Top             =   360
            Width           =   4470
         End
         Begin VB.ComboBox Cmb_Consulta_Supervisor 
            Height          =   315
            Left            =   1665
            TabIndex        =   40
            Top             =   750
            Width           =   4470
         End
         Begin VB.ComboBox Cmb_Consulta_Empleado 
            Height          =   315
            Left            =   1665
            TabIndex        =   39
            Top             =   1125
            Width           =   4470
         End
         Begin VB.CommandButton Btn_Exportar 
            Caption         =   "Exportar"
            Height          =   510
            Left            =   6240
            Picture         =   "Frm_Adm_Bolsa_Horas.frx":2140
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "A"
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker Dtp_Consulta_Fecha_Termino 
            Height          =   315
            Left            =   4170
            TabIndex        =   48
            Top             =   1560
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   135659523
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Consulta_Fecha_Inicio 
            Height          =   315
            Left            =   1665
            TabIndex        =   49
            Top             =   1560
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   135659523
            CurrentDate     =   39940
         End
         Begin MSComctlLib.ProgressBar Prbar_Exportacion 
            Height          =   165
            Left            =   6300
            TabIndex        =   50
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   3765
            TabIndex        =   52
            Top             =   1620
            Width           =   135
         End
         Begin VB.Label Lbl_Progreso_Exportacion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exportando..."
            Height          =   195
            Left            =   6300
            TabIndex        =   51
            Top             =   1440
            Visible         =   0   'False
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "Frm_Adm_Bolsa_Horas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Operacion As String
Private Simbologia As String
Private SubSimbologia As String
Private Simbologia_Consulta As String
Private SubSimbologia_Consulta As String
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios
Dim Archivo_Reporte_Abierto As Boolean  'Indica si el archivo de reporte esta abierto

Private Sub Btn_Bolsa_Horas_Agregar_Click()
    'Valida los datos del dependiente
    If Cmb_Bolsa_Horas_Empleado.Text <> "" And Trim(Txt_Hrs_Debe.Text) <> "" And _
        Txt_Hrs_Pagado.Text <> "" Then
        'Agrega el dependiente a la lista
        Grid_Bolsa_Horas.Cols = 10
        If Grid_Bolsa_Horas.Rows = 0 Then
            Grid_Bolsa_Horas.AddItem "Empleado_ID" & Chr(9) & "Empleado" & Chr(9) & _
                    "Supervisor" & Chr(9) & "" & Chr(9) & "Area" & Chr(9) & "" & Chr(9) & "Dia que Debe" & Chr(9) & "Hrs" & Chr(9) & "Dia Pagado" & Chr(9) & "Hrs"
            Grid_Bolsa_Horas.ColWidth(0) = 0  'Empleado id
            Grid_Bolsa_Horas.ColWidth(1) = 3200  'Empleado
            Grid_Bolsa_Horas.ColWidth(2) = 3200  'Supervisor
            Grid_Bolsa_Horas.ColWidth(3) = 0  'Supervisor id
            Grid_Bolsa_Horas.ColWidth(4) = 3200  'Area
            Grid_Bolsa_Horas.ColWidth(5) = 0  'Area id
            Grid_Bolsa_Horas.ColWidth(6) = 1100  'Dia Debe
            Grid_Bolsa_Horas.ColWidth(7) = 500  'Hrs
            Grid_Bolsa_Horas.ColAlignment(7) = flexAlignCenterCenter
            Grid_Bolsa_Horas.ColWidth(8) = 1100  'Dia Pagado
            Grid_Bolsa_Horas.ColWidth(9) = 500  'Hrs
            Grid_Bolsa_Horas.ColAlignment(9) = flexAlignCenterCenter
        End If
        If Cmb_Supervisor.ListIndex > -1 Then
        Grid_Bolsa_Horas.AddItem Format(Cmb_Bolsa_Horas_Empleado.ItemData(Cmb_Bolsa_Horas_Empleado.ListIndex), "00000") & Chr(9) & Trim(Cmb_Bolsa_Horas_Empleado.Text) & Chr(9) & _
            Trim(Txt_Supervisor.Text) & Chr(9) & Format(Cmb_Supervisor.ItemData(Cmb_Supervisor.ListIndex), "00000") & Chr(9) & Trim(Txt_Departmento.Text) & Chr(9) & _
            Format(Cmb_Departamento.ItemData(Cmb_Departamento.ListIndex), "00000") & Chr(9) & Format(Dtp_Fecha_Debe.Value, "dd/MMM/yyyy") & Chr(9) & Trim(Txt_Hrs_Debe.Text) & Chr(9) & _
            Format(Dtp_Fecha_Pagado.Value, "dd/MMM/yyyy") & Chr(9) & Trim(Txt_Hrs_Pagado.Text)
        Else
        Grid_Bolsa_Horas.AddItem Format(Cmb_Bolsa_Horas_Empleado.ItemData(Cmb_Bolsa_Horas_Empleado.ListIndex), "00000") & Chr(9) & Trim(Cmb_Bolsa_Horas_Empleado.Text) & Chr(9) & _
            Trim(Txt_Supervisor.Text) & Chr(9) & Null & Chr(9) & Trim(Txt_Departmento.Text) & Chr(9) & _
            Format(Cmb_Departamento.ItemData(Cmb_Departamento.ListIndex), "00000") & Chr(9) & Format(Dtp_Fecha_Debe.Value, "dd/MMM/yyyy") & Chr(9) & Trim(Txt_Hrs_Debe.Text) & Chr(9) & _
            Format(Dtp_Fecha_Pagado.Value, "dd/MMM/yyyy") & Chr(9) & Trim(Txt_Hrs_Pagado.Text)
        End If
        
            
'        Cmb_Bolsa_Horas_Empleado.ListIndex = -1
'        Txt_Supervisor.Text = ""
'        Txt_Departmento.Text = ""
        Dtp_Fecha_Debe.Value = Now
        Txt_Hrs_Debe.Text = ""
        Dtp_Fecha_Pagado.Value = Now
        Txt_Hrs_Pagado.Text = ""
        
        Grid_Bolsa_Horas.FixedRows = 1
    Else
    
    End If
End Sub

Private Sub Btn_Bolsa_Horas_Eliminar_Click()
    If Grid_Bolsa_Horas.Rows > 0 Then
        If Grid_Bolsa_Horas.Rows = 2 Then
            Grid_Bolsa_Horas.Rows = 0
        Else
            Grid_Bolsa_Horas.RemoveItem Grid_Bolsa_Horas.RowSel
        End If
    End If
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Control_Bolsa_Horas
'DESCRIPCION: Da de alta un registro en Adm_Control_Bolsa_Horas
'PARAMETROS :
'CREO       : Yazmin Flores Ramirez
'FECHA_CREO : 25-Noviembre-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Control_Bolsa_Horas()
Dim Rs_Modifica_Control_Bolsa_Horas As rdoResultset 'Informacion del registro
Dim Rs_Modifica_Control_Bolsa_Horas_Detalles As rdoResultset

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Adm_Control_Bolsa_Horas"
    Mi_SQL = Mi_SQL & " WHERE No_Control_Bolsa_Horas = '" & Trim(Txt_Bolsa_Horas_ID.Text) & "'"
    Set Rs_Modifica_Control_Bolsa_Horas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Adm_Control_Bolsa_Horas con los datos contenidos en las cajas de textos
    With Rs_Modifica_Control_Bolsa_Horas
        .Edit
            .rdoColumns("Fecha_Solicitud") = Format(Dtp_Bolsa_Horas_Fecha.Value, "MM/dd/yyyy")
            If Trim(Txt_Observaciones.Text) <> "" Then
                 .rdoColumns("Observaciones") = Trim(Txt_Observaciones.Text)
            End If
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Control_Bolsa_Horas = Nothing
    
    'Alta de Control Bolsa de horas Detalles
    Dim Cont_Fila As Integer
    Mi_SQL = "DELETE FROM Adm_Control_Bolsa_Horas_Detalles WHERE No_Control_Bolsa_Horas = '" & Trim(Txt_Bolsa_Horas_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    
    Set Rs_Modifica_Control_Bolsa_Horas_Detalles = Conectar_Ayudante.Recordset_Agregar("Adm_Control_Bolsa_Horas_Detalles")
    With Rs_Modifica_Control_Bolsa_Horas_Detalles
        For Cont_Fila = 1 To Grid_Bolsa_Horas.Rows - 1
            .AddNew
                .rdoColumns("No_Control_Bolsa_Horas") = Trim(Txt_Bolsa_Horas_ID.Text)
                .rdoColumns("Empleado_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 0)
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 3) <> "" Then
                    .rdoColumns("Supervisor_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 3)
                Else
                    .rdoColumns("Supervisor_ID") = Null
                End If
                .rdoColumns("Departamento_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 5)
                .rdoColumns("Fecha_Debe") = Format(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 6), "MM/dd/yyyy")
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 7) <> "" Then
                    .rdoColumns("Horas_Debe") = Val(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 7))
                Else
                    .rdoColumns("Horas_Debe") = Null
                End If
                .rdoColumns("Fecha_Pagado") = Format(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 8), "MM/dd/yyyy")
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 9) <> "" Then
                    .rdoColumns("Horas_Pagado") = Val(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 9))
                Else
                    .rdoColumns("Horas_Pagado") = Null
                End If
            .Update
        Next
    End With
    Set Rs_Modifica_Control_Bolsa_Horas_Detalles = Nothing
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Nuevo.Enabled = True
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Exportar.Enabled = True
    
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    
    Grid_Bolsa_Horas.Rows = 0
    Cmb_Bolsa_Horas_Empleado.Clear
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Adm_Bolsa_Horas", Me)
    MsgBox "Control Bolsa de horas ha sido modificado", vbOKOnly + vbInformation, Me.Caption
    
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Control_Bolsa_Horas
'DESCRIPCION: Da de alta un registro en Adm_Control_Bolsa_Horas
'PARAMETROS :
'CREO       : Yazmin Flores Ramirez
'FECHA_CREO : 25-Noviembre-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Control_Bolsa_Horas()
Dim Rs_Alta_Control_Bolsa_Horas As rdoResultset 'Informacion del registro
Dim Rs_Alta_Control_Bolsa_Horas_Detalles As rdoResultset

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Set Rs_Alta_Control_Bolsa_Horas = Conectar_Ayudante.Recordset_Agregar("Adm_Control_Bolsa_Horas")
    'Agrega el reigstro del Control Bolsa de horas
    With Rs_Alta_Control_Bolsa_Horas
        .AddNew
        Txt_Bolsa_Horas_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Control_Bolsa_Horas", "No_Control_Bolsa_Horas"), "0000000000")

            .rdoColumns("No_Control_Bolsa_Horas") = Trim(Txt_Bolsa_Horas_ID.Text)
            .rdoColumns("Fecha_Solicitud") = Format(Dtp_Bolsa_Horas_Fecha.Value, "MM/dd/yyyy")
            If Trim(Txt_Observaciones.Text) <> "" Then
                 .rdoColumns("Observaciones") = Trim(Txt_Observaciones.Text)
            End If
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Control_Bolsa_Horas = Nothing
    
    'Alta de Control Bolsa de horas Detalles
    Dim Cont_Fila As Integer
    Set Rs_Alta_Control_Bolsa_Horas_Detalles = Conectar_Ayudante.Recordset_Agregar("Adm_Control_Bolsa_Horas_Detalles")
    With Rs_Alta_Control_Bolsa_Horas_Detalles
        For Cont_Fila = 1 To Grid_Bolsa_Horas.Rows - 1
            .AddNew
                .rdoColumns("No_Control_Bolsa_Horas") = Trim(Txt_Bolsa_Horas_ID.Text)
                .rdoColumns("Empleado_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 0)
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 3) <> "" Then
                    .rdoColumns("Supervisor_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 3)
                Else
                    .rdoColumns("Supervisor_ID") = Null
                End If
                .rdoColumns("Departamento_ID") = Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 5)
                .rdoColumns("Fecha_Debe") = Format(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 6), "MM/dd/yyyy")
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 7) <> "" Then
                    .rdoColumns("Horas_Debe") = Val(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 7))
                Else
                    .rdoColumns("Horas_Debe") = Null
                End If
                .rdoColumns("Fecha_Pagado") = Format(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 8), "MM/dd/yyyy")
                If Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 9) <> "" Then
                    .rdoColumns("Horas_Pagado") = Val(Grid_Bolsa_Horas.TextMatrix(Cont_Fila, 9))
                Else
                    .rdoColumns("Horas_Pagado") = Null
                End If
            .Update
        Next
    End With
    Set Rs_Alta_Control_Bolsa_Horas_Detalles = Nothing
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Exportar.Enabled = True
    
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    
    Grid_Bolsa_Horas.Rows = 0
    Cmb_Bolsa_Horas_Empleado.Clear
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Adm_Bolsa_Horas", Me)
    MsgBox "Control Bolsa de horas dado de alta", vbOKOnly + vbInformation, Me.Caption
    
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
        
Private Sub Btn_Buscar_Click()
     Consulta_Control_Bolsa_Horas
End Sub
Private Sub Consulta_Control_Bolsa_Horas()
Dim Mi_SQL As String
Dim Columna As Integer
Dim Tem_Empleado As String
Dim Supervisor_ID As String
Dim Supervisor As String
Dim Rs_Empleados As rdoResultset
Dim Rs_Supervisor_Empleado As rdoResultset
Dim Total_Diferencia As Double

On Error GoTo HANDLER
Grid_Consulta_Resultados.Rows = 0
    If Grid_Consulta_Resultados.Rows = 0 Then
        Grid_Consulta_Resultados.Rows = 0
        Grid_Consulta_Resultados.Cols = 14
        Grid_Consulta_Resultados.AddItem "" _
        & Chr(9) & "Empleado" _
        & Chr(9) & "Empleado_ID" _
        & Chr(9) & "No.Control" _
        & Chr(9) & "Supervisor" _
        & Chr(9) & "Supervisor_ID" _
        & Chr(9) & "Area" _
        & Chr(9) & "Departamento_ID" _
        & Chr(9) & "Dia que Debe" _
        & Chr(9) & "Hrs" _
        & Chr(9) & "Dia Pagado" _
        & Chr(9) & "Hrs." _
        & Chr(9) & "Diferencia" _
        & Chr(9) & "Total"
        
    End If
    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Consulta la unidad del producto intermedio
    Mi_SQL = "SELECT Adm_Control_Bolsa_Horas_Detalles.*, Horas_Debe-Horas_Pagado as Diferencia, Adm_Control_Bolsa_Horas_Detalles.Empleado_ID,Adm_Control_Bolsa_Horas_Detalles.Departamento_ID,(Cat_Empleados.Apellido_paterno+' '+ Cat_Empleados.Apellido_Materno+' '+ Cat_Empleados.Nombre) as Nombre ,"
    Mi_SQL = Mi_SQL & " Cat_Departamentos.Nombre as Departamento "
    Mi_SQL = Mi_SQL & " FROM Adm_Control_Bolsa_Horas_Detalles,Cat_Empleados, Cat_Departamentos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Control_Bolsa_Horas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Control_Bolsa_Horas_Detalles.Departamento_ID=Cat_Departamentos.Departamento_ID"
    'Departamentos
    If Chk_Consulta_Departamentos.Value = 1 Then
        If Cmb_Consulta_Departamento.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Control_Bolsa_Horas_Detalles.Departamento_ID = '" & Format(Cmb_Consulta_Departamento.ItemData(Cmb_Consulta_Departamento.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun departamento", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Supervisor
    If Chk_Consulta_Supervisor.Value = 1 Then
        If Cmb_Consulta_Supervisor.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Empleados.Supervisor_ID = '" & Format(Cmb_Consulta_Supervisor.ItemData(Cmb_Consulta_Supervisor.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun supervisor", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Empleados
    If Chk_Consulta_Empleado.Value = 1 Then
        If Cmb_Consulta_Empleado.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Control_Bolsa_Horas_Detalles.Empleado_ID = '" & Format(Cmb_Consulta_Empleado.ItemData(Cmb_Consulta_Empleado.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun empleado", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Periodo
    If Chk_Consulta_Periodo.Value = 1 Then
    
        
        If DateDiff("d", Format(Dtp_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Dtp_Consulta_Fecha_Termino, "MM/dd/yyyy")) < 0 Then
            MsgBox "Rango de Fechas Incorrecto", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        Else
            Mi_SQL = Mi_SQL & " AND ((Adm_Control_Bolsa_Horas_Detalles.Fecha_Debe BETWEEN " & Par_Fecha & Format(Dtp_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND " & Par_Fecha & Format(Dtp_Consulta_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha & ")"
            Mi_SQL = Mi_SQL & " OR (Adm_Control_Bolsa_Horas_Detalles.Fecha_Pagado BETWEEN " & Par_Fecha & Format(Dtp_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND " & Par_Fecha & Format(Dtp_Consulta_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha & "))"
        End If
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empleados.Apellido_paterno"
    Set Rs_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Empleados.EOF Then
        While Not Rs_Empleados.EOF
        If Not IsNull(Rs_Empleados.rdoColumns("Supervisor_ID")) Then
            'Consulta de datos del supervisor
            Mi_SQL = "SELECT Cat_Empleados.Empleado_ID,(Cat_Empleados.Apellido_paterno+' '+ Cat_Empleados.Apellido_Materno+' '+ Cat_Empleados.Nombre)  AS Supervior"
            Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Rs_Empleados.rdoColumns("Supervisor_ID"), "00000") & "'"
            Set Rs_Supervisor_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Supervisor_Empleado.EOF Then
                Supervisor = Rs_Supervisor_Empleado.rdoColumns("Supervior")
                Supervisor_ID = Rs_Empleados.rdoColumns("Supervisor_ID")
            End If
       Else
        Supervisor_ID = ""
        Supervisor = ""
       End If
  
            If Tem_Empleado <> Rs_Empleados.rdoColumns("Nombre") Then
            
                Grid_Consulta_Resultados.AddItem "-" & Chr(9) & Rs_Empleados.rdoColumns("Nombre") & Chr(9) & Rs_Empleados.rdoColumns("Empleado_ID")
                Grid_Consulta_Resultados.AddItem "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & Rs_Empleados.rdoColumns("No_Control_Bolsa_Horas") _
                & Chr(9) & Supervisor _
                & Chr(9) & Supervisor_ID _
                & Chr(9) & Rs_Empleados.rdoColumns("Departamento") _
                & Chr(9) & Rs_Empleados.rdoColumns("Departamento_ID") _
                & Chr(9) & Format(Rs_Empleados.rdoColumns("Fecha_Debe"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Empleados.rdoColumns("Horas_Debe") _
                & Chr(9) & Format(Rs_Empleados.rdoColumns("Fecha_Pagado"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Empleados.rdoColumns("Horas_Pagado") _
                & Chr(9) & Rs_Empleados.rdoColumns("Diferencia") _
                & Chr(9) & Rs_Empleados.rdoColumns("Diferencia")
                Total_Diferencia = Rs_Empleados.rdoColumns("Diferencia")
                Tem_Empleado = Rs_Empleados.rdoColumns("Nombre")
            Else
            
                Total_Diferencia = Total_Diferencia + Rs_Empleados.rdoColumns("Horas_Debe") - Rs_Empleados.rdoColumns("Horas_Pagado")
                Grid_Consulta_Resultados.AddItem "" _
                & Chr(9) & "" _
                & Chr(9) & "" _
                & Chr(9) & Rs_Empleados.rdoColumns("No_Control_Bolsa_Horas") _
                & Chr(9) & Supervisor _
                & Chr(9) & Supervisor_ID _
                & Chr(9) & Rs_Empleados.rdoColumns("Departamento") _
                & Chr(9) & Rs_Empleados.rdoColumns("Departamento_ID") _
                & Chr(9) & Format(Rs_Empleados.rdoColumns("Fecha_Debe"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Empleados.rdoColumns("Horas_Debe") _
                & Chr(9) & Format(Rs_Empleados.rdoColumns("Fecha_Pagado"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Empleados.rdoColumns("Horas_Pagado") _
                & Chr(9) & Rs_Empleados.rdoColumns("Diferencia") _
                & Chr(9) & Total_Diferencia
            End If
        
        Rs_Empleados.MoveNext
        Wend
    
       
    'Pone el renglon en negrita
    Grid_Consulta_Resultados.Row = 0
    For Columna = 0 To Grid_Consulta_Resultados.Cols - 1
        Grid_Consulta_Resultados.Col = Columna
        Grid_Consulta_Resultados.CellFontBold = True

    Next
    If Grid_Consulta_Resultados.Rows > 0 Then
        'Pone los tamaños de las celdas
        Grid_Consulta_Resultados.ColWidth(0) = 250     'Agrupar
        Grid_Consulta_Resultados.ColAlignment(0) = flexAlignLeftTop
        Grid_Consulta_Resultados.ColWidth(1) = 3250     'Empleado
        Grid_Consulta_Resultados.ColAlignment(1) = flexAlignLeftTop
        Grid_Consulta_Resultados.ColWidth(2) = 0     'Empleado_ID
        Grid_Consulta_Resultados.ColWidth(3) = 1000     'No.Control
        Grid_Consulta_Resultados.ColAlignment(3) = flexAlignLeftTop
        Grid_Consulta_Resultados.ColWidth(4) = 2800  'Supervisor
        Grid_Consulta_Resultados.ColAlignment(4) = flexAlignLeftTop
        Grid_Consulta_Resultados.ColWidth(5) = 0    'Supervisor_ID
        Grid_Consulta_Resultados.ColWidth(6) = 2800 'Area
        Grid_Consulta_Resultados.ColAlignment(6) = flexAlignLeftTop
        Grid_Consulta_Resultados.ColWidth(7) = 0      'Departamento_ID
        Grid_Consulta_Resultados.ColWidth(8) = 1000   'Dia que Debe
        Grid_Consulta_Resultados.ColWidth(9) = 500    'Hrs.
        Grid_Consulta_Resultados.ColAlignment(9) = 4
        Grid_Consulta_Resultados.ColWidth(10) = 1000  'Dia Pagado
        Grid_Consulta_Resultados.ColWidth(11) = 500   'Hrs.
        Grid_Consulta_Resultados.ColAlignment(11) = 4
        Grid_Consulta_Resultados.ColWidth(12) = 750   'Diferencia.
        Grid_Consulta_Resultados.ColAlignment(12) = 4
        'Consigura para juntar las columnas
        Grid_Consulta_Resultados.MergeCells = flexMergeFree
        Grid_Consulta_Resultados.MergeCol(0) = True
        Grid_Consulta_Resultados.MergeCol(1) = True
        Grid_Consulta_Resultados.FixedRows = 1
    End If

   End If
    MDIFrm_Apl_Principal.MousePointer = 0

Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Consultar_Click()
    Pic_Bolsa_Horas_Consulta.Visible = True
    Pic_Bolsa_Horas_Consulta.ZOrder vbBringToFront
    Chk_Consulta_Periodo.Value = 1
    Grid_Consulta_Resultados.Rows = 0
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN  : Exportar_Exel
'DESCRIPCIÓN           : Exporta los datos a una hoja de exel
'PARÁMETROS            :
'CREO                  :Yazmin Flores Ramirez
'FECHA_CREO            :26-Noviembre-2014
'MODIFICO              :Ana Laura Huichapa
'FECHA_MODIFICO        :27-Febrero-2016
'CAUSA_MODIFICACIÓN    :Se agregó la diferencia de horas al reporte
'*******************************************************************************
Public Sub Exportar_Exel(Grid_Exportar As MSFlexGrid)

Dim Excel As New Excel.Application ' Excel Programa
Dim ExcelWBk As Excel.Workbook ' Libro de Trabajo
Dim ExcelWS As Excel.Worksheet ' Hoja
Dim Ruta_Archivo As String
Dim Fila As Integer
Dim Columna_Excel As Integer
Dim Columna As Integer
Dim Total_Debe As Double
Dim Total_Pagado As Double
Dim Diferencia As Double
Dim Total_Diferencia As Double
Dim Cantidad_Debe  As Integer
Dim Mi_SQL As String
Dim Rs_Empleado As rdoResultset

On Error GoTo MuestraError

Set Excel = CreateObject("Excel.Application") 'Crea el Objeto Excel
Set ExcelWBk = Excel.Workbooks.Add(1)  'Agrega el Libro a Excel
Set ExcelWS = ExcelWBk.Worksheets(1)   'Agrega la Hoja al Libro
ExcelWS.PageSetup.Orientation = xlLandscape

    If Grid_Consulta_Resultados.Rows < 0 Then
        MsgBox "No hay datos que exportar", vbInformation
        Exit Sub
    End If
    ' Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    'On Error GoTo ErrHandler
    ' Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.xls"
    ' Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    ' Display name of selected file
    Ruta_Archivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    If Ruta_Archivo = "" Then
        Exit Sub
    End If
    MDIFrm_Apl_Principal.MousePointer = 11
    Call ExcelWS.Shapes.AddPicture(App.Path & "\SRG_Logo_blue.jpg", 0, True, 0, 16, 200, 63)
    Columna_Excel = 1
    For Fila = 0 To Grid_Exportar.Rows - 1
    Columna_Excel = 1
        
        For Columna = 0 To Grid_Exportar.Cols - 1
        If Fila > 0 Then
            If Grid_Exportar.TextMatrix(Fila, 2) <> "" And Grid_Exportar.TextMatrix(Fila, 0) = "-" Then
                Total_Debe = 0
                Total_Pagado = 0
                Diferencia = 0
                
                'Consulta de datos del Empleado
                Mi_SQL = "SELECT SUM(Horas_Debe)AS Horas_Debe ,SUM(Horas_Pagado) AS Horas_Pagado "
                Mi_SQL = Mi_SQL & " FROM Adm_Control_Bolsa_Horas_Detalles"
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Grid_Exportar.TextMatrix(Fila, 2), "00000") & "'"
                Set Rs_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Empleado.EOF Then
                    Total_Debe = Rs_Empleado.rdoColumns("Horas_Debe")
                    Total_Pagado = Rs_Empleado.rdoColumns("Horas_Pagado")
                    Diferencia = Val(Total_Pagado - Total_Debe)
                End If
                Set Rs_Empleado = Nothing
           
            
                ExcelWS.Cells(Fila + 8, 13) = "Total Hrs. Descanzadas"
                ExcelWS.Cells(Fila + 8, 13).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 13).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 13).Font.Color = vbWhite
                ExcelWS.Cells(Fila + 8, 13).Interior.Color = &HC0C0C0
                ExcelWS.Cells(Fila + 8, 13).ColumnWidth = 18
                ExcelWS.Cells(Fila + 8, 13).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 13).BorderAround 1, xlThin, xlColorIndexAutomatic
                
                ExcelWS.Cells(Fila + 8, 14) = Total_Debe
                ExcelWS.Cells(Fila + 8, 14).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 14).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 14).ColumnWidth = 5
                ExcelWS.Cells(Fila + 8, 14).Font.Color = &H0&
                ExcelWS.Cells(Fila + 8, 14).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 14).BorderAround 1, xlThin, xlColorIndexAutomatic
                    
                ExcelWS.Cells(Fila + 8, 15) = "Total Hrs. Laboradas"
                ExcelWS.Cells(Fila + 8, 15).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 15).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 15).Font.Color = vbWhite
                ExcelWS.Cells(Fila + 8, 15).Interior.Color = &HC0C0C0
                ExcelWS.Cells(Fila + 8, 15).ColumnWidth = 18
                ExcelWS.Cells(Fila + 8, 15).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 15).BorderAround 1, xlThin, xlColorIndexAutomatic
            
                ExcelWS.Cells(Fila + 8, 16) = Total_Pagado
                ExcelWS.Cells(Fila + 8, 16).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 16).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 16).ColumnWidth = 5
                ExcelWS.Cells(Fila + 8, 16).Font.Color = &H0&
                ExcelWS.Cells(Fila + 8, 16).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 16).BorderAround 1, xlThin, xlColorIndexAutomatic
                
                ExcelWS.Cells(Fila + 8, 17) = "Diferencia"
                ExcelWS.Cells(Fila + 8, 17).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 17).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 17).Font.Color = vbWhite
                ExcelWS.Cells(Fila + 8, 17).Interior.Color = &HC0C0C0
                ExcelWS.Cells(Fila + 8, 17).ColumnWidth = 12
                ExcelWS.Cells(Fila + 8, 17).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 17).BorderAround 1, xlThin, xlColorIndexAutomatic
                
                
                ExcelWS.Cells(Fila + 8, 18) = Diferencia
                ExcelWS.Cells(Fila + 8, 18).Font.ColorIndex = 5
                ExcelWS.Cells(Fila + 8, 18).Font.Size = 10
                ExcelWS.Cells(Fila + 8, 18).Font.Color = &H0&
                ExcelWS.Cells(Fila + 8, 18).ColumnWidth = 5
                ExcelWS.Cells(Fila + 8, 18).Interior.Pattern = xlSolid
                ExcelWS.Cells(Fila + 8, 18).BorderAround 1, xlThin, xlColorIndexAutomatic

            
            End If
        End If
        If Grid_Exportar.ColWidth(Columna) <> 0 Then
            ExcelWS.Cells(Fila + 8, Columna_Excel) = Grid_Exportar.TextMatrix(Fila, Columna)
            ExcelWS.Cells(Fila + 8, Columna_Excel).Font.Size = 10
            ExcelWS.Cells(Fila + 8, Columna_Excel).Interior.Pattern = xlSolid
            ExcelWS.Cells(Fila + 8, Columna_Excel).ColumnWidth = 10
            ExcelWS.Cells(Fila + 8, Columna_Excel).BorderAround 1, xlThin, xlColorIndexAutomatic
            ExcelWS.Cells(8, Columna_Excel).Interior.Color = &HC0C0C0
            ExcelWS.Cells(8, Columna_Excel).Font.Color = vbWhite
            ExcelWS.Cells(8, Columna_Excel).Font.Size = 10
            
            
            If Fila > 0 Then
            ExcelWS.Cells(Fila + 8, 12) = ""
            ExcelWS.Cells(Fila + 8, 12).Font.ColorIndex = 5
            ExcelWS.Cells(Fila + 8, 12).Font.Size = 10
            ExcelWS.Cells(Fila + 8, 12).ColumnWidth = 22
            ExcelWS.Cells(Fila + 8, 12).Interior.Pattern = xlSolid
            ExcelWS.Cells(Fila + 8, 12).BorderAround 1, xlThin, xlColorIndexAutomatic
            Else
            ExcelWS.Cells(Fila + 8, 12) = "Firma"
            ExcelWS.Cells(Fila + 8, 12).Font.ColorIndex = 5
            ExcelWS.Cells(Fila + 8, 12).Font.Size = 10
            ExcelWS.Cells(Fila + 8, 12).Font.Color = vbWhite
            ExcelWS.Cells(Fila + 8, 12).Interior.Color = &HC0C0C0
            ExcelWS.Cells(Fila + 8, 12).ColumnWidth = 22
            ExcelWS.Cells(Fila + 8, 12).Interior.Pattern = xlSolid
            ExcelWS.Cells(Fila + 8, 12).BorderAround 1, xlThin, xlColorIndexAutomatic
            End If
            Columna_Excel = Columna_Excel + 1
            
        End If
        
        Next
        
  
    Next Fila
    

    'Campos del encabezado que se deben ajustar
    ExcelWS.Cells(3, 7) = "SRG GLOBAL A GUARDIAN COMPANY"
    ExcelWS.Cells(3, 7).Font.Size = 14
    ExcelWS.Cells(5, 7) = "CONTROL TIEMPOS POR TIEMPOS"
    ExcelWS.Cells(5, 7).Font.Size = 12
    ExcelWS.Cells(6, 1) = "FECHA :"
    ExcelWS.Cells(6, 1).Font.Size = 11
    ExcelWS.Cells(6, 2) = Format(Now, "dd/MMM/yyyy")
   
    'Campos de las firmas
    ExcelWS.Cells(Fila + 11, 2) = "_________________________________"
    ExcelWS.Cells(Fila + 12, 2) = "NOMBRE Y FIRMA DE ELABORO"
    ExcelWS.Cells(Fila + 11, 2).Font.Size = 10
    
    
    ExcelWS.Cells(Fila + 11, 6) = "____________________________________"
    ExcelWS.Cells(Fila + 12, 6) = "NOMBRE Y FIRMA GERENTE DE AREA"
    ExcelWS.Cells(Fila + 11, 6).Font.Size = 10
    
    
    ExcelWS.SaveAs Ruta_Archivo
    
    MDIFrm_Apl_Principal.MousePointer = 0
    If MsgBox("Archivo exportado ¿Desea abrir el reporte?", vbYesNo + vbInformation) = vbYes Then
        Excel.Visible = True
    Else
        ExcelWS.Application.Quit
        
    End If
Exit Sub
MuestraError:
    MsgBox Err.Description
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Btn_Exportar_Click()
      Call Exportar_Exel(Grid_Consulta_Resultados)
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
        If Trim(Txt_Bolsa_Horas_ID.Text) <> "" Then
            Pic_Bolsa_Horas.Enabled = True
            Fra_Bolsa_Horas_General.Enabled = True
        Else
            MsgBox "Seleccione una Control para poder modificar", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        Btn_Modificar.Caption = "Actualizar"
        Btn_Nuevo.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Exportar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        
    Else
        If Grid_Bolsa_Horas.Rows > 0 Then
            Modifica_Control_Bolsa_Horas
        Else
            MsgBox "Ingrese por lo menos un registro para el control de tiempos ", vbOKOnly + vbInformation, Me.Caption
        End If
    
    End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Exportar.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        
        Txt_Bolsa_Horas_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Control_Bolsa_Horas", "No_Control_Bolsa_Horas"), "0000000000")
        Dtp_Bolsa_Horas_Fecha.Value = Now
        Dtp_Fecha_Debe.Value = Now
        Dtp_Fecha_Pagado.Value = Now
        Fra_Bolsa_Horas_General.Enabled = True
        Fra_Observaciones.Enabled = True
        Fra_Control_Bolsa_Horas.Enabled = True

    Else
        If Grid_Bolsa_Horas.Rows > 0 Then
            Alta_Control_Bolsa_Horas
        Else
            MsgBox "Ingrese por lo menos un registro para el control de tiempos ", vbOKOnly + vbInformation, Me.Caption
        End If

    End If
End Sub
Private Sub Limpia_Informacion_Empleado()
    Cmb_Departamento.ListIndex = -1
    Txt_Departmento.Text = ""
    Cmb_Supervisor.ListIndex = -1
    Txt_Supervisor.Text = ""
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Bolsa_Horas_Consulta.Visible = False
End Sub

Private Sub Btn_Salir_Click()
 If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Exportar.Enabled = True
        Btn_Consultar.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        Cmb_Bolsa_Horas_Empleado.Clear
        Grid_Bolsa_Horas.Rows = 0
 End If
End Sub

Private Sub Chk_Consulta_Departamentos_Click()
    If Chk_Consulta_Departamentos.Value = 1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Consulta_Departamento, 1, "Nombre")
    Else
        Cmb_Consulta_Departamento.Clear
    End If
End Sub

Private Sub Chk_Consulta_Departamentos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_Consulta_Departamento.Clear
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Consulta_Departamento, 0, "")
    Else
        Conectar_Ayudante.Quitar_Caracter_Raro (KeyAscii)
    End If
End Sub

Private Sub Chk_Consulta_Empleado_Click()
    If Chk_Consulta_Empleado.Value = 1 Then
        If Chk_Consulta_Supervisor.Value = 1 And Cmb_Consulta_Supervisor.ListIndex > -1 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Consulta_Empleado, 1, "Apellido_paterno", " AND Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Consulta_Supervisor.ItemData(Cmb_Consulta_Supervisor.ListIndex), "00000") & "'", False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Consulta_Empleado, 1, "Apellido_paterno", " AND Estatus = 'A'", False, "TODOS")
        End If
    Else
        Cmb_Consulta_Empleado.Clear
    End If
End Sub

Private Sub Chk_Consulta_Empleado_KeyPress(KeyAscii As Integer)
    If Chk_Consulta_Empleado.Value = 1 Then
        If Chk_Consulta_Supervisor.Value = 1 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Consulta_Supervisor.ItemData(Cmb_Consulta_Supervisor.ListIndex), "00000") & "'", Cmb_Consulta_Empleado, 0, 0, False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A'", Cmb_Consulta_Empleado, 0, 0, False, "TODOS")
        End If
    Else
        Cmb_Consulta_Empleado.Clear
    End If
End Sub

Private Sub Chk_Consulta_Supervisor_Click()
    If Chk_Consulta_Supervisor.Value = 1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Consulta_Supervisor, 1, "Apellido_paterno", "AND Tipo = 'S' AND Estatus = 'A'", False, "TODOS")
    Else
        Cmb_Consulta_Supervisor.Clear
    End If
End Sub


Private Sub Cmb_Bolsa_Horas_Empleado_Click()
Dim Rs_Consuta_Cat_Empleados As rdoResultset     'Informcion de los empleados
    
    If Cmb_Bolsa_Horas_Empleado.ListIndex > -1 Then
        Call Limpia_Informacion_Empleado
        'Llena los datos del departamento, Empresa y Supervisor
        Mi_SQL = "SELECT ISNULL(Supervisor_ID,'') as Supervisor_ID ,ISNULL(Empresa_ID,'') as Empresa_ID,ISNULL(Departamento_ID,'') as Departamento_ID FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID = '" & Format(Cmb_Bolsa_Horas_Empleado.ItemData(Cmb_Bolsa_Horas_Empleado.ListIndex), "00000") & "'"
        Set Rs_Consuta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consuta_Cat_Empleados
            If Not .EOF Then
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Departamento)
                Txt_Departmento.Text = Cmb_Departamento.Text
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Supervisor)
                Txt_Supervisor.Text = Cmb_Supervisor.Text
                Txt_Supervisor.Enabled = False
                Txt_Departmento.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub Cmb_Bolsa_Horas_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Limpia_Informacion_Empleado
        If IsNumeric(Cmb_Bolsa_Horas_Empleado.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND No_Tarjeta='" & Cmb_Bolsa_Horas_Empleado.Text & "'", Cmb_Bolsa_Horas_Empleado, 0, "No_Tarjeta")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND Nombre LIKE '%" & Trim(Cmb_Bolsa_Horas_Empleado.Text) & "%' OR Apellido_Paterno LIKE '%" & Trim(Cmb_Bolsa_Horas_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Bolsa_Horas_Empleado.Text) & "%'", Cmb_Bolsa_Horas_Empleado, 0, "Apellido_Paterno")
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Bolsa_Horas_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Bolsa_Horas_Empleado, KeyCode)
End Sub
Private Sub Cmb_Consulta_Empleado_Click()
Dim Rs_Consuta_Cat_Empleados As rdoResultset     'Informcion de los empleados

    If Cmb_Bolsa_Horas_Empleado.ListIndex > -1 Then
        Call Limpia_Informacion_Empleado
        'Llena los datos del departamento, Empresa y Supervisor
        Mi_SQL = "SELECT ISNULL(Supervisor_ID,'') as Supervisor_ID ,ISNULL(Empresa_ID,'') as Empresa_ID,ISNULL(Departamento_ID,'') as Departamento_ID FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID = '" & Format(Cmb_Bolsa_Horas_Empleado.ItemData(Cmb_Bolsa_Horas_Empleado.ListIndex), "00000") & "'"
        Set Rs_Consuta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consuta_Cat_Empleados
            If Not .EOF Then
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Departamento)
                Txt_Departmento.Text = Cmb_Departamento.Text
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Supervisor)
                Txt_Supervisor.Text = Cmb_Supervisor.Text
                Txt_Supervisor.Enabled = False
                Txt_Departmento.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub Cmb_Consulta_Empleado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        Call Limpia_Informacion_Empleado
        If IsNumeric(Cmb_Consulta_Empleado.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND No_Tarjeta='" & Cmb_Consulta_Empleado.Text & "'", Cmb_Consulta_Empleado, 0, "No_Tarjeta")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND Nombre LIKE '%" & Trim(Cmb_Consulta_Empleado.Text) & "%' OR Apellido_Paterno LIKE '%" & Trim(Cmb_Consulta_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Consulta_Empleado.Text) & "%'", Cmb_Consulta_Empleado, 0, "Apellido_Paterno")
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Consulta_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Consulta_Empleado, KeyCode)
End Sub

Private Sub Form_Load()

Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Departamento, 0, "Nombre")
Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S' ", Cmb_Supervisor, 0, "Apellido_Paterno")
Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND Tipo = 'S'", Cmb_Bolsa_Horas_Empleado, 0, "Nombre")
Dtp_Bolsa_Horas_Fecha.Value = Now
End Sub

Private Sub Grid_Bolsa_Horas_Click()
    With Grid_Bolsa_Horas
        If .Rows > 1 Then

            Txt_Supervisor.Text = Trim(Grid_Bolsa_Horas.TextMatrix(.RowSel, 2))
            Txt_Departmento.Text = Trim(Grid_Bolsa_Horas.TextMatrix(.RowSel, 4))
            Cmb_Bolsa_Horas_Empleado.Text = Trim(Grid_Bolsa_Horas.TextMatrix(.RowSel, 1))
            Dtp_Fecha_Debe.Value = Format(Grid_Bolsa_Horas.TextMatrix(.RowSel, 6), "MM/dd/yyyy")
            Txt_Hrs_Debe.Text = Val(Grid_Bolsa_Horas.TextMatrix(.RowSel, 7))
            Dtp_Fecha_Pagado.Value = Format(Grid_Bolsa_Horas.TextMatrix(.RowSel, 8), "MM/dd/yyyy")
            Txt_Hrs_Pagado.Text = Val(Grid_Bolsa_Horas.TextMatrix(.RowSel, 9))
        End If
    End With
End Sub

Private Sub Grid_Consulta_Resultados_Click()
Dim Fila As Integer
    If Grid_Consulta_Resultados.Rows > 1 Then
        If Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 0) = "+" Or Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 0) = "-" Then
            If Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 0) = "+" Then
                Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 0) = "-"
                For Fila = (Grid_Consulta_Resultados.RowSel + 1) To Grid_Consulta_Resultados.Rows - 1
                    'Valida que sea un renglon de submenu para cambiar el tamaño
                    If Trim(Grid_Consulta_Resultados.TextMatrix(Fila, 0)) = "" Then
                        Grid_Consulta_Resultados.RowHeight(Fila) = 250 'Cambia el tamaño de la fila
                    Else
                        Exit For
                    End If
                Next
            Else
                Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 0) = "+"
                For Fila = (Grid_Consulta_Resultados.RowSel + 1) To Grid_Consulta_Resultados.Rows - 1
                    'Valida que sea un renglon de submenu para cambiar el tamaño
                    If Trim(Grid_Consulta_Resultados.TextMatrix(Fila, 0)) = "" Then
                        Grid_Consulta_Resultados.RowHeight(Fila) = 0 'Cambia el tamaño de la fila
                    Else
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Grid_Consulta_Resultados_DblClick()
    If Grid_Consulta_Resultados.Rows > 1 Then
        Muestra_Datos_Control_Bolsa_Horas Trim(Grid_Consulta_Resultados.TextMatrix(Grid_Consulta_Resultados.RowSel, 3))
    End If
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Muestra_Datos_Control_Bolsa_Horas
'DESCRIPCION: Carga en pantalla los datos de la bolsa de horas
'PARAMETROS : No_Control.- Número de la solicitud a mostrar en la pantalla
'CREO       : Yazmin Flores Ramirez
'FECHA_CREO : 28-Noviembre-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Muestra_Datos_Control_Bolsa_Horas(No_Control As String)
On Error GoTo HANDLER
Dim Mi_SQL As String
Dim Rs_Control As rdoResultset              'Consulta todos los datos de un pedido especifico
Dim Rs_Detalles_Control As rdoResultset     'Consulta todos los productos del pedido
Dim Rs_Supervisor_Empleado As rdoResultset
Dim Cantidad_Total As Double
Dim Cont_Filas As Integer
Dim Supervisor_ID As String
Dim Supervisor As String
    
   
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Adm_Control_Bolsa_Horas"
    Mi_SQL = Mi_SQL & " WHERE No_Control_Bolsa_Horas = '" & Format(No_Control, "0000000000") & "'"
    
    Set Rs_Control = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'LLena el grid con los datos el resultado de la busqueda anterior
    If Not Rs_Control.EOF Then
        With Rs_Control
            Txt_Bolsa_Horas_ID.Text = .rdoColumns("No_Control_Bolsa_Horas")
            Dtp_Bolsa_Horas_Fecha.Value = .rdoColumns("Fecha_Solicitud")
            
            If Not IsNull(.rdoColumns("Observaciones")) Then
                Txt_Observaciones.Text = .rdoColumns("Observaciones")
            Else
                Txt_Observaciones.Text = ""
            End If
            
            'Llena los detalles
            Mi_SQL = "SELECT Adm_Control_Bolsa_Horas_Detalles.*,Adm_Control_Bolsa_Horas_Detalles.Empleado_ID,Adm_Control_Bolsa_Horas_Detalles.Departamento_ID,(Cat_Empleados.Apellido_paterno+' '+ Cat_Empleados.Apellido_Materno+' '+ Cat_Empleados.Nombre) as Nombre ,"
            Mi_SQL = Mi_SQL & " Cat_Departamentos.Nombre as Departamento "
            Mi_SQL = Mi_SQL & " FROM Adm_Control_Bolsa_Horas_Detalles,Cat_Empleados, Cat_Departamentos"
            Mi_SQL = Mi_SQL & " WHERE Adm_Control_Bolsa_Horas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Control_Bolsa_Horas_Detalles.Departamento_ID=Cat_Departamentos.Departamento_ID"
            Mi_SQL = Mi_SQL & " AND Adm_Control_Bolsa_Horas_Detalles.No_Control_Bolsa_Horas='" & .rdoColumns("No_Control_Bolsa_Horas") & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Cat_Empleados.Apellido_paterno"
            Set Rs_Detalles_Control = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Grid_Bolsa_Horas.Rows = 0
            Grid_Bolsa_Horas.Cols = 10
            Grid_Bolsa_Horas.AddItem "Empleado_ID" & Chr(9) & "Empleado" & Chr(9) & _
                    "Supervisor" & Chr(9) & "" & Chr(9) & "Area" & Chr(9) & "" & Chr(9) & "Dia que Debe" & Chr(9) & "Hrs" & Chr(9) & "Dia Pagado" & Chr(9) & "Hrs"
            
            While Not Rs_Detalles_Control.EOF
            
            If Not IsNull(Rs_Detalles_Control.rdoColumns("Supervisor_ID")) Then
                'Consulta de datos del supervisor
                Mi_SQL = "SELECT Cat_Empleados.Empleado_ID,(Cat_Empleados.Apellido_paterno+' '+ Cat_Empleados.Apellido_Materno+' '+ Cat_Empleados.Nombre)  AS Supervior"
                Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Rs_Detalles_Control.rdoColumns("Supervisor_ID"), "00000") & "'"
                Set Rs_Supervisor_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Supervisor_Empleado.EOF Then
                    Supervisor = Rs_Supervisor_Empleado.rdoColumns("Supervior")
                    Supervisor_ID = Rs_Detalles_Control.rdoColumns("Supervisor_ID")
                End If
            Else
                Supervisor_ID = ""
                Supervisor = ""
            End If
            
                'Agrega el producto al grid de productos del pedido
                
                Grid_Bolsa_Horas.AddItem Rs_Detalles_Control.rdoColumns("Empleado_ID") _
                & Chr(9) & Rs_Detalles_Control.rdoColumns("Nombre") _
                & Chr(9) & Supervisor _
                & Chr(9) & Supervisor_ID _
                & Chr(9) & Rs_Detalles_Control.rdoColumns("Departamento") _
                & Chr(9) & Rs_Detalles_Control.rdoColumns("Departamento_ID") _
                & Chr(9) & Format(Rs_Detalles_Control.rdoColumns("Fecha_Debe"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Detalles_Control.rdoColumns("Horas_Debe") _
                & Chr(9) & Format(Rs_Detalles_Control.rdoColumns("Fecha_Pagado"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Detalles_Control.rdoColumns("Horas_Pagado")
                
                Grid_Bolsa_Horas.FixedRows = 1
                Rs_Detalles_Control.MoveNext
            Wend
            
            Grid_Bolsa_Horas.ColWidth(0) = 0  'Empleado id
            Grid_Bolsa_Horas.ColWidth(1) = 3200  'Empleado
            Grid_Bolsa_Horas.ColWidth(2) = 3200  'Supervisor
            Grid_Bolsa_Horas.ColWidth(3) = 0  'Supervisor id
            Grid_Bolsa_Horas.ColWidth(4) = 3200  'Area
            Grid_Bolsa_Horas.ColWidth(5) = 0  'Area id
            Grid_Bolsa_Horas.ColWidth(6) = 1100  'Dia Debe
            Grid_Bolsa_Horas.ColWidth(7) = 500  'Hrs
            Grid_Bolsa_Horas.ColAlignment(7) = flexAlignCenterCenter
            Grid_Bolsa_Horas.ColWidth(8) = 1100  'Dia Pagado
            Grid_Bolsa_Horas.ColWidth(9) = 500  'Hrs
            Grid_Bolsa_Horas.ColAlignment(9) = flexAlignCenterCenter
            Rs_Detalles_Control.Close
        End With
    End If
    Rs_Control.Close
    Btn_Regresar_Click
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Private Sub Txt_Hrs_Debe_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Hrs_Debe, True)
End Sub

Private Sub Txt_Hrs_Pagado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Hrs_Pagado, True)
End Sub

Private Sub Txt_Observaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
