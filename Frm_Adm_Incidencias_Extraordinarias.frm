VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Adm_Incidencias_Extraordinarias 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   645
      Left            =   2493
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6780
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   645
      Left            =   45
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "A"
      Top             =   5985
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Cancelar"
      Height          =   645
      Left            =   3105
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "B"
      Top             =   5985
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   645
      Left            =   1575
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "M"
      Top             =   5985
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   645
      Left            =   6165
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5985
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   645
      Left            =   4635
      Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "C"
      Top             =   5985
      Width           =   930
   End
   Begin VB.PictureBox Pic_Solicitud_Permisos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5910
      Left            =   15
      ScaleHeight     =   5910
      ScaleWidth      =   7125
      TabIndex        =   35
      Top             =   0
      Width           =   7125
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   585
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         Left            =   90
         TabIndex        =   49
         Top             =   4725
         Width           =   6990
         Begin VB.TextBox Txt_Permisos_Observaciones 
            Height          =   675
            Left            =   90
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   225
            Width           =   6795
         End
      End
      Begin VB.Frame Fra_Permisos_Tipo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incidencias Extraordinarias"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   90
         TabIndex        =   48
         Top             =   3645
         Width           =   7035
         Begin VB.TextBox Txt_Permiso_Horas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5265
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   8
            Top             =   675
            Width           =   1245
         End
         Begin VB.ComboBox Cmb_Permisos_Incidencias_Extraordinarias 
            Height          =   315
            ItemData        =   "Frm_Adm_Incidencias_Extraordinarias.frx":213C
            Left            =   1530
            List            =   "Frm_Adm_Incidencias_Extraordinarias.frx":213E
            TabIndex        =   7
            Top             =   315
            Width           =   5415
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "hrs."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6600
            TabIndex        =   58
            Top             =   720
            Width           =   270
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Incidencia Extraord."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   57
            Top             =   375
            Width           =   1380
         End
      End
      Begin VB.Frame Fra_Permiso_Datos_Solicitante 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del Solicitante"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   90
         TabIndex        =   40
         Top             =   1395
         Width           =   7035
         Begin VB.TextBox Txt_Permisos_Departamento 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4710
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   270
            Width           =   2220
         End
         Begin VB.TextBox Txt_Permiso_Dias_Sueldo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5070
            MaxLength       =   2
            TabIndex        =   6
            Top             =   1890
            Width           =   1245
         End
         Begin VB.ComboBox Cmb_Permisos_Empleado 
            Height          =   315
            Left            =   1395
            TabIndex        =   3
            Top             =   960
            Width           =   5535
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Inicio 
            Height          =   315
            Left            =   2175
            TabIndex        =   4
            Top             =   1470
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd/MMM/yyyy"
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Termino 
            Height          =   315
            Left            =   5070
            TabIndex        =   5
            Top             =   1470
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd/MMM/yyyy"
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin VB.TextBox Txt_Permisos_Empresa 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   270
            Width           =   2070
         End
         Begin VB.ComboBox Cmb_Permisos_Empresa 
            Height          =   315
            Left            =   1395
            TabIndex        =   17
            Top             =   270
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox Txt_Permisos_Supervisor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   615
            Width           =   5520
         End
         Begin VB.ComboBox Cmb_Permisos_Supervisor 
            Height          =   315
            Left            =   1530
            TabIndex        =   18
            Top             =   615
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.ComboBox Cmb_Permisos_Departamento 
            Height          =   315
            Left            =   4710
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   270
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   3540
            TabIndex        =   64
            Top             =   330
            Width           =   1050
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   50
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha o fechas que solicita"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   47
            Top             =   1350
            Width           =   1815
         End
         Begin VB.Label Lbl_Adm_Inasistencias_Dias_Porcentaje 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4260
            TabIndex        =   46
            Top             =   1980
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4260
            TabIndex        =   45
            Top             =   1530
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   44
            Top             =   1530
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dias"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6630
            TabIndex        =   43
            Top             =   1935
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   42
            Top             =   1020
            Width           =   690
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.Frame Fra_Permiso_Datos_Generales 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   915
         Left            =   90
         TabIndex        =   38
         Top             =   450
         Width           =   7035
         Begin VB.ComboBox Cmb_Permisos_Estatus 
            Height          =   315
            ItemData        =   "Frm_Adm_Incidencias_Extraordinarias.frx":2140
            Left            =   1395
            List            =   "Frm_Adm_Incidencias_Extraordinarias.frx":214D
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   540
            Width           =   2070
         End
         Begin VB.TextBox Txt_Adm_Permisos_No_Movimiento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   180
            Width           =   2070
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Solicitud 
            Height          =   315
            Left            =   4770
            TabIndex        =   2
            Top             =   540
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "ddd dd/MMM/yyyy"
            Format          =   124059651
            CurrentDate     =   40030
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Movimiento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   55
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3555
            TabIndex        =   39
            Top             =   600
            Width           =   450
         End
      End
      Begin VB.PictureBox Pic_Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   450
         TabIndex        =   36
         Top             =   0
         Width           =   450
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   1020
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INCIDENCIAS EXTRAORDINARIAS"
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
         Left            =   1725
         TabIndex        =   37
         Top             =   0
         Width           =   3945
      End
   End
   Begin VB.PictureBox Pic_Solicitud_Permisos_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   7125
      TabIndex        =   51
      Top             =   495
      Width           =   7125
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
         Height          =   2715
         Left            =   45
         TabIndex        =   53
         Top             =   75
         Width           =   7080
         Begin VB.CommandButton Btn_Exportar 
            Caption         =   "Exportar"
            Height          =   645
            Left            =   6060
            Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":216F
            Style           =   1  'Graphical
            TabIndex        =   60
            Tag             =   "A"
            Top             =   1095
            UseMaskColor    =   -1  'True
            Width           =   930
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Empleado 
            Height          =   315
            Left            =   1755
            TabIndex        =   26
            Top             =   1467
            Width           =   4200
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Empresa 
            Height          =   315
            Left            =   1755
            TabIndex        =   20
            Top             =   225
            Width           =   4200
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Supervisor 
            Height          =   315
            Left            =   1755
            TabIndex        =   24
            Top             =   1053
            Width           =   4200
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Departamento 
            Height          =   315
            Left            =   1755
            TabIndex        =   22
            Top             =   639
            Width           =   4200
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Tipo_Permiso 
            Height          =   315
            ItemData        =   "Frm_Adm_Incidencias_Extraordinarias.frx":26F9
            Left            =   1755
            List            =   "Frm_Adm_Incidencias_Extraordinarias.frx":2709
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1881
            Width           =   4200
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Tipo_Permiso 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Incidencia Extraord."
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
            TabIndex        =   27
            Top             =   1881
            Width           =   1770
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Departamentos 
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
            TabIndex        =   21
            Top             =   639
            Width           =   1410
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Supervisor 
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
            TabIndex        =   23
            Top             =   1053
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Regresar 
            Cancel          =   -1  'True
            Caption         =   "Regresar"
            Height          =   645
            Left            =   6060
            Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":2748
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1965
            UseMaskColor    =   -1  'True
            Width           =   930
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Periodo 
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
            TabIndex        =   29
            Top             =   2295
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Empleado 
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
            TabIndex        =   25
            Top             =   1467
            Width           =   1050
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Empresa 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
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
            TabIndex        =   19
            Top             =   225
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   645
            Left            =   6060
            Picture         =   "Frm_Adm_Incidencias_Extraordinarias.frx":2CD2
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "C"
            Top             =   225
            Width           =   930
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Permisos_Consulta_Fecha_Termino 
            Height          =   315
            Left            =   4230
            TabIndex        =   31
            Top             =   2295
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Permisos_Consulta_Fecha_Inicio 
            Height          =   315
            Left            =   1755
            TabIndex        =   30
            Top             =   2295
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin MSComctlLib.ProgressBar Prbar_Exportacion 
            Height          =   165
            Left            =   6060
            TabIndex        =   61
            Top             =   1740
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Lbl_Progreso_Exportacion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exportando..."
            Height          =   195
            Left            =   6120
            TabIndex        =   62
            Top             =   1440
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   3780
            TabIndex        =   54
            Top             =   2355
            Width           =   135
         End
      End
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
         Height          =   3360
         Left            =   45
         TabIndex        =   52
         Top             =   2775
         Width           =   7080
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Permisos_Consulta_Resultados 
            Height          =   2955
            Left            =   60
            TabIndex        =   34
            Top             =   270
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   5212
            _Version        =   393216
            Rows            =   0
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
   End
End
Attribute VB_Name = "Frm_Adm_Incidencias_Extraordinarias"
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
Dim Archivo_Reporte_Abierto As Boolean
Private Sub Btn_Buscar_Click()
    Consulta_Permisos
End Sub

Private Sub Btn_Consultar_Click()
    Pic_Solicitud_Permisos_Consulta.Visible = True
    Pic_Solicitud_Permisos_Consulta.ZOrder vbBringToFront
    Chk_Adm_Permisos_Consulta_Periodo.Value = 1
    Grid_Adm_Permisos_Consulta_Resultados.Rows = 0
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo HANDLER
    If MsgBox("¿Esta seguro de cancelar el permiso?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Select Case Operacion
            Case "Permisos": 'Catalogo de dias
                If Trim(Txt_Adm_Permisos_No_Movimiento.Text) <> "" Then
                    Dim Rs_Modifica_Adm_Movimiento
                    Dim Referencia As String
                    'Valida que el movimiento no este siendo utilizado
                    Mi_SQL = "SELECT Referencia FROM Adm_Asistencias"
                    Mi_SQL = Mi_SQL & " WHERE Referencia = '" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & "'"
                    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'E'"
                    Referencia = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Referencia")
                    If Referencia <> "" Then
                        MsgBox "El permiso esta siendo utilizado para el control de asistencias," & vbCrLf & _
                               "Si desea cancelar el permiso, " & vbCrLf & _
                               "debera primero modificar la asistencia a la que esta asociado", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                    Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
                    Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & "'"
                    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia='E'"
                    Set Rs_Modifica_Adm_Movimiento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                    If Not Rs_Modifica_Adm_Movimiento.EOF Then
                        Rs_Modifica_Adm_Movimiento.Edit
                            Rs_Modifica_Adm_Movimiento.rdoColumns("Estatus") = "C"
                        Rs_Modifica_Adm_Movimiento.Update
                    End If
                    Rs_Modifica_Adm_Movimiento.Close
                    Cmb_Permisos_Estatus.ListIndex = 1
                    'Quita los datos del usuario contenidos en el Grid
                    MsgBox "El permiso ha sido cancelado", vbInformation + vbOKOnly, Me.Caption
                Else
                    MsgBox "Seleccione un Permisos para poder cancelar", vbInformation + vbOKOnly, Me.Caption
                End If
        End Select
        'Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia los textos de la forma
    End If
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub


Private Sub Btn_Exportar_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String

On Error GoTo HANDLER
    Cmd_Exportar.CancelError = True
    Cmd_Exportar.DialogTitle = "Seleccione el directorio"
    Cmd_Exportar.Flags = cdlOFNHideReadOnly
    Cmd_Exportar.Filter = "Archivos de Excel(*.xls)|*.xls"
    Cmd_Exportar.FilterIndex = 2
    Cmd_Exportar.FileName = "Incidencias_Extraordinarias" & ".xls"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
        Call Exportar_Excel(Ruta_Temporal & "Incidencias_Extraordinarias" & "xls.txt", Ruta_Exportacion, Prbar_Exportacion, Lbl_Progreso_Exportacion, Me)
    End If
    Exit Sub
HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Imprimir_Click()
    If Trim(Txt_Adm_Permisos_No_Movimiento) = "" Then
        MsgBox "Seleccione una incidencia", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = 11
    Rpt_Reporte.Reset
    Rpt_Reporte.WindowTitle = "Incidencia Extraordinaria"
    Rpt_Reporte.Connect = Conexion_Base.Connect
    Rpt_Reporte.ReportFileName = App.Path & "\Reportes\Incidencia_Extraordinaria.frmx"
    Rpt_Reporte.ParameterFields(1) = "No_Movimiento;" & Trim(Txt_Adm_Permisos_No_Movimiento) & ";true"
    Rpt_Reporte.Destination = crptToWindow
    Rpt_Reporte.WindowState = crptMaximized
    Rpt_Reporte.ProgressDialog = True
    Rpt_Reporte.WindowShowPrintBtn = True
    Rpt_Reporte.WindowShowPrintSetupBtn = True
    Rpt_Reporte.PrintReport
    Me.MousePointer = 0
End Sub

Private Sub Btn_Modificar_Click()
Dim Tipo_Permiso As Boolean
    If Btn_Modificar.Caption = "Modificar" Then
        Select Case Operacion
            Case "Permisos":
                If Trim(Txt_Adm_Permisos_No_Movimiento.Text) <> "" Then
                    Pic_Solicitud_Permisos.Enabled = True
                    Dtp_Permiso_Fecha_Inicio.SetFocus
                    Cmb_Permisos_Estatus.Enabled = True
                Else
                    MsgBox "Seleccione una Permiso para poder modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
        End Select
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = False
        Btn_Nuevo.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Imprimir.Enabled = False
        Btn_Salir.Caption = "Regresar"
    Else
        Select Case Operacion
            Case "Permisos": 'Captura de Vacaciones
                Tipo_Permiso = False
                'valida que la informacion este completa
                If Cmb_Permisos_Departamento.ListIndex > -1 Then    'Valida el departamento
                    If Cmb_Permisos_Empresa.ListIndex > -1 Then     'valida la empresa
'                        If Cmb_Permisos_Supervisor.ListIndex > -1 Then  'Valida el supervisor
                            If Cmb_Permisos_Empleado.ListIndex > -1 Then    'Valida el empleado
                                If Cmb_Permisos_Incidencias_Extraordinarias.ListIndex > -1 Then
                                    If Val(Txt_Permiso_Dias_Sueldo.Text) > 0 Then  'valida los dias de permiso
                                        If Simbologia = "A" Then
                                            If Val(Txt_Permiso_Horas.Text) <= 0 Then
                                                MsgBox "Debe agregar las horas de acuerdo para la asistencia", vbInformation + vbOKOnly, Me.Caption
                                                Txt_Permiso_Horas.SetFocus
                                                Exit Sub
                                            End If
                                        End If
                                        Modifica_Permisos
                                    Else
                                        MsgBox "No ha ingresado los dias para el permiso", vbInformation + vbOKOnly, Me.Caption
                                        Dtp_Permiso_Fecha_Termino.SetFocus
                                    End If
                                Else
                                    MsgBox "Debe seleccionar la incidencia extraordinaria", vbInformation + vbOKOnly, Me.Caption
                                    Cmb_Permisos_Incidencias_Extraordinarias.SetFocus
                                End If
                            Else
                                MsgBox "Debe proporcionar el empleado", vbInformation + vbOKOnly, Me.Caption
                                Cmb_Permisos_Empleado.SetFocus
                            End If
'                        Else
'                            MsgBox "Supervisor no asignado a empleado," & vbCrLf & _
'                                "debera configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
'                            'Cmb_Permisos_Supervisor.SetFocus
'                        End If
                    Else
                        MsgBox "Empresa no asignada a empleado," & vbCrLf & _
                               "deberá configurarla en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                        'Cmb_Permisos_Empresa.SetFocus
                    End If
                Else
                    MsgBox "Departamento no asinado a empleado," & vbCrLf & _
                               "deberá configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                    'Cmb_Permisos_Departamento.SetFocus
                End If
        End Select
    End If
End Sub

'***************************************Acciones de Botones*********************************
Private Sub Btn_Nuevo_Click()
Dim Tipo_Permiso As Boolean     'Indica si se ha seleccionado algun tipo de permiso
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Imprimir.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        Select Case Operacion
            Case "Permisos": 'Captura de Inasistencias
                'Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias WHERE Tipo_Incidencia='E'", "No_Movimiento"), "0000000000")
                Pic_Solicitud_Permisos.Enabled = True
                Cmb_Permisos_Estatus.ListIndex = 0
                Dtp_Permiso_Fecha_Inicio.Value = Now
                Dtp_Permiso_Fecha_Termino.Value = Now
                Dtp_Permiso_Fecha_Solicitud.Value = Now
        End Select
    Else
        Select Case Operacion
            Case "Permisos": 'Captura de Vacaciones
                Tipo_Permiso = False
                'valida que la informacion este completa
                If Cmb_Permisos_Departamento.ListIndex > -1 Then    'Valida el departamento
                    If Cmb_Permisos_Empresa.ListIndex > -1 Then     'valida la empresa
'                        If Cmb_Permisos_Supervisor.ListIndex > -1 Then  'Valida el supervisor
                            If Cmb_Permisos_Empleado.ListIndex > -1 Then    'Valida el empleado
                                If Cmb_Permisos_Incidencias_Extraordinarias.ListIndex > -1 Then
                                    If Val(Txt_Permiso_Dias_Sueldo.Text) > 0 Then  'valida los dias de permiso
                                        If Simbologia = "A" Then
                                            If Val(Txt_Permiso_Horas.Text) <= 0 Then
                                                MsgBox "Debe agregar las horas de acuerdo para la asistencia", vbInformation + vbOKOnly, Me.Caption
                                                Txt_Permiso_Horas.SetFocus
                                                Exit Sub
                                            End If
                                        End If
                                        Alta_Permisos
                                    Else
                                        MsgBox "No ha ingresado los dias para el permiso", vbInformation + vbOKOnly, Me.Caption
                                        Dtp_Permiso_Fecha_Termino.SetFocus
                                    End If
                                Else
                                    MsgBox "Debe seleccionar la incidencia extraordinaria", vbInformation + vbOKOnly, Me.Caption
                                    Cmb_Permisos_Incidencias_Extraordinarias.SetFocus
                                End If
                            Else
                                MsgBox "Debe proporcionar el empleado", vbInformation + vbOKOnly, Me.Caption
                                Cmb_Permisos_Empleado.SetFocus
                            End If
'                        Else
'                            MsgBox "Supervisor no asignado a empleado," & vbCrLf & _
'                                "debera configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
'                        End If
                    Else
                        MsgBox "Empresa no asignada a empleado," & vbCrLf & _
                               "deberá configurarla en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                        'Cmb_Permisos_Empresa.SetFocus
                    End If
                Else
                    MsgBox "Departamento no asinado a empleado," & vbCrLf & _
                               "deberá configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                    'Cmb_Permisos_Departamento.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Solicitud_Permisos_Consulta.Visible = False
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Eliminar.Enabled = True
        Btn_Consultar.Enabled = True
        Btn_Imprimir.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
      
        Select Case Operacion
            Case "Permisos": 'Catalogo de Dias no laborales
                Pic_Solicitud_Permisos.Enabled = False
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Sub SubMenu_Adm_Incidencias_Extraordinarias", Me)
    
        End Select
    End If
End Sub

'*****************************************Fin Acciones de Botones*****************************
Public Sub Inicializa()
    Dtp_Permiso_Fecha_Solicitud.Value = Now
    Simbologia = ""
    SubSimbologia = ""
    Dtp_Permiso_Fecha_Inicio.Value = Now
    Dtp_Permiso_Fecha_Termino.Value = Now
    Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value = Now
    Dtp_Adm_Permisos_Consulta_Fecha_Termino.Value = Now
    Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Permisos_Departamento, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Permisos_Empresa, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Permisos_Consulta_Empresa, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S'", Cmb_Permisos_Supervisor, 0, "Apellido_Paterno")
    'Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID, Descripcion", "Cat_Tipos_Faltas", Cmb_Permisos_Incidencias_Extraordinarias, 0, "Descripcion")
    Call Cmb_Permisos_Incidencias_Extraordinarias_KeyPress(13)
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Departamentos_Click()
    If Chk_Adm_Permisos_Consulta_Departamentos.Value = 1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Adm_Permisos_Consulta_Departamento, 1, "Nombre")
    Else
        Cmb_Adm_Permisos_Consulta_Departamento.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Departamentos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cmb_Adm_Permisos_Consulta_Departamento.Clear
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Adm_Permisos_Consulta_Departamento, 0, "")
    Else
        Conectar_Ayudante.Quitar_Caracter_Raro (KeyAscii)
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Empleado_Click()
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 Then
        If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", "AND Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", "AND Estatus = 'A'", False, "TODOS")
        End If
    Else
        Cmb_Adm_Permisos_Consulta_Empleado.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Empleado_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 Then
        If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 And Cmb_Adm_Permisos_Consulta_Empresa.ListIndex > -1 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", Cmb_Adm_Permisos_Consulta_Empleado, 0, 0, False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A'", Cmb_Adm_Permisos_Consulta_Empleado, 0, 0, False, "TODOS")
        End If
    Else
        Cmb_Adm_Permisos_Consulta_Empleado.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Empresa_Click()
    If Chk_Adm_Permisos_Consulta_Empresa.Value = 1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Permisos_Consulta_Empresa, 1, "Nombre", "", False, "")
    Else
        Cmb_Adm_Permisos_Consulta_Empresa.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Supervisor_Click()
    If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 Then
        If Chk_Adm_Permisos_Consulta_Empresa.Value = 1 And Cmb_Adm_Permisos_Consulta_Empresa.ListIndex > 0 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Supervisor, 1, "Apellido_paterno", "AND Tipo = 'S' AND Estatus = 'A' AND Empresa_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Empresa.ItemData(Cmb_Adm_Permisos_Consulta_Empresa.ListIndex), "00000") & "'", False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Supervisor, 1, "Apellido_paterno", "AND Tipo = 'S' AND Estatus = 'A'", False, "TODOS")
        End If
    Else
        Cmb_Adm_Permisos_Consulta_Supervisor.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Tipo_Permiso_Click()
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Locked = False
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID, Descripcion", "Cat_Tipos_Faltas", Cmb_Adm_Permisos_Consulta_Tipo_Permiso, 0, "Descripcion")
    Else
        'Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Clear
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Clear
        Simbologia_Consulta = ""
        SubSimbologia_Consulta = ""
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Locked = True
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Departamento_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Departamentos.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Adm_Permisos_Consulta_Departamento, 1, "Nombre")
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Departamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Departamento, KeyCode)
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Empleado_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 Then
        If KeyAscii = 13 Then
            'If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
                If IsNumeric(Cmb_Adm_Permisos_Consulta_Empleado.Text) Then
                    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE No_Tarjeta='" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "'", Cmb_Adm_Permisos_Consulta_Empleado, 0, "No_Tarjeta")
                Else
                    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", " OR Nombre LIKE '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%'" & _
                         " OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%'", False, "")
                End If
            'Else
            '    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "No_Tarjeta", " AND (Nombre like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR " & _
                     "Apellido_Paterno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%') ", False, "")
            'End If
        Else
            Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Empleado, KeyCode)
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Empresa_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Empresa.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Permisos_Consulta_Empresa, 1, "Nombre")
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Empresa, KeyCode)
End Sub


Private Sub Cmb_Adm_Permisos_Consulta_Supervisor_Click()
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", "AND Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", False, "")
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Supervisor_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S' ORDER BY Apellido_Paterno", Cmb_Adm_Permisos_Consulta_Supervisor, 0, "")
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Supervisor, 1, "Apellido_Paterno", "AND Estatus = 'A' AND Tipo='S' AND(Nombre like '%" & Trim(Cmb_Adm_Permisos_Consulta_Supervisor.Text) & "%' OR " & _
                "Apellido_Paterno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Supervisor.Text) & "%') ", False, "")
            If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", "AND (Nombre like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR " & _
                     "Apellido_Paterno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%') AND Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", False, "")
            End If
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Supervisor, KeyCode)
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Tipo_Permiso_Click()
    If Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex > -1 Then
        Simbologia_Consulta = ""
        SubSimbologia_Consulta = ""
        Simbologia = Conectar_Ayudante.Busca_Dato_BD("SELECT Tipo_Falta_ID, Simbologia FROM Cat_Tipos_Faltas WHERE Tipo_Falta_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ItemData(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex), "00000") & "'", "Simbologia")
    Else
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Clear
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Tipo_Permiso_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
        If KeyAscii = 13 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID, Descripcion", "Cat_Tipos_Faltas", Cmb_Adm_Permisos_Consulta_Tipo_Permiso, 1, "Descripcion")
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Tipo_Permiso_KeyUp(KeyCode As Integer, Shift As Integer)
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Tipo_Permiso, KeyCode)
    Else
        KeyCode = 0
    End If
End Sub

Private Sub Cmb_Permisos_Empleado_Click()
Dim Rs_Consuta_Cat_Empleados As rdoResultset     'Informcion de los empleados
    If Cmb_Permisos_Empleado.ListIndex > -1 Then
        Call Limpia_Informacion_Empleado
        'Llena los datos del departamento, Empresa y Supervisor
        Mi_SQL = "SELECT ISNULL(Supervisor_ID,'') as Supervisor_ID ,ISNULL(Empresa_ID,'') as Empresa_ID,ISNULL(Departamento_ID,'') as Departamento_ID FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID = '" & Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000") & "'"
        Set Rs_Consuta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consuta_Cat_Empleados
            If Not .EOF Then
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Permisos_Departamento)
                Txt_Permisos_Departamento.Text = Cmb_Permisos_Departamento.Text
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Empresa_ID"), Cmb_Permisos_Empresa)
                Txt_Permisos_Empresa.Text = Cmb_Permisos_Empresa.Text
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Permisos_Supervisor)
                Txt_Permisos_Supervisor.Text = Cmb_Permisos_Supervisor.Text
            End If
        End With
        
    End If
End Sub

Private Sub Cmb_Permisos_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Limpia_Informacion_Empleado
        If IsNumeric(Cmb_Permisos_Empleado.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND No_Tarjeta='" & Trim(Cmb_Permisos_Empleado.Text) & "'", Cmb_Permisos_Empleado, 0, "Apellido_Paterno")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Permisos_Empleado.Text) & "%' OR " & "Apellido_Paterno like '%" & Trim(Cmb_Permisos_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Permisos_Empleado.Text) & "%')", Cmb_Permisos_Empleado, 0, "Apellido_Paterno")
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Permisos_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Permisos_Empleado, KeyCode)
End Sub

Private Sub Cmb_Permisos_Incidencias_Extraordinarias_Click()
    If Cmb_Permisos_Incidencias_Extraordinarias.ListIndex > -1 Then
'        Txt_Permiso_Horas.Locked = True
'        Txt_Permiso_Horas.Text = ""
        Simbologia = Conectar_Ayudante.Busca_Dato_BD("SELECT Tipo_Falta_ID, Simbologia FROM Cat_Tipos_Faltas WHERE Tipo_Falta_ID = '" & Format(Cmb_Permisos_Incidencias_Extraordinarias.ItemData(Cmb_Permisos_Incidencias_Extraordinarias.ListIndex), "00000") & "'", "Simbologia")
'        If Simbologia = "AS" Then
            Txt_Permiso_Horas.Locked = False
'            If Fra_Permisos_Tipo.Enabled = True Then Txt_Permiso_Horas.SetFocus
'        End If
    End If
End Sub

Private Sub Cmb_Permisos_Incidencias_Extraordinarias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If Usuario_ID = "00001" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID,Descripcion", "Cat_Tipos_Faltas", Cmb_Permisos_Incidencias_Extraordinarias, 1, "Descripcion")
        'Else
        '    Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID,Descripcion", "Cat_Tipos_Faltas", Cmb_Permisos_Incidencias_Extraordinarias, 1, "Descripcion", " AND (Clasificacion='TODOS' OR Clasificacion IS NULL)")
        'End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Permisos_Incidencias_Extraordinarias_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Permisos_Incidencias_Extraordinarias, KeyCode)
End Sub

Private Sub Dtp_Permiso_Fecha_Inicio_Change()
    Dtp_Permiso_Fecha_Inicio_Click
End Sub

Private Sub Dtp_Permiso_Fecha_Inicio_Click()
    Txt_Permiso_Dias_Sueldo.Text = DateDiff("d", Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")) + 1
End Sub

Private Sub Dtp_Permiso_Fecha_Inicio_LostFocus()
    Dtp_Permiso_Fecha_Inicio_Click
End Sub

Private Sub Dtp_Permiso_Fecha_Termino_Change()
    Dtp_Permiso_Fecha_Inicio_Click
End Sub

Private Sub Dtp_Permiso_Fecha_Termino_Click()
    Dtp_Permiso_Fecha_Inicio_Click
End Sub

Private Sub Dtp_Permiso_Fecha_Termino_LostFocus()
    Dtp_Permiso_Fecha_Inicio_Click
End Sub

Private Sub Grid_Adm_Permisos_Consulta_Resultados_DblClick()
    If Grid_Adm_Permisos_Consulta_Resultados.Rows > 0 Then
        Llenar_Informacion_Permiso (Grid_Adm_Permisos_Consulta_Resultados.TextMatrix(Grid_Adm_Permisos_Consulta_Resultados.RowSel, 0))
    End If
End Sub

Private Sub Txt_Permiso_Horas_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Permiso_Horas.Text, True)
End Sub

Private Sub Txt_Permisos_Observaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

'***************************************Inicio Movimientos de permiso********************************
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Permisos
'DESCRIPCION: Da de alta un registro de incidencia del empleado seleccionado
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 15-Mayo-2009
'MODIFICO   : Sergio Ulises Durán Hernández
'FECHA_MODIFICO: 30-Enero-2014
'CAUSA_MODIFICO: Se valida para no repetir una incidencia en la misma fecha
'******************************************************************************
Private Sub Alta_Permisos()
Dim Rs_Alta_Adm_Movimiento As rdoResultset 'Informacion del Maquinas
Dim Rs_Consulta_Adm_Movimiento As rdoResultset 'Informacion del Maquinas
Dim No_Movimiento As String
Dim Motivo As String

On Error GoTo HANDLER
    'Valida si no existen ya un permiso para esas fechas
    Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND Fecha_Inicio>='" & Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Fecha_Termino<='" & Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Set Rs_Consulta_Adm_Movimiento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Movimiento.EOF Then
        'If MsgBox("Ya existe un registro de incidencia en las fechas señaladas." & Chr(13) & "¿Desea darlo de alta de todos modos?", vbQuestion + vbYesNo) = vbNo Then
            MsgBox "Ya existe un registro de incidencia en las fechas señaladas.", vbExclamation
            Rs_Consulta_Adm_Movimiento.Close
            Exit Sub
        'End If
    End If
    Rs_Consulta_Adm_Movimiento.Close
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Motivo = ""
    Set Rs_Alta_Adm_Movimiento = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos_Asistencias")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Adm_Movimiento
        .AddNew
            'Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias WHERE Tipo_Incidencia = 'E'", "No_Movimiento"), "0000000000")
            Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias", "No_Movimiento"), "0000000000")
            No_Movimiento = Trim(Txt_Adm_Permisos_No_Movimiento.Text)
            .rdoColumns("No_Movimiento") = Trim(No_Movimiento)
            .rdoColumns("Empresa_ID") = Format(Cmb_Permisos_Empresa.ItemData(Cmb_Permisos_Empresa.ListIndex), "00000")
            .rdoColumns("Empleado_ID") = Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000")
            .rdoColumns("Departamento_ID") = Format(Cmb_Permisos_Departamento.ItemData(Cmb_Permisos_Departamento.ListIndex), "00000")
            .rdoColumns("Tipo_Falta_ID") = Format(Cmb_Permisos_Incidencias_Extraordinarias.ItemData(Cmb_Permisos_Incidencias_Extraordinarias.ListIndex), "00000")
            .rdoColumns("Tipo_Incidencia") = "E"
            .rdoColumns("Fecha_Solicitud") = Format(Dtp_Permiso_Fecha_Solicitud.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Inicio") = Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Termino") = Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")
            .rdoColumns("Dias_Permiso") = Val(Txt_Permiso_Dias_Sueldo.Text)
            .rdoColumns("Periodo") = 0
            .rdoColumns("Horas_Acuerdo") = Val(Txt_Permiso_Horas.Text)
            .rdoColumns("Hora_Regreso") = Format("00:00:00", "HH:mm")
            .rdoColumns("Motivo") = Trim(Cmb_Permisos_Incidencias_Extraordinarias.Text)
            If Trim(Cmb_Permisos_Incidencias_Extraordinarias.Text) = "VACACIONES" Or Trim(Cmb_Permisos_Incidencias_Extraordinarias.Text) = "PERMISO A CUENTA DE VACACIONES" Then
                'Descuenta de vacaciones
                Mi_SQL = "UPDATE Cat_Empleados SET Salario_Diario_Variable=Salario_Diario_Variable-" & Val(Txt_Permiso_Dias_Sueldo.Text)
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
                Call Alta_Vacaciones
                Conexion_Base.Execute Mi_SQL
            End If
            .rdoColumns("Observaciones") = Trim(Txt_Permisos_Observaciones.Text)
            .rdoColumns("Simbologia") = Simbologia
            .rdoColumns("Subsimbologia") = SubSimbologia
            .rdoColumns("Estatus") = "A"
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Adm_Movimiento = Nothing
    'Habilita y deshabilita los controles de la forma
    Pic_Solicitud_Permisos.Enabled = False
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Imprimir.Enabled = True
    Dtp_Permiso_Fecha_Solicitud.Value = Now
    Dtp_Permiso_Fecha_Inicio.Value = Now
    Dtp_Permiso_Fecha_Termino.Value = Now
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Incidencias_Extraordinarias", Me)
    MsgBox "Permiso Registrado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Permisos
'DESCRIPCION: Actualiza el registro de la incidencia, si viene desde permiso lo transforma
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 15-Mayo-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Permisos()
Dim Rs_Modifica_Adm_Movimiento As rdoResultset 'Informacion del Maquinas
Dim Motivo As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
    Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & "'"
    'Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'E'"
    Motivo = ""
    Set Rs_Modifica_Adm_Movimiento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Adm_Movimiento.EOF Then
        With Rs_Modifica_Adm_Movimiento
            .Edit
                .rdoColumns("Empresa_ID") = Format(Cmb_Permisos_Empresa.ItemData(Cmb_Permisos_Empresa.ListIndex), "00000")
                .rdoColumns("Empleado_ID") = Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000")
                .rdoColumns("Departamento_ID") = Format(Cmb_Permisos_Departamento.ItemData(Cmb_Permisos_Departamento.ListIndex), "00000")
                .rdoColumns("Tipo_Falta_ID") = Format(Cmb_Permisos_Incidencias_Extraordinarias.ItemData(Cmb_Permisos_Incidencias_Extraordinarias.ListIndex), "00000")
                .rdoColumns("Tipo_Incidencia") = "E"
                .rdoColumns("Fecha_Solicitud") = Format(Dtp_Permiso_Fecha_Solicitud.Value, "MM/dd/yyyy")
                .rdoColumns("Fecha_Inicio") = Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy")
                .rdoColumns("Fecha_Termino") = Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")
                .rdoColumns("Dias_Permiso") = Val(Txt_Permiso_Dias_Sueldo.Text)
                .rdoColumns("Horas_Acuerdo") = Val(Txt_Permiso_Horas.Text)
                .rdoColumns("Periodo") = 0
                .rdoColumns("Motivo") = Trim(Cmb_Permisos_Incidencias_Extraordinarias.Text)
                .rdoColumns("Observaciones") = Trim(Txt_Permisos_Observaciones.Text)
                .rdoColumns("Simbologia") = Simbologia
                .rdoColumns("Subsimbologia") = SubSimbologia
                '.rdoColumns("Estatus") = Mid(Cmb_Permisos_Estatus.Text, 1, 1)
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
    End If
    Rs_Modifica_Adm_Movimiento.Close
    Set Rs_Modifica_Adm_Movimiento = Nothing
    'Habilita y deshabilita los controles de la forma
    Pic_Solicitud_Permisos.Enabled = False
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Imprimir.Enabled = True
    Cmb_Permisos_Estatus.Enabled = False
    Dtp_Permiso_Fecha_Solicitud.Value = Now
    Dtp_Permiso_Fecha_Inicio.Value = Now
    Dtp_Permiso_Fecha_Termino.Value = Now
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Incidencias_Extraordinarias", Me)
    MsgBox "La incidencia ha sido modificada", vbInformation + vbOKOnly, Me.Caption
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'************************************************Termino Vacaciones***************************************
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Permisos
'DESCRIPCION: Consulta los permisos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 28-Junio-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Permisos()
Dim Rs_Consulta_Adm_Movimientos As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios

On Error GoTo HANDLER:
    Grid_Adm_Permisos_Consulta_Resultados.Rows = 0
    Grid_Adm_Permisos_Consulta_Resultados.Cols = 11
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT AM.No_Movimiento,AM.Empleado_ID,AM.Departamento_ID,AM.Tipo_Falta_ID,AM.Tipo_Incidencia"
    Mi_SQL = Mi_SQL & " ,AM.Fecha_Inicio,AM.Fecha_Termino,AM.Estatus, AM.Simbologia, AM.SubSimbologia,"
    Mi_SQL = Mi_SQL & " AM.Motivo, AM.Observaciones,"
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre,CE.No_Tarjeta"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias AM, Cat_Empleados CE, Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AM.Empleado_Id = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND AM.Departamento_ID = CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND AM.Tipo_Falta_ID IS NOT NULL"
    'Mi_SQL = Mi_SQL & " AND AM.Tipo_Incidencia = 'E'"
    'Manejo de Filtros
    'Empresas
    If Chk_Adm_Permisos_Consulta_Empresa.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Empresa.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AM.Empresa_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Empresa.ItemData(Cmb_Adm_Permisos_Consulta_Empresa.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ninguna empresa", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Departamentos
    If Chk_Adm_Permisos_Consulta_Departamentos.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Departamento.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AM.Departamento_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Departamento.ItemData(Cmb_Adm_Permisos_Consulta_Departamento.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun departamento", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Supervisor
    If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun supervisor", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Empleados
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Empleado.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AM.Empleado_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Empleado.ItemData(Cmb_Adm_Permisos_Consulta_Empleado.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado ningun empleado", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Tipo
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AM.Tipo_Falta_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ItemData(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado la incidencia extraordinaria", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'Periodo
    If Chk_Adm_Permisos_Consulta_Periodo.Value = 1 Then
        If DateDiff("d", Format(Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Dtp_Adm_Permisos_Consulta_Fecha_Termino, "MM/dd/yyyy")) < 0 Then
            MsgBox "Rango de Fechas Incorrecto", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        Else
            Mi_SQL = Mi_SQL & " AND ((AM.Fecha_Inicio BETWEEN " & Par_Fecha & Format(Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND " & Par_Fecha & Format(Dtp_Adm_Permisos_Consulta_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha & ")"
            Mi_SQL = Mi_SQL & " OR (AM.Fecha_Termino BETWEEN " & Par_Fecha & Format(Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND " & Par_Fecha & Format(Dtp_Adm_Permisos_Consulta_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha & "))"
        End If
    End If
    Mi_SQL = Mi_SQL & " ORDER BY AM.Fecha_Inicio"
    Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    With Rs_Consulta_Adm_Movimientos
    If Not .EOF Then
        Me.MousePointer = 11
        'Agrega el encabezado al reporte
        If Chk_Adm_Permisos_Consulta_Periodo.Value = 1 Then
            Call Encabezado_Reporte("REPORTE INCIDENCIAS EXTRAORDINARIAS", DateAdd("s", 1, Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Adm_Permisos_Consulta_Fecha_Termino.Value))
        Else
            Call Encabezado_Reporte("REPORTE INCIDENCIAS EXTRAORDINARIAS")
        End If
        Print #2, "No_Movimiento|No.Empleado|Nombre|F. Inicio|F. Termino|Tipo|SubTipo|Estatus|Motivo_Horas|Observaciones"
        'Coloca un encabezado en el grid
        Grid_Adm_Permisos_Consulta_Resultados.AddItem "No_Movimiento" & Chr(9) & "No." & Chr(9) & "Nombre" & Chr(9) & "F. Inicio" & Chr(9) & "F. Termino" _
            & Chr(9) & "Tipo" & Chr(9) & "SubTipo" & Chr(9) & "Estatus" & Chr(9) & "Motivo_Horas" & Chr(9) & "Observaciones"
        While Not .EOF
            Grid_Adm_Permisos_Consulta_Resultados.AddItem .rdoColumns("No_Movimiento") _
                & Chr(9) & .rdoColumns("No_Tarjeta") _
                & Chr(9) & .rdoColumns("Nombre") _
                & Chr(9) & .rdoColumns("Fecha_Inicio") _
                & Chr(9) & .rdoColumns("Fecha_Termino") _
                & Chr(9) & .rdoColumns("Tipo_Incidencia") _
                & Chr(9) & .rdoColumns("Simbologia") _
                & Chr(9) & .rdoColumns("Estatus") _
                & Chr(9) & .rdoColumns("Tipo_Falta_ID") _
                & Chr(9) & .rdoColumns("Motivo") _
                & Chr(9) & .rdoColumns("Observaciones")
            Print #2, .rdoColumns("No_Movimiento") _
                & "|" & .rdoColumns("No_Tarjeta") _
                & "|" & .rdoColumns("Nombre") _
                & "|" & .rdoColumns("Fecha_Inicio") _
                & "|" & .rdoColumns("Fecha_Termino") _
                & "|" & .rdoColumns("Tipo_Incidencia") _
                & "|" & .rdoColumns("Simbologia") _
                & "|" & .rdoColumns("Estatus") _
                & "|" & .rdoColumns("Tipo_Falta_ID") _
                & "|" & .rdoColumns("Motivo") _
                & "|" & .rdoColumns("Observaciones")
            .MoveNext
        Wend
        .Close
        Finalizar_Reporte (False)
        'Configura el tamaño de las columnas del grid_usuarios
        With Grid_Adm_Permisos_Consulta_Resultados
            .FixedRows = 1
            .ColWidth(0) = 0    'No Movimiento
            .ColWidth(1) = 700  'No_Tarjeta
            .ColWidth(2) = 1900 'Empleado
            .ColWidth(3) = 1000 'Fecha_Inicio
            .ColWidth(4) = 1000 'Fecha_Termino
            .ColWidth(5) = 500  'Tipo
            .ColWidth(6) = 800  'Simbologia
            .ColWidth(7) = 650  'Estatus
            .ColWidth(8) = 0    'Tipo Falta
            .ColWidth(9) = 0    'Motivo_Horas
            .ColWidth(10) = 0   'Observaciones
        End With
    End If
    End With
    Set Rs_Consulta_Adm_Movimientos = Nothing
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    Finalizar_Reporte (False)
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Llenar_Informacion_Permiso
'DESCRIPCION: Llena los datos de la incidencia en pantalla
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 09-Noviembre-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Llenar_Informacion_Permiso(No_Movimiento As String)
Dim Rs_Consulta_Adm_Movimientos As rdoResultset     'Informacion del permiso

    Mi_SQL = "SELECT AM.*, ISNULL(CE.Supervisor_ID,'') as Supervisor_ID, Tipo_Falta_ID "
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias AM, Cat_Empleados CE"
    Mi_SQL = Mi_SQL & " WHERE AM.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND No_Movimiento = '" & No_Movimiento & "'"
    'Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'E'"
    Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Movimientos.EOF Then
        With Rs_Consulta_Adm_Movimientos
            Txt_Adm_Permisos_No_Movimiento.Text = Trim(.rdoColumns("No_Movimiento"))
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Permisos_Departamento)
            Txt_Permisos_Departamento.Text = Cmb_Permisos_Departamento.Text
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Empresa_ID"), Cmb_Permisos_Empresa)
            Txt_Permisos_Empresa.Text = Cmb_Permisos_Empresa.Text
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Permisos_Supervisor)
            Txt_Permisos_Supervisor.Text = Cmb_Permisos_Supervisor.Text
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Empleado_ID = '" & Trim(.rdoColumns("Empleado_ID")) & "'", Cmb_Permisos_Empleado, 0, "Apellido_Paterno")
            Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Empleado_ID"), Cmb_Permisos_Empleado)
            Dtp_Permiso_Fecha_Solicitud.Value = .rdoColumns("Fecha_Solicitud")
            Dtp_Permiso_Fecha_Inicio.Value = .rdoColumns("Fecha_Inicio")
            Dtp_Permiso_Fecha_Termino.Value = .rdoColumns("Fecha_Termino")
            Txt_Permiso_Dias_Sueldo.Text = Val(.rdoColumns("Dias_Permiso"))
            '.rdoColumns("Periodo") = 0
            Txt_Permisos_Observaciones.Text = .rdoColumns("Observaciones")
            Cmb_Permisos_Incidencias_Extraordinarias.Text = .rdoColumns("Tipo_Falta_ID")
            Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID, Descripcion", "Cat_Tipos_Faltas", Cmb_Permisos_Incidencias_Extraordinarias, 1, "Tipo_Falta_ID")
            If Cmb_Permisos_Incidencias_Extraordinarias.ListCount > 0 Then
                Cmb_Permisos_Incidencias_Extraordinarias.ListIndex = 0
            End If
            'Valida si viene desde permiso
            If .rdoColumns("Tipo_Incidencia") = "E" Then
                If .rdoColumns("Estatus") = "A" Then
                    Cmb_Permisos_Estatus.ListIndex = 0
                Else
                    Cmb_Permisos_Estatus.ListIndex = 1
                End If
            Else
                Cmb_Permisos_Estatus.ListIndex = 2
            End If
            If Not IsNull(.rdoColumns("Horas_Acuerdo")) Then Txt_Permiso_Horas.Text = Val(.rdoColumns("Horas_Acuerdo"))
            Pic_Solicitud_Permisos.ZOrder vbBringToFront
            Pic_Solicitud_Permisos_Consulta.Visible = False
        End With
    End If
End Sub

'***************************************Termino Movimientos de permiso********************************
Private Sub Limpia_Informacion_Empleado()
    Cmb_Permisos_Departamento.ListIndex = -1
    Txt_Permisos_Departamento.Text = ""
    Cmb_Permisos_Empresa.ListIndex = -1
    Txt_Permisos_Empresa.Text = ""
    Cmb_Permisos_Supervisor.ListIndex = -1
    Txt_Permisos_Supervisor.Text = ""
End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Open Ruta_Temporal & "Incidencias_Extraordinarias" & ".txt" For Output As #1
    Open Ruta_Temporal & "Incidencias_Extraordinarias" & "xls.txt" For Output As #2 'Reporte a xls
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

Private Sub Finalizar_Reporte(Abrir As Boolean)
    Close #1, #2
    Archivo_Reporte_Abierto = False
    If Abrir Then
        
    End If
End Sub

'NOMBRE_FUNCION: Alta_Vacaciones
'DESCRIPCION: Da de alta un registro de incidencia del empleado cuando se toma como vacaciones
'PARAMETROS :
'CREO       : Ana Laura Huichapa Ramírez
'FECHA_CREO : 14-Marzo-2016
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Vacaciones()
Dim Rs_Alta_Adm_Vacaciones As rdoResultset
On Error GoTo HANDLER
    Set Rs_Alta_Adm_Vacaciones = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos_Vacaciones")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Adm_Vacaciones
        .AddNew
            Dim No_Movimiento_Vacaciones As String
            No_Movimiento_Vacaciones = Trim(Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Vacaciones", "No_Movimiento_Vacaciones"), "0000000000"))
            .rdoColumns("No_Movimiento_Vacaciones") = Trim(No_Movimiento_Vacaciones)
            .rdoColumns("Empleado_ID") = Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000")
            .rdoColumns("Fecha_Inicio") = Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Fin") = Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")
            .rdoColumns("Dias") = Val(Txt_Permiso_Dias_Sueldo.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Adm_Vacaciones = Nothing
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
