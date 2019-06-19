VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Adm_Solicitud_Permisos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9150
   ClientLeft      =   9960
   ClientTop       =   3465
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Solicitud_Permisos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8385
      Left            =   0
      ScaleHeight     =   8385
      ScaleWidth      =   7125
      TabIndex        =   33
      Top             =   0
      Width           =   7125
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   6600
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Crystal.CrystalReport Rpt_Reporte 
         Left            =   6165
         Top             =   345
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
         TabIndex        =   45
         Top             =   7365
         Width           =   6990
         Begin VB.TextBox Txt_Permisos_Observaciones 
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
            TabIndex        =   12
            Top             =   225
            Width           =   6795
         End
      End
      Begin VB.Frame Fra_Permisos_Tipo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Permiso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   90
         TabIndex        =   44
         Top             =   3285
         Width           =   7035
         Begin VB.ListBox List_Tipo_Permiso 
            Height          =   3660
            ItemData        =   "Frm_Adm_Solicitud_Permisos.frx":0000
            Left            =   90
            List            =   "Frm_Adm_Solicitud_Permisos.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   63
            Top             =   270
            Width           =   6825
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
         Height          =   2085
         Left            =   90
         TabIndex        =   36
         Top             =   1215
         Width           =   7035
         Begin VB.TextBox Txt_Permisos_Departamento 
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
            Left            =   4725
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   270
            Width           =   2220
         End
         Begin VB.TextBox Txt_Permiso_Dias_Sueldo 
            Alignment       =   1  'Right Justify
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
            Left            =   4710
            MaxLength       =   2
            TabIndex        =   11
            Top             =   1680
            Width           =   1875
         End
         Begin VB.ComboBox Cmb_Permisos_Empleado 
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
            Left            =   1200
            TabIndex        =   8
            Top             =   974
            Width           =   5760
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Inicio 
            Height          =   315
            Left            =   2115
            TabIndex        =   9
            Top             =   1320
            Width           =   1410
            _ExtentX        =   2487
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
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Termino 
            Height          =   315
            Left            =   4710
            TabIndex        =   10
            Top             =   1320
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   124059651
            CurrentDate     =   39940
         End
         Begin VB.TextBox Txt_Permisos_Empresa 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   270
            Width           =   2265
         End
         Begin VB.ComboBox Cmb_Permisos_Empresa 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   270
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox Txt_Permisos_Supervisor 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   622
            Width           =   5745
         End
         Begin VB.ComboBox Cmb_Permisos_Supervisor 
            Height          =   315
            Left            =   1530
            TabIndex        =   14
            Top             =   615
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.ComboBox Cmb_Permisos_Departamento 
            Height          =   315
            Left            =   4755
            Style           =   2  'Dropdown List
            TabIndex        =   57
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
            Left            =   3525
            TabIndex        =   58
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
            TabIndex        =   46
            Top             =   682
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas que solicita"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   135
            TabIndex        =   43
            Top             =   1290
            Width           =   945
            WordWrap        =   -1  'True
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
            Left            =   3600
            TabIndex        =   42
            Top             =   1740
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
            Left            =   3600
            TabIndex        =   41
            Top             =   1380
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
            Left            =   1275
            TabIndex        =   40
            Top             =   1380
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
            TabIndex        =   39
            Top             =   1740
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
            TabIndex        =   38
            Top             =   1034
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
            TabIndex        =   37
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.Frame Fra_Permiso_Datos_Generales 
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   90
         TabIndex        =   35
         Top             =   270
         Width           =   7035
         Begin VB.ComboBox Cmb_Permisos_Estatus 
            Enabled         =   0   'False
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
            ItemData        =   "Frm_Adm_Solicitud_Permisos.frx":0004
            Left            =   1395
            List            =   "Frm_Adm_Solicitud_Permisos.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   540
            Width           =   2070
         End
         Begin VB.TextBox Txt_Adm_Permisos_No_Movimiento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
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
            TabIndex        =   51
            Top             =   180
            Width           =   2070
         End
         Begin MSComCtl2.DTPicker Dtp_Permiso_Fecha_Solicitud 
            Height          =   315
            Left            =   4770
            TabIndex        =   60
            Top             =   540
            Width           =   2220
            _ExtentX        =   3916
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
            CustomFormat    =   "ddd dd MMMM yyyy"
            Format          =   124059651
            CurrentDate     =   40030
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
            TabIndex        =   62
            Top             =   600
            Width           =   450
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
            TabIndex        =   61
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
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUD DE PERMISO"
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
         Left            =   2295
         TabIndex        =   34
         Top             =   0
         Width           =   2805
      End
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   645
      Left            =   2493
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":0025
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   645
      Left            =   45
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":05AF
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "A"
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Cancelar"
      Height          =   645
      Left            =   3717
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":0B39
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "B"
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   645
      Left            =   1269
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":10C3
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "M"
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   645
      Left            =   6165
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":164D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   645
      Left            =   4941
      Picture         =   "Frm_Adm_Solicitud_Permisos.frx":1BD7
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "C"
      Top             =   8400
      Width           =   930
   End
   Begin VB.PictureBox Pic_Solicitud_Permisos_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8850
      Left            =   0
      ScaleHeight     =   8850
      ScaleWidth      =   7125
      TabIndex        =   47
      Top             =   315
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
         Height          =   2895
         Left            =   45
         TabIndex        =   49
         Top             =   75
         Width           =   7080
         Begin VB.CommandButton Btn_Exportar 
            Caption         =   "Exportar"
            Height          =   510
            Left            =   5880
            Picture         =   "Frm_Adm_Solicitud_Permisos.frx":2161
            Style           =   1  'Graphical
            TabIndex        =   53
            Tag             =   "A"
            Top             =   1275
            UseMaskColor    =   -1  'True
            Width           =   1110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso 
            Height          =   315
            ItemData        =   "Frm_Adm_Solicitud_Permisos.frx":26EB
            Left            =   1665
            List            =   "Frm_Adm_Solicitud_Permisos.frx":2710
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2135
            Width           =   4110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Empleado 
            Height          =   315
            Left            =   1665
            TabIndex        =   22
            Top             =   1371
            Width           =   4110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Empresa 
            Height          =   315
            Left            =   1665
            TabIndex        =   16
            Top             =   225
            Width           =   4110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Supervisor 
            Height          =   315
            Left            =   1665
            TabIndex        =   20
            Top             =   989
            Width           =   4110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Departamento 
            Height          =   315
            Left            =   1665
            TabIndex        =   18
            Top             =   607
            Width           =   4110
         End
         Begin VB.ComboBox Cmb_Adm_Permisos_Consulta_Tipo_Permiso 
            Height          =   315
            ItemData        =   "Frm_Adm_Solicitud_Permisos.frx":27D1
            Left            =   1665
            List            =   "Frm_Adm_Solicitud_Permisos.frx":27D3
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1753
            Width           =   4110
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Simbologia"
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
            Top             =   2135
            Width           =   1500
         End
         Begin VB.CheckBox Chk_Adm_Permisos_Consulta_Tipo_Permiso 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Permiso"
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
            Top             =   1753
            Width           =   1230
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
            TabIndex        =   17
            Top             =   607
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
            TabIndex        =   19
            Top             =   989
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Regresar 
            Cancel          =   -1  'True
            Caption         =   "Regresar"
            Height          =   510
            Left            =   5880
            Picture         =   "Frm_Adm_Solicitud_Permisos.frx":27D5
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2325
            UseMaskColor    =   -1  'True
            Width           =   1110
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
            TabIndex        =   27
            Top             =   2520
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
            TabIndex        =   21
            Top             =   1371
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
            TabIndex        =   15
            Top             =   225
            Width           =   1050
         End
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   510
            Left            =   5880
            Picture         =   "Frm_Adm_Solicitud_Permisos.frx":2D5F
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "C"
            Top             =   225
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Permisos_Consulta_Fecha_Termino 
            Height          =   315
            Left            =   4050
            TabIndex        =   29
            Top             =   2520
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
            Left            =   1665
            TabIndex        =   28
            Top             =   2520
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
            Left            =   5940
            TabIndex        =   54
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   5940
            TabIndex        =   55
            Top             =   1800
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   3652
            TabIndex        =   50
            Top             =   2580
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
         Height          =   5805
         Left            =   45
         TabIndex        =   48
         Top             =   3000
         Width           =   7080
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Permisos_Consulta_Resultados 
            Height          =   5430
            Left            =   90
            TabIndex        =   32
            Top             =   270
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   9578
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
   End
End
Attribute VB_Name = "Frm_Adm_Solicitud_Permisos"
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
Dim Usando_ListBox As Boolean
Dim Tipo_Falta_ID As String

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
                    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'P'"
                    Referencia = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Referencia")
                    If Referencia <> "" Then
                        MsgBox "El permiso esta siendo utilizado para el control de asistencias," & vbCrLf & _
                               "Si desea cancelar el permiso, " & vbCrLf & _
                               "debera primero modificar la asistencia a la que esta asociado", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                    Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
                    Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & "'"
                    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia='P'"
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
HANDLER:
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
    Cmd_Exportar.FileName = "Solicitud_Permisos" & ".xls"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
        Call Exportar_Excel(Ruta_Temporal & "Solicitud_Permisos" & "xls.txt", Ruta_Exportacion, Prbar_Exportacion, Lbl_Progreso_Exportacion, Me)
    End If
  ' Display name of selected file
  Exit Sub

HANDLER:
    Exit Sub

End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Solicitud_Permisos
'DESCRIPCION: Genera el reporte de la solicitud de permiso
'PARAMETROS :
'CREO       : Yazmin Flores Ramirez
'FECHA_CREO : 24-Noviembre-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Btn_Imprimir_Click()
'''---------- Crystal 11
''    Generar_Reporte_Solicitud_Permisos
'    Rpt_Reporte.Connect = Conexion_Base.Connect
'    Rpt_Reporte.ReportFileName = App.Path & "\Reportes\Rpt_Solicitud_Permiso.rpt"
'    Rpt_Reporte.ParameterFields(1) = "No_Movimiento;" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & ";true"
'    Rpt_Reporte.ParameterFields(2) = "Supervisor_ID;" & Trim(Txt_Permisos_Supervisor.Text) & ";true"
'    Rpt_Reporte.Destination = crptToWindow
'    Rpt_Reporte.PrintReport
'    Rpt_Reporte.WindowState = crptMaximized
'    Rpt_Reporte.PageZoom 100
Dim Nombre As String
Dim Nombre_RPT As String
Dim Hoora As Date
Hoora = Format$(Now, "d-mmmm-yy h:mm:ss")
Dim hora As String
hora = Replace(Hoora, " ", "")
hora = Replace(hora, ":", "_")
hora = Replace(hora, ".", "")
hora = Replace(hora, "/", "")
Nombre_RPT = "Rpt_Solicitud_Permiso"
Nombre = "Solicitud_Permiso_" & hora
        
Crea_PDF Nombre_RPT, Nombre

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
            Case "No_Movimiento"
                    parametro = Trim(Txt_Adm_Permisos_No_Movimiento.Text)
                    crParamDef.AddCurrentValue ("'" & parametro & "'")
            
            Case "Supervisor_ID"
                  parametro = Trim(Txt_Permisos_Supervisor.Text)
                crParamDef.AddCurrentValue (parametro)
'
            
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
Private Sub Generar_Reporte_Solicitud_Permisos()
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
    'Asigna el formato de la factura a la variable
    Set crxReport = crxApplication.OpenReport(App.Path & "\Reportes\Rpt_Solicitud_Permiso.Rpt")
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
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "No_Movimiento"
                crParamDef.AddCurrentValue (Trim(Txt_Adm_Permisos_No_Movimiento.Text))
            Case "Supervisor_ID"
                crParamDef.AddCurrentValue (Trim(Txt_Permisos_Supervisor.Text))
        End Select
    Next
    'Asigna el nombre del reporte
    Nombre_Reporte = App.Path & "\Faltas\P" & Format(Now, "yyyyMMddHHmm") & ".pdf"
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
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Modificar_Click()
Dim Tipo_Permiso As Boolean
Dim Cont_List_Permisos As Integer

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
                'Verifica si seleccionó un permiso
                For Cont_List_Permisos = 0 To List_Tipo_Permiso.ListCount - 1
                    If List_Tipo_Permiso.Selected(Cont_List_Permisos) = True Then
                        Tipo_Falta_ID = Mid(List_Tipo_Permiso.List(Cont_List_Permisos), 1, 5)
                        Tipo_Permiso = True
                        Exit For
                    End If
                Next
                'valida que la informacion este completa
                If Cmb_Permisos_Departamento.ListIndex > -1 Then    'Valida el departamento
                    If Cmb_Permisos_Empresa.ListIndex > -1 Then     'valida la empresa
                        If Cmb_Permisos_Supervisor.ListIndex > -1 Then  'Valida el supervisor
                            If Cmb_Permisos_Empleado.ListIndex > -1 Then    'Valida el empleado
                                If Val(Txt_Permiso_Dias_Sueldo.Text) > 0 Then  'valida los dias de permiso
                                    If Tipo_Permiso Then
                                        Modifica_Permisos
                                    Else
                                        MsgBox "No ha seleccionado ningun tipo de permiso", vbInformation + vbOKOnly, Me.Caption
                                    End If
                                Else
                                    MsgBox "No ha ingresado los dias para el permiso", vbInformation + vbOKOnly, Me.Caption
                                    Dtp_Permiso_Fecha_Termino.SetFocus
                                End If
                            Else
                                MsgBox "Debe proporcionar el empleado", vbInformation + vbOKOnly, Me.Caption
                                Cmb_Permisos_Empleado.SetFocus
                            End If
                        Else
                            MsgBox "Supervisor no asignado a empleado," & vbCrLf & _
                                "debera configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Empresa no asignada a empleado," & vbCrLf & _
                               "deberá configurarla en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                    End If
                Else
                    MsgBox "Departamento no asinado a empleado," & vbCrLf & _
                               "deberá configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                End If
        End Select
    End If
End Sub

'***************************************Acciones de Botones*********************************
Private Sub Btn_Nuevo_Click()
Dim Tipo_Permiso As Boolean     'Indica si se ha seleccionado algun tipo de permiso
Dim Cont_List_Permisos As Integer
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Imprimir.Enabled = False
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        Select Case Operacion
            Case "Permisos": 'Captura de Inasistencias
                'Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias WHERE Tipo_Incidencia = 'P'", "No_Movimiento"), "0000000000")
                Pic_Solicitud_Permisos.Enabled = True
                Cmb_Permisos_Estatus.ListIndex = 0
                Dtp_Permiso_Fecha_Inicio.Value = Now
                Dtp_Permiso_Fecha_Termino.Value = Now
                Dtp_Permiso_Fecha_Solicitud.Value = Now
                Dtp_Permiso_Fecha_Solicitud.SetFocus
        End Select
    Else
        Select Case Operacion
            Case "Permisos": 'Captura de Vacaciones
                Tipo_Permiso = False
                'Verifica si seleccionó un permiso
                For Cont_List_Permisos = 0 To List_Tipo_Permiso.ListCount - 1
                    If List_Tipo_Permiso.Selected(Cont_List_Permisos) = True Then
                        Tipo_Falta_ID = Mid(List_Tipo_Permiso.List(Cont_List_Permisos), 1, 5)
                        Tipo_Permiso = True
                        Exit For
                    End If
                Next
                'valida que la informacion este completa
                If Cmb_Permisos_Departamento.ListIndex > -1 Then    'Valida el departamento
                    If Cmb_Permisos_Empresa.ListIndex > -1 Then     'valida la empresa
                        If Cmb_Permisos_Empleado.ListIndex > -1 Then    'Valida el empleado
                            If Val(Txt_Permiso_Dias_Sueldo.Text) > 0 Or Txt_Permiso_Dias_Sueldo.Enabled = False Then 'valida los dias de permiso
                                If Tipo_Permiso Then
                                    Alta_Permisos
                                Else
                                    MsgBox "No ha seleccionado ningun tipo de permiso", vbInformation + vbOKOnly, Me.Caption
                                End If
                            Else
                                MsgBox "No ha ingresado los dias para el permiso", vbInformation + vbOKOnly, Me.Caption
                                Dtp_Permiso_Fecha_Termino.SetFocus
                            End If
                        Else
                            MsgBox "Debe proporcionar el empleado", vbInformation + vbOKOnly, Me.Caption
                            Cmb_Permisos_Empleado.SetFocus
                        End If
                    Else
                        MsgBox "Empresa no asignada a empleado," & vbCrLf & _
                               "deberá configurarla en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
                    End If
                Else
                    MsgBox "Departamento no asinado a empleado," & vbCrLf & _
                               "deberá configurarlo en el catalogo de empleados", vbInformation + vbOKOnly, Me.Caption
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
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Sub SubMenu_Adm_Solicitud_Permisos", Me)
    
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
    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S' ", Cmb_Permisos_Supervisor, 0, "Apellido_Paterno")
    Call Llena_List_View
End Sub

Public Sub Llena_List_View()
Dim Rs_Consulta_Tipos_Faltas As rdoResultset

    'Limpia el List
    List_Tipo_Permiso.Clear
    'Consulta los tipos
    Mi_SQL = "SELECT Tipo_Falta_ID,Descripcion FROM Cat_Tipos_Faltas"
    Mi_SQL = Mi_SQL & " WHERE Tipo_Falta_ID>'00002'"
    Mi_SQL = Mi_SQL & " ORDER BY Tipo_Falta_ID"
    Set Rs_Consulta_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Tipos_Faltas.EOF
        List_Tipo_Permiso.AddItem Rs_Consulta_Tipos_Faltas.rdoColumns("Tipo_Falta_ID") & " - " & Rs_Consulta_Tipos_Faltas.rdoColumns("Descripcion")
        Rs_Consulta_Tipos_Faltas.MoveNext
    Wend
    Rs_Consulta_Tipos_Faltas.Close
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
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_paterno", " AND Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_paterno", " AND Estatus = 'A'", False, "TODOS")
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
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Permisos_Consulta_Empresa, 1, "Nombre")
    Else
        Cmb_Adm_Permisos_Consulta_Empresa.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso_Click()
    If Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 1 Then
        If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
            Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Locked = False
        Else
            Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 0
            MsgBox "Debe seleccionar un Tipo de Permiso", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        SubSimbologia_Consulta = ""
        Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Clear
        Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Locked = True
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Supervisor_Click()
    If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 Then
        If Chk_Adm_Permisos_Consulta_Empresa.Value = 1 And Cmb_Adm_Permisos_Consulta_Empresa.ListIndex > 0 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Adm_Permisos_Consulta_Supervisor, 1, "Apellido_paterno", "AND Tipo = 'S' AND Estatus = 'A' AND Empresa_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Empresa.ItemData(Cmb_Adm_Permisos_Consulta_Empresa.ListIndex), "00000") & "'", False, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Adm_Permisos_Consulta_Supervisor, 1, "Apellido_paterno", "AND Tipo = 'S' AND Estatus = 'A'", False, "TODOS")
        End If
    Else
        Cmb_Adm_Permisos_Consulta_Supervisor.Clear
    End If
End Sub

Private Sub Chk_Adm_Permisos_Consulta_Tipo_Permiso_Click()
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Locked = False
    Else
        'Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Clear
        Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex = -1
        Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 0
        Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Clear
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


Private Sub Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso_Click()
    If Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.ListCount > 0 Then
        SubSimbologia_Consulta = ""
        If Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.ListIndex > -1 Then
            Select Case Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Text
                    Case "ENFERMEDAD GENERAL"
                        SubSimbologia_Consulta = "EG"
                    Case "MATERNIDAD"
                        SubSimbologia_Consulta = "MA"
                    Case "RIESGO DE TRABAJO"
                        SubSimbologia_Consulta = "RT"
                    Case "VACACIONES"
                        SubSimbologia_Consulta = "VA"
                    Case "ALUMBRAMIENTO DE CONYUGE"
                        SubSimbologia_Consulta = "AL"
                    Case "DENFUNCION"
                        SubSimbologia_Consulta = "DE"
                    Case "MATRIMONIO"
                        SubSimbologia_Consulta = "MO"
            End Select
        End If
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 1 Then
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso_KeyUp(KeyCode As Integer, Shift As Integer)
    If Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 1 Then
        Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso, KeyCode)
    Else
        KeyCode = 0
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Supervisor_Click()
    If Chk_Adm_Permisos_Consulta_Empleado.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Adm_Permisos_Consulta_Supervisor.ItemData(Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex), "00000") & "'", Cmb_Adm_Permisos_Consulta_Empleado, 0, "")
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
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "Apellido_Paterno", "AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR " & _
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
        Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Clear
        Select Case Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Text
            Case "INCAPACIDAD"
                Simbologia_Consulta = ""
                SubSimbologia_Consulta = ""
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "ENFERMEDAD GENERAL"
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "MATERNIDAD"
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "RIESGO DE TRABAJO"
                Simbologia_Consulta = "II"
            Case "DERECHO"
                Simbologia_Consulta = ""
                SubSimbologia_Consulta = ""
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "VACACIONES"
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "ALUMBRAMIENTO DE CONYUGE"
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "DENFUNCION"
                Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.AddItem "MATRIMONIO"
                Simbologia_Consulta = "ID"
            Case "FALTA JUSTIFICADA"
                Simbologia_Consulta = ""
                SubSimbologia_Consulta = ""
                Simbologia_Consulta = "FJ"
            Case "PERMISO TEMPORAL"
                Simbologia_Consulta = ""
                SubSimbologia_Consulta = ""
                Simbologia_Consulta = "PE"
            Case "SANCION"
                Simbologia_Consulta = ""
                SubSimbologia_Consulta = ""
                Simbologia_Consulta = "SA"
        End Select
    End If
End Sub

Private Sub Cmb_Adm_Permisos_Consulta_Tipo_Permiso_KeyPress(KeyAscii As Integer)
    If Chk_Adm_Permisos_Consulta_Tipo_Permiso.Value = 1 Then
        If KeyAscii = 13 Then
            Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Falta_ID,Descripcion", "Cat_Tipos_Faltas", Cmb_Adm_Permisos_Consulta_Tipo_Permiso, 1, "Descripcion")
            If Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListCount > 0 Then Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex = 0 Else Cmb_Adm_Permisos_Consulta_Tipo_Permiso.Text = ""
        Else
            Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
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
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND No_Tarjeta='" & Cmb_Permisos_Empleado.Text & "'", Cmb_Permisos_Empleado, 0, "No_Tarjeta")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Permisos_Empleado.Text) & "%' OR Apellido_Paterno LIKE '%" & Trim(Cmb_Permisos_Empleado.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Permisos_Empleado.Text) & "%')", Cmb_Permisos_Empleado, 0, "Apellido_Paterno")
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Permisos_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Permisos_Empleado, KeyCode)
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

Private Sub List_Tipo_Permiso_Click()
Dim Item_Seleccionado As Integer
Dim Cont_List_Permisos As Integer

    If Btn_Nuevo.Caption = "Dar de Alta" Or Btn_Modificar.Caption = "Actualizar" Then
        If Usando_ListBox = False Then
            Usando_ListBox = True
            Item_Seleccionado = List_Tipo_Permiso.ListIndex
            'Verifica si seleccionó previo para dejar sólo uno seleccionado
            For Cont_List_Permisos = 0 To List_Tipo_Permiso.ListCount - 1
                If Item_Seleccionado <> Cont_List_Permisos Then
                    List_Tipo_Permiso.Selected(Cont_List_Permisos) = False
                End If
            Next
            Usando_ListBox = False
        End If
    End If
End Sub

Private Sub Txt_Permiso_Dias_Sueldo_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Permiso_Dias_Sueldo, False)
End Sub

Private Sub Txt_Permisos_Observaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

'***************************************Inicio Movimientos de permiso********************************
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Permisos
'DESCRIPCION: Se da de alta una solicitud de permiso
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 25-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Permisos()
Dim Rs_Alta_Adm_Movimiento As rdoResultset 'Informacion del Maquinas
Dim No_Movimiento As String
Dim Motivo As String
Dim Horas_Acuerdo As Double                 'Horas de acuerdo por el turno o tipo de permiso

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Motivo = ""
    Set Rs_Alta_Adm_Movimiento = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos_Asistencias")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Adm_Movimiento
        .AddNew
            'Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias WHERE Tipo_Incidencia = 'P'", "No_Movimiento"), "0000000000")
            Txt_Adm_Permisos_No_Movimiento.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos_Asistencias", "No_Movimiento"), "0000000000")
            No_Movimiento = Trim(Txt_Adm_Permisos_No_Movimiento.Text)
            .rdoColumns("No_Movimiento") = Trim(No_Movimiento)
            .rdoColumns("Empresa_ID") = Format(Cmb_Permisos_Empresa.ItemData(Cmb_Permisos_Empresa.ListIndex), "00000")
            .rdoColumns("Empleado_ID") = Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000")
            .rdoColumns("Departamento_ID") = Format(Cmb_Permisos_Departamento.ItemData(Cmb_Permisos_Departamento.ListIndex), "00000")
            .rdoColumns("Tipo_Falta_ID") = Tipo_Falta_ID
            .rdoColumns("Fecha_Solicitud") = Format(Dtp_Permiso_Fecha_Solicitud.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Inicio") = Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Termino") = Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")
            .rdoColumns("Dias_Permiso") = Val(Txt_Permiso_Dias_Sueldo.Text)
            .rdoColumns("Periodo") = 0
            .rdoColumns("Hora_Regreso") = Format(0, "HH:mm")
            '.rdoColumns("Horas_Acuerdo") = Horas_Acuerdo
            .rdoColumns("Motivo") = Trim(Motivo)
            .rdoColumns("Tipo_Incidencia") = "P"
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
    Conexion_Base.CommitTrans
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
    Tipo_Falta_ID = ""
    'Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos", Me)
    If MsgBox("La solicitud de permiso ha sido registrada" & Chr(13) & "¿Desea enviarlo a imprimir?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Btn_Imprimir_Click
    End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Permisos
'DESCRIPCION: Se realiza la modificacion de los permisos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 25-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Permisos()
Dim Rs_Modifica_Adm_Movimiento As rdoResultset 'Informacion del Maquinas
Dim Motivo As String
Dim Horas_Acuerdo As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Adm_Movimientos_Asistencias"
    Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(Txt_Adm_Permisos_No_Movimiento.Text) & "'"
    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'P'"
    Motivo = ""
    Set Rs_Modifica_Adm_Movimiento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Modifica_Adm_Movimiento
        .Edit
            .rdoColumns("Empresa_ID") = Format(Cmb_Permisos_Empresa.ItemData(Cmb_Permisos_Empresa.ListIndex), "00000")
            .rdoColumns("Empleado_ID") = Format(Cmb_Permisos_Empleado.ItemData(Cmb_Permisos_Empleado.ListIndex), "00000")
            .rdoColumns("Departamento_ID") = Format(Cmb_Permisos_Departamento.ItemData(Cmb_Permisos_Departamento.ListIndex), "00000")
            .rdoColumns("Tipo_Falta_ID") = Tipo_Falta_ID
            .rdoColumns("Fecha_Solicitud") = Format(Dtp_Permiso_Fecha_Solicitud.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Inicio") = Format(Dtp_Permiso_Fecha_Inicio.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Termino") = Format(Dtp_Permiso_Fecha_Termino.Value, "MM/dd/yyyy")
            .rdoColumns("Dias_Permiso") = Val(Txt_Permiso_Dias_Sueldo.Text)
            .rdoColumns("Periodo") = 0
            .rdoColumns("Hora_Regreso") = Format(0, "HH:mm")
            '.rdoColumns("Horas_Acuerdo") = Horas_Acuerdo
            .rdoColumns("Motivo") = Trim(Motivo)
            .rdoColumns("Observaciones") = Trim(Txt_Permisos_Observaciones.Text)
            .rdoColumns("Tipo_Incidencia") = "P"
            .rdoColumns("Simbologia") = Simbologia
            .rdoColumns("Subsimbologia") = SubSimbologia
            .rdoColumns("Estatus") = Mid(Cmb_Permisos_Estatus.Text, 1, 1)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Adm_Movimiento = Nothing
    Conexion_Base.CommitTrans
    'Habilita y deshabilita los controles de la forma
    Pic_Solicitud_Permisos.Enabled = False
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Imprimir.Enabled = True
    Cmb_Permisos_Estatus.Enabled = False
    Dtp_Permiso_Fecha_Solicitud.Value = Now
    Dtp_Permiso_Fecha_Inicio.Value = Now
    Dtp_Permiso_Fecha_Termino.Value = Now
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos", Me)
    MsgBox "La solicitud de permiso ha sido modificada", vbInformation + vbOKOnly, Me.Caption
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Permisos
'DESCRIPCION: Consulta los registros de permisos en el sistema
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 26-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Permisos()
Dim Rs_Consulta_Adm_Movimientos As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios

On Error GoTo HANDLER:
    Grid_Adm_Permisos_Consulta_Resultados.Rows = 0
    Grid_Adm_Permisos_Consulta_Resultados.Cols = 9
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT AM.No_Movimiento,AM.Empleado_ID,AM.Departamento_ID,"
    Mi_SQL = Mi_SQL & " AM.Fecha_Inicio,AM.Fecha_Termino,AM.Estatus, AM.Simbologia, AM.SubSimbologia,"
    Mi_SQL = Mi_SQL & " AM.Motivo, AM.Observaciones,"
    Mi_SQL = Mi_SQL & " (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias AM,Cat_Empleados CE, Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE AM.Empleado_ID=CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND AM.Departamento_ID=CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND AM.Tipo_Incidencia='P'"
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
            Mi_SQL = Mi_SQL & " AND AM.Tipo_Falta_ID='" & Format(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ItemData(Cmb_Adm_Permisos_Consulta_Tipo_Permiso.ListIndex), "00000") & "'"
        Else
            MsgBox "No ha seleccionado tipo de permiso", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
    End If
    'SubTipo
    If Chk_Adm_Permisos_Consulta_Sub_Tipo_Permiso.Value = 1 Then
        If Cmb_Adm_Permisos_Consulta_Sub_Tipo_Permiso.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND AM.Simbologia = '" & Simbologia_Consulta & "'"
        Else
            MsgBox "No ha seleccionado subtipo de permiso", vbInformation + vbOKOnly, Me.Caption
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
        'Coloca un encabezado en el grid
        'Agrega el encabezado al reporte
        If Chk_Adm_Permisos_Consulta_Periodo.Value = 1 Then
            Call Encabezado_Reporte("REPORTE SOLICITUD DE PERMISOS", DateAdd("s", 1, Dtp_Adm_Permisos_Consulta_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Adm_Permisos_Consulta_Fecha_Termino.Value))
        Else
            Call Encabezado_Reporte("REPORTE SOLICITUD DE PERMISOS")
        End If
        Print #2, "No_Movimiento|Nombre|F. Inicio|F. Termino|Tipo|SubTipo|Estatus|Motivo-Horas|Observaciones"
        Grid_Adm_Permisos_Consulta_Resultados.AddItem "No_Movimiento" & Chr(9) & "Nombre" & Chr(9) & "F. Inicio" & Chr(9) & _
                        "F. Termino" & Chr(9) & "Tipo" & Chr(9) & "SubTipo" & Chr(9) & "Estatus" & Chr(9) & _
                        "Motivo-Horas" & Chr(9) & "Observaciones"
        While Not .EOF
            Grid_Adm_Permisos_Consulta_Resultados.AddItem .rdoColumns("No_Movimiento") _
            & Chr(9) & .rdoColumns("Nombre") _
            & Chr(9) & .rdoColumns("Fecha_Inicio") _
            & Chr(9) & .rdoColumns("Fecha_Termino") _
            & Chr(9) & .rdoColumns("Simbologia") _
            & Chr(9) & .rdoColumns("SubSimbologia") _
            & Chr(9) & .rdoColumns("Estatus") _
            & Chr(9) & .rdoColumns("Motivo") _
            & Chr(9) & .rdoColumns("Observaciones")
            Print #2, .rdoColumns("No_Movimiento") _
            & "|" & .rdoColumns("Nombre") _
            & "|" & .rdoColumns("Fecha_Inicio") _
            & "|" & .rdoColumns("Fecha_Termino") _
            & "|" & .rdoColumns("Simbologia") _
            & "|" & .rdoColumns("SubSimbologia") _
            & "|" & .rdoColumns("Estatus") _
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
            .ColWidth(1) = 2500 'Empleado
            .ColWidth(2) = 1000 'Fecha_Inicio
            .ColWidth(3) = 1000 'Fecha_Termino
            .ColWidth(4) = 800  'Simbologia
            .ColWidth(5) = 800  'Subsimbologia
            .ColWidth(6) = 650  'Estatus
            .ColWidth(7) = 0  'Motivo_Horas
            .ColWidth(8) = 0  'Observaciones
        End With
    End If
    End With
    Set Rs_Consulta_Adm_Movimientos = Nothing
    Me.MousePointer = 0
Exit Sub
HANDLER:
Me.MousePointer = 0
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Llenar_Informacion_Permiso
'DESCRIPCION: Llena la información del permiso en la interfaz
'PARAMETROS : No_Movimiento- Es el número de permiso seleccionado
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 25-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Llenar_Informacion_Permiso(No_Movimiento As String)
Dim Rs_Consulta_Adm_Movimientos As rdoResultset     'Informacion del permiso
Dim Cont_List_Permisos As Integer

    'Informacion del permiso
    Mi_SQL = "SELECT AM.*, ISNULL(CE.Supervisor_ID,'') as Supervisor_ID "
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias AM, Cat_Empleados CE"
    Mi_SQL = Mi_SQL & " WHERE AM.Empleado_ID = CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND No_Movimiento = '" & No_Movimiento & "'"
    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = 'P'"
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
            'Selecciona el item del permiso
            For Cont_List_Permisos = 0 To List_Tipo_Permiso.ListCount - 1
                If .rdoColumns("Tipo_Falta_ID") = Mid(List_Tipo_Permiso.List(Cont_List_Permisos), 1, 5) Then
                    List_Tipo_Permiso.Selected(Cont_List_Permisos) = True
                End If
            Next
            If .rdoColumns("Estatus") = "A" Then
                Cmb_Permisos_Estatus.ListIndex = 0
            Else
                Cmb_Permisos_Estatus.ListIndex = 1
            End If
            Pic_Solicitud_Permisos.ZOrder vbBringToFront
            Pic_Solicitud_Permisos_Consulta.Visible = False
        End With
    End If
    Rs_Consulta_Adm_Movimientos.Close
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
    Open Ruta_Temporal & "Solicitud_Permisos" & ".txt" For Output As #1
    Open Ruta_Temporal & "Solicitud_Permisos" & "xls.txt" For Output As #2 'Reporte a xls
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
'*******************************************************************************************************
'NOMBRE_FUNCION:
'DESCRIPCION:
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************************************
Private Sub Mostrar_PDF(Nombre_Archivo As String, Ruta_Archivo As String)
Dim Archivo_Pdf As String
Dim Extension As String
Dim Frm_PDFS As New Frm_Previo_Pdf

On Error GoTo HANDLER
    Extension = Right(Nombre_Archivo, 3)
    Frm_PDFS.Label1.Caption = "REPORTES"
    'Frm_Solicitud_Cotizacion_.lblproyecto.Caption = Cmb_Nombre_Cliente.Text
    Frm_PDFS.OLE1.SourceDoc = Ruta_Archivo
    Frm_PDFS.OLE1.SourceItem = Nombre_Archivo
    Archivo_Pdf = Frm_PDFS.OLE1.SourceDoc
    If Dir(Frm_PDFS.OLE1.SourceDoc) <> "" Then
        'Carga la forma de ver archivo.
        Load Frm_PDFS
        If UCase(Extension) = "PDF" Then
            Frm_PDFS.AcroPDF1.Visible = True
            Frm_PDFS.OLE1.Visible = False
            Frm_PDFS.AcroPDF1.src = Archivo_Pdf
            Frm_PDFS.AcroPDF1.LoadFile (Archivo_Pdf)
        End If
    Else
        Unload Frm_PDFS
        MsgBox "El archivo que está intentando abrir no se encontró en el directorio indicado.", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    MsgBox Err.Description
End Sub

