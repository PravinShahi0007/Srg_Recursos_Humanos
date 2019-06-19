VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Empleados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8415
   Begin VB.CommandButton Btn_Imprimir_Contrato 
      Caption         =   "Contrato"
      Height          =   555
      Left            =   5640
      Picture         =   "Frm_Cat_Empleados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   155
      Tag             =   "A"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1160
   End
   Begin VB.CommandButton Btn_Salir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   555
      Left            =   7020
      Picture         =   "Frm_Cat_Empleados.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1160
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   4200
      Picture         =   "Frm_Cat_Empleados.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "C"
      Top             =   6600
      Width           =   1160
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   555
      Left            =   1440
      Picture         =   "Frm_Cat_Empleados.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "M"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1160
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   555
      Left            =   2760
      Picture         =   "Frm_Cat_Empleados.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "B"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1160
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   45
      Picture         =   "Frm_Cat_Empleados.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "A"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1160
   End
   Begin VB.PictureBox Pic_Cat_Empleados 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6495
      ScaleWidth      =   8400
      TabIndex        =   101
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Empleados 
         BackColor       =   &H8000000E&
         Caption         =   "Empleados"
         Height          =   1695
         Left            =   120
         TabIndex        =   152
         Top             =   4800
         Width           =   8175
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Empleados 
            Height          =   1335
            Left            =   120
            TabIndex        =   153
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin TabDlg.SSTab Tab_Cat_Empleados 
         Height          =   4290
         Left            =   90
         TabIndex        =   23
         Top             =   450
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   7567
         _Version        =   393216
         Tabs            =   6
         Tab             =   2
         TabsPerRow      =   6
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "Personales"
         TabPicture(0)   =   "Frm_Cat_Empleados.frx":213C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Fra_Cat_Empleados_Datos_Personales"
         Tab(0).Control(1)=   "Btn_Imprimir"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Domicilio"
         TabPicture(1)   =   "Frm_Cat_Empleados.frx":2158
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Cat_Empleados_Datos_Dependientes"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Laborales"
         TabPicture(2)   =   "Frm_Cat_Empleados.frx":2174
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Fra_Cat_Empleados_Datos_Laborales"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Otros"
         TabPicture(3)   =   "Frm_Cat_Empleados.frx":2190
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Fra_Cat_Empleados_Datos_Otros"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Cursos"
         TabPicture(4)   =   "Frm_Cat_Empleados.frx":21AC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Fra_Cursos"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Adicionales"
         TabPicture(5)   =   "Frm_Cat_Empleados.frx":21C8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Fra_Evaluaciones"
         Tab(5).ControlCount=   1
         Begin VB.Frame Fra_Evaluaciones 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3420
            Left            =   -74910
            TabIndex        =   132
            Top             =   390
            Width           =   8085
            Begin VB.TextBox Txt_Campo_5 
               Height          =   600
               Left            =   1095
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   149
               Top             =   2775
               Width           =   6915
            End
            Begin VB.TextBox Txt_Campo_4 
               Height          =   600
               Left            =   1095
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   147
               Top             =   2124
               Width           =   6915
            End
            Begin VB.TextBox Txt_Campo_3 
               Height          =   600
               Left            =   1095
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   145
               Top             =   1476
               Width           =   6915
            End
            Begin VB.TextBox Txt_Campo_2 
               Height          =   600
               Left            =   1095
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   143
               Top             =   828
               Width           =   6915
            End
            Begin VB.TextBox Txt_Campo_1 
               Height          =   600
               Left            =   1095
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   141
               Top             =   180
               Width           =   6915
            End
            Begin VB.TextBox Txt_Evaluacion 
               Height          =   315
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   71
               Top             =   3510
               Width           =   3240
            End
            Begin VB.CommandButton Btn_Eliminar_Evaluaciones 
               Caption         =   "Eliminar"
               Height          =   255
               Left            =   6945
               TabIndex        =   134
               Top             =   6390
               Width           =   1005
            End
            Begin VB.CommandButton Btn_Agregar_Evaluacion 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   7170
               TabIndex        =   133
               Top             =   3495
               Width           =   750
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Evaluacion 
               Height          =   315
               Left            =   4350
               TabIndex        =   72
               Top             =   3510
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MMM/yyyy"
               Format          =   113704963
               CurrentDate     =   41020
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Evaluaciones 
               Height          =   2535
               Left            =   75
               TabIndex        =   135
               Top             =   3840
               Width           =   7875
               _ExtentX        =   13891
               _ExtentY        =   4471
               _Version        =   393216
               Rows            =   0
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               ScrollBars      =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               Appearance      =   0
            End
            Begin MSComCtl2.DTPicker Dtp_Proxima_Evaluacion 
               Height          =   315
               Left            =   5760
               TabIndex        =   73
               Top             =   3510
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MMM/yyyy"
               Format          =   113704963
               CurrentDate     =   41020
            End
            Begin VB.Label Lbl_Campo_5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Campo 5"
               Height          =   195
               Left            =   135
               TabIndex        =   148
               Top             =   2978
               Width           =   630
            End
            Begin VB.Label Lbl_Campo_4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Campo 4"
               Height          =   195
               Left            =   135
               TabIndex        =   146
               Top             =   2327
               Width           =   630
            End
            Begin VB.Label Lbl_Campo_3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Campo 3"
               Height          =   195
               Left            =   135
               TabIndex        =   144
               Top             =   1679
               Width           =   630
            End
            Begin VB.Label Lbl_Campo_2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Campo 2"
               Height          =   195
               Left            =   135
               TabIndex        =   142
               Top             =   1031
               Width           =   630
            End
            Begin VB.Label Lbl_Campo_1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Campo 1"
               Height          =   195
               Left            =   135
               TabIndex        =   140
               Top             =   383
               Width           =   630
            End
            Begin VB.Label Lbl_Evaluacion 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Evaluacion"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   136
               Top             =   3570
               Width           =   795
            End
         End
         Begin VB.Frame Fra_Cursos 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3420
            Left            =   -74910
            TabIndex        =   127
            Top             =   390
            Width           =   8085
            Begin VB.TextBox Txt_Horas_Curso 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6675
               Locked          =   -1  'True
               TabIndex        =   131
               Top             =   555
               Width           =   1275
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio 
               Height          =   315
               Left            =   3105
               TabIndex        =   66
               Top             =   915
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MMM/yyyy"
               Format          =   113704963
               CurrentDate     =   41020
            End
            Begin VB.ComboBox Cmb_Estatus_Curso 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":21E4
               Left            =   1305
               List            =   "Frm_Cat_Empleados.frx":21F1
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   915
               Width           =   1800
            End
            Begin VB.CommandButton Btn_Agregar_Curso 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   5910
               TabIndex        =   68
               Top             =   915
               Width           =   1005
            End
            Begin VB.CommandButton Btn_Eliminar_Curso 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   6930
               TabIndex        =   70
               Top             =   915
               Width           =   1005
            End
            Begin VB.ComboBox Cmb_Curso 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2217
               Left            =   1305
               List            =   "Frm_Cat_Empleados.frx":2219
               TabIndex        =   63
               Top             =   195
               Width           =   6630
            End
            Begin VB.TextBox Txt_Comentarios_Curso 
               Height          =   315
               Left            =   1305
               MaxLength       =   50
               TabIndex        =   64
               Top             =   555
               Width           =   5340
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Cursos 
               Height          =   2040
               Left            =   105
               TabIndex        =   69
               Top             =   1290
               Width           =   7830
               _ExtentX        =   13811
               _ExtentY        =   3598
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               ScrollBars      =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               Appearance      =   0
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Fin 
               Height          =   315
               Left            =   4515
               TabIndex        =   67
               Top             =   915
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MMM/yyyy"
               Format          =   113704963
               CurrentDate     =   41020
            End
            Begin VB.Label Lbl_Curso 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Estatus"
               Height          =   195
               Index           =   3
               Left            =   165
               TabIndex        =   130
               Top             =   975
               Width           =   525
            End
            Begin VB.Label Lbl_Curso 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Instructor"
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   129
               Top             =   615
               Width           =   660
            End
            Begin VB.Label Lbl_Curso 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Curso"
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   128
               Top             =   255
               Width           =   405
            End
         End
         Begin VB.CommandButton Btn_Imprimir 
            Height          =   675
            Left            =   -67800
            Picture         =   "Frm_Cat_Empleados.frx":221B
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "C"
            ToolTipText     =   "Imprimir Hoja de Registro"
            Top             =   3240
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.Frame Fra_Cat_Empleados_Datos_Otros 
            BackColor       =   &H00FFFFFF&
            Height          =   3480
            Left            =   -74880
            TabIndex        =   74
            Top             =   360
            Width           =   8070
            Begin VB.ComboBox Cmb_Transporte 
               Height          =   315
               Left            =   1320
               TabIndex        =   53
               Top             =   165
               Width           =   6675
            End
            Begin VB.TextBox Txt_Cat_Empleados_Alergias1 
               Height          =   315
               Left            =   90
               MaxLength       =   50
               TabIndex        =   57
               Top             =   1230
               Width           =   7890
            End
            Begin VB.TextBox Txt_Cat_Empleados_Alergias2 
               Height          =   315
               Left            =   90
               MaxLength       =   50
               TabIndex        =   58
               Top             =   1560
               Width           =   7890
            End
            Begin VB.TextBox Txt_Cat_Empleados_Alergias3 
               Height          =   315
               Left            =   90
               MaxLength       =   50
               TabIndex        =   59
               Top             =   1905
               Width           =   7890
            End
            Begin VB.TextBox Txt_Cat_Empleados_Llamar_Telefono2 
               Height          =   315
               Left            =   6120
               MaxLength       =   50
               TabIndex        =   56
               Top             =   690
               Width           =   1860
            End
            Begin VB.TextBox Txt_Cat_Empleados_Llamar_Telefono1 
               Height          =   315
               Left            =   4140
               MaxLength       =   50
               TabIndex        =   55
               Top             =   690
               Width           =   1860
            End
            Begin VB.TextBox Txt_Cat_Empleados_LLamar_A 
               Height          =   315
               Left            =   90
               MaxLength       =   50
               TabIndex        =   54
               Top             =   690
               Width           =   3885
            End
            Begin VB.Frame Fra_Cat_Empleados_Baja 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Baja"
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
               Height          =   1230
               Left            =   90
               TabIndex        =   75
               Top             =   2190
               Width           =   7845
               Begin VB.TextBox Txt_Cat_Empleados_Observaciones_Baja 
                  Height          =   555
                  Left            =   60
                  MaxLength       =   1000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   62
                  Top             =   615
                  Width           =   7665
               End
               Begin VB.ComboBox Cmb_Cat_Empleados_Motivos_Baja 
                  Height          =   315
                  ItemData        =   "Frm_Cat_Empleados.frx":2AE5
                  Left            =   2880
                  List            =   "Frm_Cat_Empleados.frx":2AF5
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   270
                  Width           =   4845
               End
               Begin MSComCtl2.DTPicker Dtp_Cat_Empleados_Fecha_Baja 
                  Height          =   315
                  Left            =   630
                  TabIndex        =   60
                  Top             =   270
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   556
                  _Version        =   393216
                  CustomFormat    =   "dd MMM yyyy"
                  Format          =   113704963
                  CurrentDate     =   39941
               End
               Begin VB.Label Lbl_Baja 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Motivo"
                  Height          =   195
                  Index           =   1
                  Left            =   2250
                  TabIndex        =   76
                  Top             =   315
                  Width           =   480
               End
               Begin VB.Label Lbl_Baja 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "F. Baja"
                  Height          =   195
                  Index           =   0
                  Left            =   45
                  TabIndex        =   77
                  Top             =   330
                  Width           =   495
               End
            End
            Begin VB.Label Lbl_Transporte 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Transporte"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   105
               TabIndex        =   119
               Top             =   210
               Width           =   930
            End
            Begin VB.Label Lbl_Otros 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Alergias ..."
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   78
               Top             =   1005
               Width           =   735
            End
            Begin VB.Label Lbl_Otros 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "ó  Al Telefono"
               Height          =   195
               Index           =   2
               Left            =   6570
               TabIndex        =   79
               Top             =   465
               Width           =   990
            End
            Begin VB.Label Lbl_Otros 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Al Telefono"
               Height          =   195
               Index           =   1
               Left            =   4665
               TabIndex        =   80
               Top             =   465
               Width           =   810
            End
            Begin VB.Label Lbl_Otros 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "En caso de Emergencia llamar a..."
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   83
               Top             =   465
               Width           =   2415
            End
         End
         Begin VB.Frame Fra_Cat_Empleados_Datos_Laborales 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3525
            Left            =   45
            TabIndex        =   84
            Top             =   345
            Width           =   8205
            Begin VB.TextBox Txt_Clave_SAP 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4470
               MaxLength       =   10
               TabIndex        =   41
               Top             =   1665
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Subdivision 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2B17
               Left            =   6825
               List            =   "Frm_Cat_Empleados.frx":2B24
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   1296
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Gap 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2B37
               Left            =   1260
               List            =   "Frm_Cat_Empleados.frx":2B3E
               TabIndex        =   50
               Top             =   3135
               Width           =   4560
            End
            Begin VB.TextBox Txt_Cat_Empleados_Vacaciones 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4470
               MaxLength       =   10
               TabIndex        =   48
               Top             =   2790
               Width           =   1335
            End
            Begin VB.TextBox Txt_Cat_Empleados_Seccion 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   52
               Top             =   2790
               Width           =   2130
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Contratacion 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2B5C
               Left            =   4470
               List            =   "Frm_Cat_Empleados.frx":2B66
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   2415
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Tipo_Empleado 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2B7C
               Left            =   1260
               List            =   "Frm_Cat_Empleados.frx":2B86
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   2415
               Width           =   2145
            End
            Begin VB.TextBox Txt_Cat_Empleados_Antiguedad 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4470
               Locked          =   -1  'True
               TabIndex        =   82
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Puesto 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2BA4
               Left            =   1260
               List            =   "Frm_Cat_Empleados.frx":2BAB
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   1296
               Width           =   4605
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Turno 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2BC9
               Left            =   6825
               List            =   "Frm_Cat_Empleados.frx":2BCB
               TabIndex        =   44
               Text            =   "Cmb_Cat_Empleados_Turno"
               Top             =   2040
               Width           =   1335
            End
            Begin VB.CheckBox Chk_Cat_Empleados_Trabaja_Domingos 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Trabaja Domingos"
               Height          =   240
               Left            =   5940
               TabIndex        =   51
               Top             =   3165
               Width           =   1635
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Tipo 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2BCD
               Left            =   1260
               List            =   "Frm_Cat_Empleados.frx":2BD7
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1665
               Width           =   2145
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Empresa 
               Height          =   315
               Left            =   1260
               TabIndex        =   35
               Text            =   "Cmb_Cat_Empleados_Empresa"
               Top             =   180
               Width           =   6900
            End
            Begin VB.TextBox Txt_Cat_Empleados_Salario_Diario 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6825
               MaxLength       =   10
               TabIndex        =   49
               Top             =   2760
               Width           =   1335
            End
            Begin VB.TextBox Txt_Cat_Empleados_No_Tarjeta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6825
               MaxLength       =   10
               TabIndex        =   42
               Top             =   1665
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Supervisor 
               Height          =   315
               Left            =   1260
               TabIndex        =   36
               Text            =   "Cmb_Cat_Empleados_Supervisor"
               Top             =   552
               Width           =   6900
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Departamento 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":2BF1
               Left            =   1260
               List            =   "Frm_Cat_Empleados.frx":2BF3
               TabIndex        =   37
               Text            =   "Cmb_Cat_Empleados_Departamento"
               Top             =   924
               Width           =   6900
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Empleados_Fecha_Ingreso 
               Height          =   315
               Left            =   1260
               TabIndex        =   43
               Top             =   2040
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   113704963
               CurrentDate     =   39941
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Empleados_Fecha_Termino_Contrato 
               Height          =   315
               Left            =   6825
               TabIndex        =   47
               Top             =   2415
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   113704963
               CurrentDate     =   39941
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Clave SAP"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   3405
               TabIndex        =   139
               Top             =   1725
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Subdivisión"
               Height          =   195
               Index           =   12
               Left            =   5895
               TabIndex        =   138
               Top             =   1350
               Width           =   810
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tripulacion"
               Height          =   195
               Index           =   15
               Left            =   45
               TabIndex        =   126
               Top             =   3195
               Width           =   780
            End
            Begin VB.Label Lbl_Cat_Empleados_Fecha_Termino_Contrato 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "F. Termino"
               Height          =   195
               Left            =   5865
               TabIndex        =   85
               Top             =   2475
               Visible         =   0   'False
               Width           =   750
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vacaciones"
               Height          =   195
               Index           =   14
               Left            =   3405
               TabIndex        =   86
               Top             =   2850
               Width           =   840
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Seccion"
               Height          =   195
               Index           =   13
               Left            =   45
               TabIndex        =   87
               Top             =   2850
               Width           =   585
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Contratación"
               Height          =   195
               Index           =   11
               Left            =   3405
               TabIndex        =   88
               Top             =   2475
               Width           =   900
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tipo MO"
               Height          =   195
               Index           =   10
               Left            =   45
               TabIndex        =   89
               Top             =   2475
               Width           =   615
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Antigüedad"
               Height          =   195
               Index           =   8
               Left            =   3405
               TabIndex        =   90
               Top             =   2100
               Width           =   810
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Puesto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   45
               TabIndex        =   91
               Top             =   1356
               Width           =   600
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Supervisor ?"
               Height          =   195
               Index           =   4
               Left            =   45
               TabIndex        =   92
               Top             =   1725
               Width           =   885
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Empresa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   93
               Top             =   225
               Width           =   735
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "S. Diario"
               Height          =   195
               Index           =   5
               Left            =   5865
               TabIndex        =   94
               Top             =   2820
               Width           =   600
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "No.Nomina"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   5865
               TabIndex        =   95
               Top             =   1725
               Width           =   945
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Turno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   5865
               TabIndex        =   96
               Top             =   2100
               Width           =   510
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Supervisor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   45
               TabIndex        =   97
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "F. Ingreso"
               Height          =   195
               Index           =   7
               Left            =   45
               TabIndex        =   98
               Top             =   2100
               Width           =   705
            End
            Begin VB.Label Lbl_Laborales 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Departamento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   45
               TabIndex        =   99
               Top             =   975
               Width           =   1200
            End
         End
         Begin VB.Frame Fra_Cat_Empleados_Datos_Personales 
            BackColor       =   &H00FFFFFF&
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
            Height          =   3855
            Left            =   -74940
            TabIndex        =   104
            Top             =   360
            Width           =   8175
            Begin VB.TextBox Txt_Cat_Empleados_Ruta_Imagen 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   100
               Top             =   1620
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.CommandButton Btn_Cat_Empleados_Foto 
               Caption         =   "Agregar Foto"
               Height          =   315
               Left            =   6240
               TabIndex        =   21
               Top             =   1620
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.TextBox Txt_Cat_Empleados_Email 
               Height          =   285
               Left            =   1200
               TabIndex        =   18
               Top             =   3480
               Width           =   4935
            End
            Begin VB.CommandButton Btn_Huella_Comedor 
               Height          =   675
               Left            =   6240
               MouseIcon       =   "Frm_Cat_Empleados.frx":2BF5
               Picture         =   "Frm_Cat_Empleados.frx":34BF
               Style           =   1  'Graphical
               TabIndex        =   151
               Tag             =   "C"
               ToolTipText     =   "Registro de Huella para Comedor"
               Top             =   2880
               UseMaskColor    =   -1  'True
               Width           =   855
            End
            Begin VB.TextBox Txt_Cat_Empleados_Estatus_Puesto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   137
               Top             =   3075
               Width           =   1335
            End
            Begin VB.TextBox Txt_Cat_Empleados_Lugar_Nacimiento 
               Height          =   315
               Left            =   1215
               MaxLength       =   50
               TabIndex        =   10
               Top             =   1605
               Width           =   4920
            End
            Begin VB.TextBox Txt_Cat_Empleados_Clave_Elector 
               Height          =   315
               Left            =   6210
               TabIndex        =   14
               Top             =   2400
               Width           =   1830
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Nivel_Estudio 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":3901
               Left            =   1215
               List            =   "Frm_Cat_Empleados.frx":3903
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   1965
               Width           =   4350
            End
            Begin VB.TextBox Txt_Cat_Empleados_NSS 
               Height          =   315
               Left            =   3735
               TabIndex        =   17
               Top             =   3090
               Width           =   1830
            End
            Begin VB.TextBox Txt_Cat_Empleados_Curp 
               Height          =   315
               Left            =   3735
               TabIndex        =   16
               Top             =   2715
               Width           =   1830
            End
            Begin VB.TextBox Txt_Cat_Empleados_RFC 
               Height          =   315
               Left            =   1215
               TabIndex        =   15
               Top             =   2715
               Width           =   1875
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Sexo 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":3905
               Left            =   6210
               List            =   "Frm_Cat_Empleados.frx":390F
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   1965
               Width           =   1830
            End
            Begin VB.TextBox Txt_Cat_Empleados_Edad 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3735
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   2355
               Width           =   1830
            End
            Begin VB.TextBox Txt_Cat_Empleados_Apellido_Materno 
               Height          =   315
               Left            =   1215
               MaxLength       =   50
               TabIndex        =   9
               Top             =   1260
               Width           =   4920
            End
            Begin VB.TextBox Txt_Cat_Empleados_Apellido_Paterno 
               Height          =   315
               Left            =   1215
               MaxLength       =   50
               TabIndex        =   8
               Top             =   900
               Width           =   4920
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Estatus 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":3928
               Left            =   3780
               List            =   "Frm_Cat_Empleados.frx":3932
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   180
               Width           =   2370
            End
            Begin VB.TextBox Txt_Cat_Empleados_Empleado_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1215
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   180
               Width           =   1875
            End
            Begin VB.TextBox Txt_Cat_Empleados_Nombre 
               Height          =   315
               Left            =   1215
               MaxLength       =   50
               TabIndex        =   7
               Top             =   540
               Width           =   4920
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Empleados_Fecha_Nacimiento 
               Height          =   315
               Left            =   1215
               TabIndex        =   13
               Top             =   2340
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   113704963
               CurrentDate     =   39941
            End
            Begin MSComDlg.CommonDialog Cmd_Cat_Empleados_Foto 
               Left            =   7470
               Top             =   225
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Estado_Civil 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":3948
               Left            =   1215
               List            =   "Frm_Cat_Empleados.frx":3964
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   3090
               Width           =   1875
            End
            Begin VB.CheckBox Chk_Cat_Empleados_Infonavit 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tiene Infonavit"
               Height          =   210
               Left            =   3120
               TabIndex        =   20
               Top             =   3840
               Width           =   1635
            End
            Begin VB.Label Label1 
               BackColor       =   &H8000000E&
               Caption         =   "Email"
               Height          =   255
               Left            =   120
               TabIndex        =   154
               Top             =   3480
               Width           =   855
            End
            Begin VB.Label Lbl_NSS 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "NSS"
               Height          =   195
               Left            =   3165
               TabIndex        =   150
               Top             =   3150
               Width           =   330
            End
            Begin VB.Label Lbl_Estado_Civil 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Estado Civil"
               Height          =   195
               Left            =   90
               TabIndex        =   111
               Top             =   3150
               Width           =   825
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "L. Nacimiento"
               Height          =   195
               Left            =   90
               TabIndex        =   112
               Top             =   1665
               Width           =   975
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "P. Académico"
               Height          =   195
               Left            =   90
               TabIndex        =   113
               Top             =   2025
               Width           =   990
            End
            Begin VB.Image Img_Cat_Empleados_Foto 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1755
               Left            =   6255
               Picture         =   "Frm_Cat_Empleados.frx":39B0
               Stretch         =   -1  'True
               ToolTipText     =   "Doble click para cambiar la imagen"
               Top             =   180
               Width           =   1725
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "RFC                                             CURP"
               Height          =   195
               Left            =   90
               TabIndex        =   114
               Top             =   2775
               Width           =   2790
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Sexo"
               Height          =   195
               Left            =   5625
               TabIndex        =   115
               Top             =   2025
               Width           =   360
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Nacimiento                                                  Edad                                               C.Elec"
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   2400
               Width           =   6000
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "A. Materno"
               Height          =   195
               Left            =   90
               TabIndex        =   109
               Top             =   1320
               Width           =   780
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "A. Paterno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   108
               Top             =   960
               Width           =   915
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Estatus"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3105
               TabIndex        =   107
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Empleado ID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   106
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   105
               Top             =   600
               Width           =   660
            End
         End
         Begin VB.Frame Fra_Cat_Empleados_Datos_Dependientes 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3525
            Left            =   -74955
            TabIndex        =   103
            Top             =   345
            Width           =   8175
            Begin VB.TextBox Txt_Cat_Empleados_Estado 
               Height          =   315
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   28
               Top             =   870
               Width           =   2625
            End
            Begin VB.TextBox Txt_Cat_Empleados_CP 
               Height          =   315
               Left            =   5400
               MaxLength       =   50
               TabIndex        =   26
               Top             =   540
               Width           =   2625
            End
            Begin VB.TextBox Txt_Cat_Empleados_Ciudad 
               Height          =   315
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   27
               Top             =   870
               Width           =   3315
            End
            Begin VB.TextBox Txt_Cat_Empleados_Colonia 
               Height          =   315
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   25
               Top             =   540
               Width           =   3315
            End
            Begin VB.TextBox Txt_Cat_Empleados_Direccion 
               Height          =   315
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   24
               Top             =   195
               Width           =   6915
            End
            Begin VB.TextBox Txt_Cat_Empleados_Dependiente_Nombre 
               Height          =   315
               Left            =   1110
               MaxLength       =   50
               TabIndex        =   29
               Top             =   1425
               Width           =   5595
            End
            Begin VB.CommandButton Btn_Cat_Empleados_Dependientes_Eliminar 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   6780
               TabIndex        =   34
               Top             =   1770
               Width           =   1335
            End
            Begin VB.CommandButton Btn_Cat_Empleados_Dependientes_Agregar 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   6780
               TabIndex        =   32
               Top             =   1410
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Parentesco 
               Height          =   315
               ItemData        =   "Frm_Cat_Empleados.frx":F9F2
               Left            =   1110
               List            =   "Frm_Cat_Empleados.frx":FA08
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   1770
               Width           =   2370
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Empleados_Dependientes 
               Height          =   1290
               Left            =   75
               TabIndex        =   33
               Top             =   2130
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   2275
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               ScrollBars      =   2
               AllowUserResizing=   1
               Appearance      =   0
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento 
               Height          =   315
               Left            =   4980
               TabIndex        =   31
               Top             =   1770
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   113704963
               CurrentDate     =   39941
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Nombre"
               Height          =   195
               Index           =   8
               Left            =   75
               TabIndex        =   125
               Top             =   1470
               Width           =   555
            End
            Begin VB.Line Line1 
               X1              =   1185
               X2              =   8130
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Estado"
               Height          =   195
               Index           =   7
               Left            =   4530
               TabIndex        =   124
               Top             =   930
               Width           =   495
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "C. Postal"
               Height          =   195
               Index           =   6
               Left            =   4530
               TabIndex        =   123
               Top             =   600
               Width           =   630
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ciudad"
               Height          =   195
               Index           =   5
               Left            =   75
               TabIndex        =   122
               Top             =   930
               Width           =   495
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Colonia"
               Height          =   195
               Index           =   4
               Left            =   75
               TabIndex        =   121
               Top             =   600
               Width           =   525
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Direccion"
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   120
               Top             =   255
               Width           =   675
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "F. Nacimiento"
               Height          =   195
               Index           =   3
               Left            =   3675
               TabIndex        =   116
               Top             =   1830
               Width           =   975
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Parentescos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   75
               TabIndex        =   117
               Top             =   1200
               Width           =   1065
            End
            Begin VB.Label Lbl_Domicilio 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Parentesco"
               Height          =   195
               Index           =   2
               Left            =   75
               TabIndex        =   118
               Top             =   1830
               Width           =   810
            End
         End
      End
      Begin VB.TextBox Txt_Log 
         Height          =   1365
         Left            =   3840
         MultiLine       =   -1  'True
         TabIndex        =   156
         Top             =   5040
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMPLEADOS"
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
         Left            =   3000
         TabIndex        =   102
         Top             =   15
         Width           =   2385
      End
   End
   Begin VB.Image Img_Logo_Empresa 
      Height          =   450
      Left            =   0
      Picture         =   "Frm_Cat_Empleados.frx":FA36
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "Frm_Cat_Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Imagen_Logo As String
Option Explicit
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
' Constantes
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260 ' Para Buffer de caracteres del path
' Funcion Api CoTaskMemFree
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
' Funcion Api CoTaskMemFree lstrcat
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
' Funcion Api SHBrowseForFolder
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
' Funcion Api SHGetPathFromIDList
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios
Public Catalogo As String          'Indicar que formulario se va a abrir
Private Sub Btn_Agregar_Curso_Click()
Dim Fila As Integer
    'Valida los datos del dependiente
    If Cmb_Curso.ListIndex > -1 And Trim(Txt_Comentarios_Curso.Text) <> "" Then
        'Agrega el dependiente a la lista
        Grid_Cursos.Cols = 6
        If Grid_Cursos.Rows = 0 Then
            Grid_Cursos.AddItem "Curso_ID" & Chr(9) & "Curso" _
                & Chr(9) & "Instructor" & Chr(9) & "Estatus" _
                & Chr(9) & "Inicio" & Chr(9) & "Fin"
            Grid_Cursos.ColWidth(0) = 0     'Curso_ID
            Grid_Cursos.ColWidth(1) = 2500  'Curso
            Grid_Cursos.ColWidth(2) = 1400  'Instructor
            Grid_Cursos.ColWidth(3) = 1200  'Estatus
            Grid_Cursos.ColWidth(4) = 1200  'Inicio
            Grid_Cursos.ColWidth(5) = 1200  'Fin
        End If
        'Valida que no este ya dado de alta el curso en el listado
        For Fila = 1 To Grid_Cursos.Rows - 1
            If Format(Cmb_Curso.ItemData(Cmb_Curso.ListIndex), "00000") = Trim(Grid_Cursos.TextMatrix(Fila, 0)) Then
                MsgBox "El curso ya ha sido agregado previamente", vbExclamation
                Exit Sub
            End If
        Next
        Grid_Cursos.AddItem Format(Cmb_Curso.ItemData(Cmb_Curso.ListIndex), "00000") _
            & Chr(9) & Cmb_Curso.Text _
            & Chr(9) & Trim(Txt_Comentarios_Curso.Text) _
            & Chr(9) & Cmb_Estatus_Curso.Text _
            & Chr(9) & Format(Dtp_Fecha_Inicio.Value, "dd/MMM/yyyy") _
            & Chr(9) & Format(Dtp_Fecha_Fin.Value, "dd/MMM/yyyy")
        Grid_Cursos.FixedRows = 1
        Cmb_Curso.Text = ""
        Txt_Comentarios_Curso.Text = ""
        Txt_Horas_Curso.Text = ""
        Dtp_Fecha_Inicio.Value = Now
        Dtp_Fecha_Fin.Value = Now
    Else
        MsgBox "Faltan datos para poder agregar el curso", vbExclamation
    End If
End Sub

Private Sub Btn_Agregar_Evaluacion_Click()
Dim Fila As Integer
    'Valida los datos del dependiente
    If Trim(Txt_Evaluacion.Text) <> "" Then
        'Agrega el dependiente a la lista
        If Grid_Evaluaciones.Rows = 0 Then
            Grid_Evaluaciones.AddItem "Evaluacion" & Chr(9) & "Fecha" & Chr(9) & "Siguiente"
            Grid_Evaluaciones.ColWidth(0) = 4900  'Evaluacion
            Grid_Evaluaciones.ColWidth(1) = 1300  'Inicio
            Grid_Evaluaciones.ColWidth(2) = 1300  'Fin
        End If
        Grid_Evaluaciones.AddItem Trim(Txt_Evaluacion.Text) _
            & Chr(9) & Format(Dtp_Fecha_Evaluacion.Value, "dd/MMM/yyyy") _
            & Chr(9) & Format(Dtp_Proxima_Evaluacion.Value, "dd/MMM/yyyy")
        Grid_Evaluaciones.FixedRows = 1
        Txt_Evaluacion.Text = ""
        Dtp_Fecha_Evaluacion.Value = Now
        Dtp_Proxima_Evaluacion.Value = Now
    Else
        MsgBox "Faltan datos para poder agregar la evaluacion", vbExclamation
    End If
End Sub

Private Sub Btn_Cat_Empleados_Dependientes_Agregar_Click()
Dim Edad As String
    'Valida los datos del dependiente
    If Trim(Txt_Cat_Empleados_Dependiente_Nombre.Text) <> "" And _
        Cmb_Cat_Empleados_Parentesco.Text <> "" Then
        'Agrega el dependiente a la lista
        Grid_Cat_Empleados_Dependientes.Cols = 4
        If Grid_Cat_Empleados_Dependientes.Rows = 0 Then
            Grid_Cat_Empleados_Dependientes.AddItem "Parentesco" & Chr(9) & "Nombre" & Chr(9) & _
                    "F. Nacimiento" & Chr(9) & "Edad"
            Grid_Cat_Empleados_Dependientes.ColWidth(0) = 1000  'Parentesco
            Grid_Cat_Empleados_Dependientes.ColWidth(1) = 2500  'Nombre
            Grid_Cat_Empleados_Dependientes.ColWidth(2) = 1300  'F. Nacimiento
            Grid_Cat_Empleados_Dependientes.ColWidth(3) = 1800  'Edad
        End If
        Edad = Calcula_Edad(Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value)
        Grid_Cat_Empleados_Dependientes.AddItem Trim(Cmb_Cat_Empleados_Parentesco.Text) & Chr(9) & _
            Trim(Txt_Cat_Empleados_Dependiente_Nombre.Text) & Chr(9) & _
            Format(Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value, "dd/MMM/yyyy") & Chr(9) & _
            Edad
        Cmb_Cat_Empleados_Parentesco.ListIndex = -1
        Txt_Cat_Empleados_Dependiente_Nombre.Text = ""
        Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value = Now
        
        Grid_Cat_Empleados_Dependientes.FixedRows = 1
    End If
End Sub

Private Sub Btn_Cat_Empleados_Dependientes_Eliminar_Click()
    If Grid_Cat_Empleados_Dependientes.Rows > 0 Then
        If Grid_Cat_Empleados_Dependientes.Rows = 2 Then
            Grid_Cat_Empleados_Dependientes.Rows = 0
        Else
            Grid_Cat_Empleados_Dependientes.RemoveItem Grid_Cat_Empleados_Dependientes.RowSel
        End If
    End If
End Sub

Private Sub Btn_Cat_Empleados_Foto_Click()
On Error GoTo HANDLER
Dim hNew2 As Long
Dim Extension_Imagen As String
Dim Punto_Extension As Integer
    'Set CancelError is True
    Cmd_Cat_Empleados_Foto.CancelError = True
    'Titulo de la ventana
    Cmd_Cat_Empleados_Foto.DialogTitle = "Seleccione el Archivo de Imagen"
    'Set flags
    Cmd_Cat_Empleados_Foto.Flags = cdlOFNHideReadOnly
    'Set filters
    Cmd_Cat_Empleados_Foto.Filter = "Archivo de Imagen (*.jpg;*.gif)|*.jpg;*.gif"
    'Specify default filter
    Cmd_Cat_Empleados_Foto.FilterIndex = 2
    'Display the Open dialog box
    Cmd_Cat_Empleados_Foto.ShowOpen
    'Display name of selected file
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Cmd_Cat_Empleados_Foto.FileName, "ARCHIVO") = True Then
        Txt_Cat_Empleados_Ruta_Imagen.Text = Cmd_Cat_Empleados_Foto.FileName
        Img_Cat_Empleados_Foto.picture = LoadPicture(Cmd_Cat_Empleados_Foto.FileName)
        
        Punto_Extension = InStrRev(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), ".")
        Extension_Imagen = Mid(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), Punto_Extension, Len(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text)))
        'Image1.picture = LoadPicture("C:\Users\desarr\Desktop\foto.jpg")
        
        hNew2 = CopyImage(Img_Cat_Empleados_Foto.picture, IMAGE_BITMAP, Val(128), Val(128), LR_COPYRETURNORG)
        OpenClipboard Me.hwnd
        EmptyClipboard
        SetClipboardData CF_BITMAP, hNew2
        CloseClipboard
        
        Img_Cat_Empleados_Foto.picture = Clipboard.GetData(2)
        SavePicture Img_Cat_Empleados_Foto.picture, Ruta_Temporal & "P_Temporal" & Extension_Imagen
        Txt_Cat_Empleados_Ruta_Imagen.Text = Ruta_Temporal & "P_Temporal" & Extension_Imagen
    End If
    Exit Sub
HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Consultar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
    
    Nombre = InputBox("Proporcione el No. Nómina, Nombre, Apellido, RFC, NSS para buscar Empleados")
    Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Select Case Catalogo
        Case "Cat_Empleados"
            Call Consulta_Cat_Empleados(Nombre)
    End Select
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo HANDLER
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Conexion_Base.BeginTrans
            Select Case Catalogo
                Case "Cat_Empleados":
                    If Trim(Txt_Cat_Empleados_Empleado_ID.Text) <> "" Then
'                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Empleados", "Empleado_ID", Trim(Txt_Cat_Empleados_Empleado_ID.Text)) = True Then
                         If Conectar_Ayudante.Elimina_Actualiza_Catalogo("Cat_Empleados", "Estatus", "E", "Empleado_ID", Trim(Txt_Cat_Empleados_Empleado_ID.Text)) Then
                         
                            If Grid_Cat_Empleados.Rows = 2 Then
                                Grid_Cat_Empleados.Rows = 0
                            Else
                                Grid_Cat_Empleados.RemoveItem Grid_Cat_Empleados.RowSel
                            End If
                            Call Conectar_Ayudante.Limpiar_Textos(Me)
                            MsgBox "Empleado eliminado", vbInformation + vbOKOnly, Me.Caption
                        Else
                            MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
                        End If '
                    Else
                        MsgBox "Seleccione un empleado para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
            End Select
            
            Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia los textos de la forma
        Conexion_Base.CommitTrans
    End If
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Eliminar_Curso_Click()
    If Grid_Cursos.Rows > 0 Then
        If Grid_Cursos.Rows = 2 Then
            Grid_Cursos.Rows = 0
        Else
            Grid_Cursos.RemoveItem Grid_Cursos.RowSel
        End If
    End If
End Sub

Private Sub Btn_Eliminar_Evaluaciones_Click()
    If Grid_Evaluaciones.Rows > 0 Then
        If Grid_Evaluaciones.Rows = 2 Then
            Grid_Evaluaciones.Rows = 0
        Else
            Grid_Evaluaciones.RemoveItem Grid_Evaluaciones.RowSel
        End If
    End If
End Sub

Private Sub Btn_Huella_Comedor_Click()
    If Trim(Txt_Cat_Empleados_No_Tarjeta.Text) <> "" Then
        Unload Frm_Adm_Enrollment
        Load Frm_Adm_Enrollment
        Frm_Adm_Enrollment.Caption = "REGISTRO DE HUELLAS"
        Frm_Adm_Enrollment.Txt_Cat_Empleados_Empleado_ID.Text = Txt_Cat_Empleados_Empleado_ID.Text
        Frm_Adm_Enrollment.Lbl_No_Empleado.Caption = Txt_Cat_Empleados_No_Tarjeta.Text
        Frm_Adm_Enrollment.Lbl_Nombre_Empleado.Caption = Txt_Cat_Empleados_Apellido_Paterno.Text & " " & Txt_Cat_Empleados_Apellido_Materno.Text & " " & Txt_Cat_Empleados_Nombre.Text
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), "ARCHIVO") = True Then
            Frm_Adm_Enrollment.Img_Cat_Empleados_Foto.picture = LoadPicture(PG_Ruta_Fotos & "\" & Trim(Txt_Cat_Empleados_Ruta_Imagen.Text))
        Else
            Frm_Adm_Enrollment.Img_Cat_Empleados_Foto.picture = LoadPicture("")
        End If
    End If
End Sub

Private Sub Btn_Imprimir_Click()
'Dim Ruta_Imagen As String
'Dim Ruta_Aplicacion  As String
'
'    Ruta_Aplicacion = App.Path
'    If Mid(Ruta_Aplicacion, Len(Ruta_Aplicacion), 1) = "\" Then
'        Ruta_Aplicacion = Mid(Ruta_Aplicacion, 1, Len(Ruta_Aplicacion) - 1)
'    End If
'    If Trim(Txt_Cat_Empleados_Empleado_ID.Text) <> "" Then
'        'Consulta la ruta de la imagen y su extension
'        Mi_SQL = "SELECT ISNULL(Imagen_Perfil,'') AS Imagen_Perfil FROM Cat_Empleados WHERE Empleado_ID = '" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
'        Ruta_Imagen = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Imagen_Perfil")
'        If Crea_PDF_Empleado_Expediente(Ruta_Temporal, Trim(Txt_Cat_Empleados_Empleado_ID.Text), Trim(Txt_Cat_Empleados_Empleado_ID.Text), Ruta_Aplicacion & "\Perfil\" & Ruta_Imagen) Then
'            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Temporal & "\" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & ".pdf", "ARCHIVO") = True Then
'                Call Mostrar_PDF(Trim(Txt_Cat_Empleados_Empleado_ID.Text) & ".pdf", Ruta_Temporal & "\" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & ".pdf")
'            End If
'        Else
'            MsgBox "No se puede crear el reporte del empleado", vbInformation + vbOKOnly, Me.Caption
'        End If
'    Else
'        MsgBox "Seleccione un Empleado", vbInformation + vbOKOnly, Me.Caption
'    End If
'    If Trim(Txt_Cat_Empleados_Empleado_ID.Text) <> "" Then
'        Call Barcode("39", Txt_Cat_Empleados_No_Tarjeta.Text, Printer, 15, 600, 1100, 2200)
'        Printer.EndDoc
'    End If
    If Trim(Txt_Cat_Empleados_No_Tarjeta.Text) <> "" Then
        Unload Frm_Cat_Empleados_Credencial
        Load Frm_Cat_Empleados_Credencial
        
        'Frm_Cat_Empleados_Credencial.Lbl_Nombre_Empleado.Caption = Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) & " " & Trim(Txt_Cat_Empleados_Apellido_Materno.Text) & " " & Trim(Txt_Cat_Empleados_Nombre.Text)
        Frm_Cat_Empleados_Credencial.Lbl_Apellido.Caption = Trim(Txt_Cat_Empleados_Apellido_Paterno.Text)
        Frm_Cat_Empleados_Credencial.Lbl_Nombre.Caption = Trim(Txt_Cat_Empleados_Nombre.Text)
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), "ARCHIVO") = True Then
            Frm_Cat_Empleados_Credencial.Img_Cat_Empleados_Foto.picture = LoadPicture(PG_Ruta_Fotos & "\" & Trim(Txt_Cat_Empleados_Ruta_Imagen.Text))
        Else
            Frm_Cat_Empleados_Credencial.Img_Cat_Empleados_Foto.picture = LoadPicture("")
        End If
        Frm_Cat_Empleados_Credencial.Ruta = App.Path & "\Logos_Empresas\" & Cmb_Cat_Empleados_Empresa.Text & "\" & Imagen_Logo
        Frm_Cat_Empleados_Credencial.Cargar_Credencial
    End If
End Sub

Public Function Barcode(CodeType As String, strCode As String, pic As Object, barscale As Integer, barHeight As Single, StartX As Single, StartY As Single)
Dim barWidth As Single
Dim i0 As Integer
Dim barStart As Single
Dim Nombre_Empleado As String
Dim Ultima_Posicion As Integer
Dim Aux_Espacio As Integer
Dim Espacio As Integer
Dim Cadena As String
Dim Cortada As String
Dim Contador_Renglon As Integer
Dim Salto_Linea As Double
    
    Select Case CodeType
        Case "39": strCode = UCase(strCode)
    End Select
    pic.Orientation = vbPRORLandscape
    'Se configura el tipo y el tamaño de letra que va a tener al momento de imprimir
    pic.ScaleMode = vbCentimeters
    pic.FontBold = True
    pic.FontSize = 8
    pic.Font = "MS Sans Serif"
    pic.FontSize = 8
    pic.Font = "COURIER NEW"
    pic.FontSize = 8
    pic.Font = "MS Sans Serif"
    pic.PaintPicture Img_Logo_Empresa.picture, 0.3, 0.3 'Con esto se imprime la imagen del logo
    pic.PaintPicture Img_Cat_Empleados_Foto.picture, 5.5, 0.3
    pic.FontSize = 10
    pic.CurrentX = 1.5
    pic.CurrentY = 3
    Nombre_Empleado = Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) & " " & Trim(Txt_Cat_Empleados_Apellido_Materno.Text) & " " & Trim(Txt_Cat_Empleados_Nombre.Text)
    Ultima_Posicion = 1
    Espacio = 1
    Aux_Espacio = 1
    Nombre_Empleado = Nombre_Empleado & Chr(13)
    Cadena = Mid(Nombre_Empleado, Ultima_Posicion, 20)
    Contador_Renglon = 3
    Salto_Linea = 0.5
    While Cadena <> ""
        Espacio = 0
        Aux_Espacio = 1
        While Aux_Espacio > 0
            Espacio = Aux_Espacio
            Aux_Espacio = InStr(Espacio + 1, Cadena, Chr(13), vbTextCompare)
            If Aux_Espacio = 0 Then
                Aux_Espacio = InStr(Espacio + 1, Cadena, " ", vbTextCompare)
            Else
                Espacio = Aux_Espacio + 1
                Aux_Espacio = 0
                Cadena = Mid(Cadena, 1, Espacio - 2)
            End If
        Wend
        If Espacio > 0 Then
            pic.Print Mid(Cadena, 1, Espacio)
            Contador_Renglon = Contador_Renglon + Salto_Linea
        End If
        Ultima_Posicion = Ultima_Posicion + Espacio
        Cadena = Mid(Nombre_Empleado, Ultima_Posicion, 20)
    Wend
    'pic.Print Nombre_Empleado
    barStart = StartX
    pic.ScaleMode = 1
    pic.ScaleMode = 1
    pic.FontBold = False
    pic.FontSize = 8: pic.CurrentX = StartX: pic.CurrentY = (StartY * 1) + barHeight: pic.Print strCode
End Function

Function Imprime_Varias_Lineas(Real As String, Tamaño As Integer, Contador_Renglon, Salto_Linea) As Double
Dim Ultima_Posicion As Integer
Dim Aux_Espacio As Integer
Dim Espacio As Integer
Dim Cadena As String
Dim Cortada As String

    Ultima_Posicion = 1
    Espacio = 1
    Aux_Espacio = 1
    Real = Real & Chr(13)
    Cadena = Mid(Real, Ultima_Posicion, Tamaño)
    While Cadena <> ""
        Espacio = 0
        Aux_Espacio = 1
        While Aux_Espacio > 0
            Espacio = Aux_Espacio
            Aux_Espacio = InStr(Espacio + 1, Cadena, Chr(13), vbTextCompare)
            If Aux_Espacio = 0 Then
                Aux_Espacio = InStr(Espacio + 1, Cadena, " ", vbTextCompare)
            Else
                Espacio = Aux_Espacio + 1
                Aux_Espacio = 0
                Cadena = Mid(Cadena, 1, Espacio - 2)
            End If
        Wend
        If Espacio > 0 Then
            'pic.Print Mid(Cadena, 1, Espacio)
            Contador_Renglon = Contador_Renglon + Salto_Linea
        End If
        Ultima_Posicion = Ultima_Posicion + Espacio
        Cadena = Mid(Real, Ultima_Posicion, Tamaño)
    Wend
    Imprime_Varias_Lineas = Contador_Renglon
End Function

Private Sub Btn_Imprimir_Contrato_Click()
    If Validar_Imprimir_Contrato Then
        Txt_Log.Visible = True
        Txt_Log.Text = Txt_Log.Text & "1. Iniciado Imprimir Contrato..." & vbCrLf
         Call Imprimir_Contrato
    Else
        MsgBox ("Seleccione un Empleado")
    End If
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
        Select Case Catalogo
            Case "Cat_Empleados"
                'Revisa que exista un registro a modificar y prepara la interfaz
                If Trim(Txt_Cat_Empleados_Empleado_ID.Text) <> "" Then
                    Fra_Cat_Empleados_Datos_Personales.Enabled = True
                    Fra_Cat_Empleados_Datos_Dependientes.Enabled = True
                    Fra_Cat_Empleados_Datos_Laborales.Enabled = True
                    Fra_Cat_Empleados_Baja.Enabled = True
                    Fra_Cursos.Enabled = True
                    Fra_Evaluaciones.Enabled = True
                    Fra_Cat_Empleados.Enabled = False
'                    Tab_Cat_Empleados.Tab = 0
                    Txt_Cat_Empleados_Nombre.SetFocus
                Else
                    MsgBox "Seleccione un empleado para poder modificar", vbOKOnly + vbInformation, Me.Caption
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
        Select Case Catalogo
            Case "Cat_Empleados":   'Modifica Empleados
                If Trim(Txt_Cat_Empleados_Nombre.Text) <> "" Then
                    If Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) <> "" Then
                        If Cmb_Cat_Empleados_Empresa.ListIndex > -1 Then
                            If Cmb_Cat_Empleados_Turno.ListIndex > -1 Then
                                If Cmb_Cat_Empleados_Departamento.ListIndex > -1 Then
                                    If Cmb_Cat_Empleados_Puesto.ListIndex > -1 Then
                                        If Trim(Txt_Cat_Empleados_No_Tarjeta.Text) <> "" Then
'                                            If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Empleados", "Nomipaq_ID", Trim(Txt_Cat_Empleados_Noi_ID.Text), "Empleado_ID", Trim(Txt_Cat_Empleados_Empleado_ID)) = False Then
                                                If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Empleados", "No_Tarjeta", Trim(Txt_Cat_Empleados_No_Tarjeta.Text), "Empleado_ID", Trim(Txt_Cat_Empleados_Empleado_ID)) = False Then
                                                    'Verfica si es cambio de estatus para solictar el motivo de baja
                                                    If Cmb_Cat_Empleados_Estatus.Text = "INACTIVO" Then
                                                        If Cmb_Cat_Empleados_Motivos_Baja.ListIndex = -1 Then
                                                            MsgBox "Ingrese el motivo de la baja", vbInformation + vbOKOnly, Me.Caption
                                                            Tab_Cat_Empleados.Tab = 3
                                                            Cmb_Cat_Empleados_Motivos_Baja.SetFocus
                                                            Exit Sub
                                                        End If
                                                        If Txt_Cat_Empleados_Observaciones_Baja.Text = "" Then
                                                            MsgBox "Ingrese algun comentario de la baja", vbInformation + vbOKOnly, Me.Caption
                                                            Tab_Cat_Empleados.Tab = 3
                                                            Txt_Cat_Empleados_Observaciones_Baja.SetFocus
                                                            Exit Sub
                                                        End If
                                                    End If
                                                    Modifica_Cat_Empleados
                                            Else
                                                MsgBox "El No. de Nómina ya se ha registrado", vbOKOnly + vbInformation, Me.Caption
                                                Tab_Cat_Empleados.Tab = 2
                                                Txt_Cat_Empleados_No_Tarjeta.SetFocus
                                            End If
'                                            Else
'                                                MsgBox "El identificador Nomipaq ya se ha registrado", vbOKOnly + vbInformation, Me.Caption
'                                                Tab_Cat_Empleados.Tab = 2
'                                                Txt_Cat_Empleados_Noi_ID.SetFocus
'                                            End If
                                        Else
                                            MsgBox "Ingrese el No de Nómina del empleado", vbOKOnly + vbInformation, Me.Caption
                                            Tab_Cat_Empleados.Tab = 2
                                            Txt_Cat_Empleados_No_Tarjeta.SetFocus
                                        End If
                                    Else
                                        MsgBox "Ingrese el Puestp del empleado", vbOKOnly + vbInformation, Me.Caption
                                    End If
                                Else
                                    MsgBox "Ingrese el Departamento del empleado", vbOKOnly + vbInformation, Me.Caption
                                    Tab_Cat_Empleados.Tab = 2
                                    Cmb_Cat_Empleados_Departamento.SetFocus
                                End If
                            Else
                                MsgBox "Seleccione el turno", vbOKOnly + vbInformation, Me.Caption
                                Tab_Cat_Empleados.Tab = 2
                                Cmb_Cat_Empleados_Turno.SetFocus
                            End If
                        Else
                            MsgBox "Seleccione la empresa", vbOKOnly + vbInformation, Me.Caption
                            Tab_Cat_Empleados.Tab = 2
                            Cmb_Cat_Empleados_Empresa.SetFocus
                        End If
                    Else
                        MsgBox "Ingrese el Apellido Paterno del empleado", vbOKOnly + vbInformation, Me.Caption
                        Tab_Cat_Empleados.Tab = 0
                        Txt_Cat_Empleados_Apellido_Paterno.SetFocus
                    End If
                Else
                    MsgBox "Ingrese el Nombre del empleado", vbOKOnly + vbInformation, Me.Caption
                    Tab_Cat_Empleados.Tab = 0
                    Txt_Cat_Empleados_Nombre.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub Btn_Nuevo_Click()
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Imprimir.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        'Muestra el picture del catalogo seleccionado
        Select Case Catalogo
            Case "Cat_Empleados": 'Catalogo de Empleados, Prepara la interfaz
                Txt_Cat_Empleados_Empleado_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empleados", "Empleado_ID"), "00000")
                Dtp_Cat_Empleados_Fecha_Ingreso.Value = Now
                Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Now
                Dtp_Cat_Empleados_Fecha_Baja.Value = Now
                Fra_Cat_Empleados_Datos_Personales.Enabled = True
                Fra_Cat_Empleados_Datos_Dependientes.Enabled = True
                Fra_Cat_Empleados_Datos_Laborales.Enabled = True
                Fra_Cat_Empleados_Baja.Enabled = True
                Fra_Cursos.Enabled = True
                Fra_Evaluaciones.Enabled = True
'                Fra_Cat_Empleados.Enabled = False
                Cmb_Cat_Empleados_Estatus.ListIndex = 0
                Cmb_Cat_Empleados_Estatus.Enabled = False
                Tab_Cat_Empleados.Tab = 0
                Chk_Cat_Empleados_Trabaja_Domingos.Value = 0
                Chk_Cat_Empleados_Infonavit.Value = 0
                Txt_Cat_Empleados_Nombre.SetFocus
        End Select
    Else
        Select Case Catalogo
            Case "Cat_Empleados":   'Modifica Empleados
                If Trim(Txt_Cat_Empleados_Nombre.Text) <> "" Then
                    If Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) <> "" Then
                        If Cmb_Cat_Empleados_Empresa.ListIndex > -1 Then
                            If Cmb_Cat_Empleados_Turno.ListIndex > -1 Then
                                If Cmb_Cat_Empleados_Departamento.ListIndex > -1 Then
                                    If Cmb_Cat_Empleados_Puesto.ListIndex > -1 Then
                                        If Trim(Txt_Cat_Empleados_No_Tarjeta.Text) <> "" Then
                                            If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Empleados", "No_Tarjeta", Trim(Txt_Cat_Empleados_No_Tarjeta.Text), "Empleado_ID", Trim(Txt_Cat_Empleados_Empleado_ID)) = False Then
                                                'Verfica si es cambio de estatus para solictar el motivo de baja
                                                If Cmb_Cat_Empleados_Estatus.Text = "INACTIVO" Then
                                                    If Cmb_Cat_Empleados_Motivos_Baja.ListIndex = -1 Then
                                                        MsgBox "Ingrese el motivo de la baja", vbInformation + vbOKOnly, Me.Caption
                                                        Tab_Cat_Empleados.Tab = 3
                                                        Cmb_Cat_Empleados_Motivos_Baja.SetFocus
                                                        Exit Sub
                                                    End If
                                                    If Txt_Cat_Empleados_Observaciones_Baja.Text = "" Then
                                                        MsgBox "Ingrese algun comentario de la baja", vbInformation + vbOKOnly, Me.Caption
                                                        Tab_Cat_Empleados.Tab = 3
                                                        Txt_Cat_Empleados_Observaciones_Baja.SetFocus
                                                        Exit Sub
                                                    End If
                                                End If
                                                Alta_Cat_Empleados
                                            Else
                                                MsgBox "El No. de Nómina ya se ha registrado", vbOKOnly + vbInformation, Me.Caption
                                                Tab_Cat_Empleados.Tab = 2
                                                Txt_Cat_Empleados_No_Tarjeta.SetFocus
                                            End If
                                        Else
                                            MsgBox "Ingrese el No de Nómina del empleado", vbOKOnly + vbInformation, Me.Caption
                                            Tab_Cat_Empleados.Tab = 2
                                            Txt_Cat_Empleados_No_Tarjeta.SetFocus
                                        End If
                                    Else
                                        MsgBox "Ingrese el Puesto del empleado", vbOKOnly + vbInformation, Me.Caption
                                    End If
                                Else
                                    MsgBox "Ingrese el Departamento del empleado", vbOKOnly + vbInformation, Me.Caption
                                    Tab_Cat_Empleados.Tab = 2
                                    Cmb_Cat_Empleados_Departamento.SetFocus
                                End If
                            Else
                                MsgBox "Seleccione el turno", vbOKOnly + vbInformation, Me.Caption
                                Tab_Cat_Empleados.Tab = 2
                                Cmb_Cat_Empleados_Turno.SetFocus
                            End If
                        Else
                            MsgBox "Seleccione la empresa", vbOKOnly + vbInformation, Me.Caption
                            Tab_Cat_Empleados.Tab = 2
                            Cmb_Cat_Empleados_Empresa.SetFocus
                        End If
                    Else
                        MsgBox "Ingrese el Apellido Paterno del empleado", vbOKOnly + vbInformation, Me.Caption
                        Tab_Cat_Empleados.Tab = 0
                        Txt_Cat_Empleados_Apellido_Paterno.SetFocus
                    End If
                Else
                    MsgBox "Ingrese el Nombre del empleado", vbOKOnly + vbInformation, Me.Caption
                    Tab_Cat_Empleados.Tab = 0
                    Txt_Cat_Empleados_Nombre.SetFocus
                End If
        End Select
    End If
End Sub

'Cierra la forma
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
        Select Case Catalogo
            Case "Cat_Empleados"
                Fra_Cat_Empleados_Datos_Personales.Enabled = False
                Fra_Cat_Empleados_Datos_Dependientes.Enabled = False
                Fra_Cat_Empleados_Datos_Laborales.Enabled = False
                Fra_Cat_Empleados_Baja.Enabled = False
                Fra_Cursos.Enabled = False
                Fra_Evaluaciones.Enabled = True
                Fra_Cat_Empleados.Enabled = True
                Cmb_Cat_Empleados_Estatus.Enabled = True
                Grid_Cat_Empleados_Dependientes.Rows = 0
                Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value = Now
                Dtp_Cat_Empleados_Fecha_Baja.Value = Now
                Dtp_Cat_Empleados_Fecha_Ingreso.Value = Now
                Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Now
                Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = Now
                Cmb_Cat_Empleados_Contratacion.ListIndex = -1
                Cmb_Cat_Empleados_Estado_Civil.ListIndex = -1
                Cmb_Cat_Empleados_Parentesco.ListIndex = -1
                Cmb_Cat_Empleados_Tipo.ListIndex = -1
                Cmb_Cat_Empleados_Sexo.ListIndex = -1
                Cmb_Cat_Empleados_Tipo_Empleado.ListIndex = -1
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Empleados", Me)
        End Select
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Contratacion_Change()
    Dtp_Cat_Empleados_Fecha_Termino_Contrato.Visible = False
    Lbl_Cat_Empleados_Fecha_Termino_Contrato.Visible = False
    If Cmb_Cat_Empleados_Contratacion.Text = "EVENTUAL" Then
        Dtp_Cat_Empleados_Fecha_Termino_Contrato.Visible = True
        Lbl_Cat_Empleados_Fecha_Termino_Contrato.Visible = True
        Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = DateAdd("M", 6, Dtp_Cat_Empleados_Fecha_Ingreso.Value)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Contratacion_Click()
    Cmb_Cat_Empleados_Contratacion_Change
End Sub

Private Sub Cmb_Cat_Empleados_Departamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Cat_Empleados_Departamento, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Departamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Cat_Empleados_Departamento, KeyCode)
End Sub

Private Sub Cmb_Cat_Empleados_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Cat_Empleados_Empresa, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Cat_Empleados_Empresa, KeyCode)
End Sub

Private Sub Cmb_Cat_Empleados_Supervisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Cat_Empleados_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Cat_Empleados_Supervisor.Text) & "%' OR " & "Apellido_Paterno LIKE '%" & Trim(Cmb_Cat_Empleados_Supervisor.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Cat_Empleados_Supervisor.Text) & "%')", False, "")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub



Private Sub Cmb_Cat_Empleados_Turno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Cat_Empleados_Turno, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Turno_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Cat_Empleados_Turno, KeyCode)
End Sub

Private Sub Cmb_Curso_Click()
Dim Rs_Curso As rdoResultset
    If Cmb_Curso.ListIndex > -1 Then
        Mi_SQL = "SELECT Curso_ID,ISNULL(Instructor,'') AS Instructor,ISNULL(Horas,0) AS Horas"
        Mi_SQL = Mi_SQL & " FROM Cat_Cursos WHERE Curso_ID='" & Format(Cmb_Curso.ItemData(Cmb_Curso.ListIndex), "00000") & "'"
        Set Rs_Curso = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Curso.EOF Then
            Txt_Comentarios_Curso.Text = Rs_Curso.rdoColumns("Instructor")
            Txt_Horas_Curso.Text = Rs_Curso.rdoColumns("Horas") & " hrs"
        End If
        Rs_Curso.Close
    End If
End Sub

Private Sub Cmb_Curso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID,Nombre", "Cat_Cursos", Cmb_Curso, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Gap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Gap_ID,Nombre", "Cat_Gaps", Cmb_Gap, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Transporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Transporte_ID,Nombre", "Cat_Transportes", Cmb_Transporte, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Ingreso_Change()
    Txt_Cat_Empleados_Antiguedad.Text = Calcula_Edad(Dtp_Cat_Empleados_Fecha_Ingreso.Value)
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Ingreso_Click()
    Dtp_Cat_Empleados_Fecha_Ingreso_Change
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Ingreso_LostFocus()
    Dtp_Cat_Empleados_Fecha_Ingreso_Change
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Nacimiento_Change()
    Dim anios As Double
    Txt_Cat_Empleados_Edad.ForeColor = vbBlack
    Txt_Cat_Empleados_Edad.Text = Calcula_Edad(Dtp_Cat_Empleados_Fecha_Nacimiento.Value, anios)
    If anios < Edad_Minima_Contratacion Then
        Txt_Cat_Empleados_Edad.ForeColor = vbRed
    End If
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Nacimiento_Click()
    Dtp_Cat_Empleados_Fecha_Nacimiento_Change
End Sub

Private Sub Dtp_Cat_Empleados_Fecha_Nacimiento_LostFocus()
    Dtp_Cat_Empleados_Fecha_Nacimiento_Change
End Sub

Private Sub Dtp_Fecha_Evaluacion_Change()
    Dtp_Proxima_Evaluacion.Value = DateAdd("M", 6, Dtp_Fecha_Evaluacion.Value)
End Sub

Private Sub Form_Load()
    Me.Height = 7665
    Me.Width = 8505
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
    'Tab_Cat_Empleados.TabEnabled(4) = False
End Sub

Private Sub Grid_Cat_Empleados_Click()
Dim Rs_Consulta_Cat_Empleados As rdoResultset    'Informacion de la empresa
Dim Rs_Consulta_Cat_Empleados_Dependientes As rdoResultset
Dim Edad As String
Dim anios As Double
Dim Meses As Double
Dim anios_trabajo As String
Dim Edad_Dependiente As String
    
    If Grid_Cat_Empleados.Rows > 1 Then
        With Grid_Cat_Empleados
            'Consulta los empleados
            Mi_SQL = "SELECT * FROM Cat_Empleados, Cat_Empresas"
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(.TextMatrix(.RowSel, 0)) & "'"
            Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID = Cat_Empresas.Empresa_ID "
            Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Cat_Empleados.EOF Then
                With Rs_Consulta_Cat_Empleados
                    If Not IsNull(.rdoColumns("LOGO")) Then
                        Imagen_Logo = Trim(.rdoColumns("LOGO"))
                    End If
                    Txt_Cat_Empleados_Empleado_ID.Text = Trim(.rdoColumns("Empleado_ID"))
                    If .rdoColumns("Empresa_ID") <> "" Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Empresa_ID"), Cmb_Cat_Empleados_Empresa)
                    End If
                    If Not IsNull(.rdoColumns("Supervisor_ID")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Cat_Empleados_Supervisor)
                    Else
                        Cmb_Cat_Empleados_Supervisor.ListIndex = -1
                    End If
                    If .rdoColumns("Departamento_ID") <> "" Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Cat_Empleados_Departamento)
                    End If
                    If .rdoColumns("Puesto_ID") <> "" Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Puesto_ID"), Cmb_Cat_Empleados_Puesto)
                    End If
                    If .rdoColumns("Turno_ID") <> "" Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Turno_ID"), Cmb_Cat_Empleados_Turno)
                    End If
                    If Not IsNull(.rdoColumns("Nivel_Academico_ID")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Nivel_Academico_ID"), Cmb_Cat_Empleados_Nivel_Estudio)
                    End If
                    If Not IsNull(.rdoColumns("Transporte_ID")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Transporte_ID"), Cmb_Transporte)
                    Else
                        Cmb_Transporte.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Gap_ID")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Gap_ID"), Cmb_Gap)
                    Else
                        Cmb_Gap.Text = ""
                    End If
                    'Datos Personales
                    If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Cat_Empleados_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Cat_Empleados_Estatus, 1)
                    End If
                    Cmb_Cat_Empleados_Tipo.ListIndex = 0
                    If Not IsNull(.rdoColumns("Tipo")) Then
                        Cmb_Cat_Empleados_Tipo.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Tipo"), Cmb_Cat_Empleados_Tipo, 1)
                    End If
                    Txt_Cat_Empleados_Nombre.Text = Trim(UCase(.rdoColumns("Nombre")))
                    Txt_Cat_Empleados_Apellido_Paterno.Text = Trim(UCase(.rdoColumns("Apellido_Paterno")))
                    Txt_Cat_Empleados_Apellido_Materno.Text = Trim(UCase(.rdoColumns("Apellido_Materno")))
                    Txt_Cat_Empleados_Lugar_Nacimiento.Text = Trim(UCase(.rdoColumns("Lugar_Nacimiento")))
                    If Not IsNull(.rdoColumns("Sexo")) Then
                    Cmb_Cat_Empleados_Sexo.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Sexo"), Cmb_Cat_Empleados_Sexo)
                    End If
                    Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Format(.rdoColumns("Fecha_Nacimiento"), "MM/dd/yyyy")
                    If Not IsNull(.rdoColumns("Email")) Then
                        Txt_Cat_Empleados_Email.Text = Trim(.rdoColumns("Email"))
                    End If
                    'Direccion
                    Txt_Cat_Empleados_Direccion.Text = .rdoColumns("Direccion")
                    Txt_Cat_Empleados_Colonia.Text = .rdoColumns("Colonia")
                    Txt_Cat_Empleados_CP.Text = .rdoColumns("Codigo_Postal")
                    Txt_Cat_Empleados_Ciudad.Text = .rdoColumns("Ciudad")
                    Txt_Cat_Empleados_Estado.Text = .rdoColumns("Estado")
                    Edad = Calcula_Edad(.rdoColumns("Fecha_Nacimiento"))
                    Txt_Cat_Empleados_Edad.Text = Edad
                    Txt_Cat_Empleados_Clave_Elector.Text = Trim(UCase(.rdoColumns("Clave_Elector")))
                    Cmb_Cat_Empleados_Estado_Civil.ListIndex = -1
                    If Not IsNull(.rdoColumns("Estado_Civil")) Then
                        Cmb_Cat_Empleados_Estado_Civil.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(UCase(.rdoColumns("Estado_Civil"))), Cmb_Cat_Empleados_Estado_Civil)
                    End If
                    Txt_Cat_Empleados_RFC.Text = Trim(.rdoColumns("RFC"))
                    Txt_Cat_Empleados_Curp.Text = Trim(.rdoColumns("Curp"))
                    Txt_Cat_Empleados_NSS.Text = Trim(.rdoColumns("Nss"))
                    'Foto
                    Txt_Cat_Empleados_Ruta_Imagen.Text = ""
                    Img_Cat_Empleados_Foto.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que elarchivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Cat_Empleados_Foto.picture = LoadPicture(PG_Ruta_Fotos & "\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                                Img_Cat_Empleados_Foto.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                            End If
                        End If
                        Txt_Cat_Empleados_Ruta_Imagen.Text = .rdoColumns("Imagen_Perfil")
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Cat_Empleados_Foto.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                    If Not IsNull(.rdoColumns("Nomipaq_ID")) Then
                        Txt_Cat_Empleados_Seccion.Text = Trim(.rdoColumns("Nomipaq_ID"))
                    Else
                        Txt_Cat_Empleados_Seccion.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Clave_SAP")) Then
                        Txt_Clave_SAP.Text = Trim(.rdoColumns("Clave_SAP"))
                    Else
                        Txt_Clave_SAP.Text = ""
                    End If
                    Txt_Cat_Empleados_No_Tarjeta.Text = Trim(.rdoColumns("No_Tarjeta"))
                    Dtp_Cat_Empleados_Fecha_Ingreso.Value = Format(.rdoColumns("Fecha_Ingreso"), "MM/dd/yyyy")
                    anios_trabajo = Calcula_Edad(.rdoColumns("Fecha_Ingreso"))
                    Txt_Cat_Empleados_Antiguedad.Text = anios_trabajo
                    
                    If Not IsNull(.rdoColumns("Tipo_Empleado")) Then
                        Cmb_Cat_Empleados_Tipo_Empleado.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Tipo_Empleado"), Cmb_Cat_Empleados_Tipo_Empleado)
                    End If
                    
                    If Not IsNull(.rdoColumns("Tipo_Contratacion")) Then
                        Cmb_Cat_Empleados_Contratacion.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Tipo_Contratacion"), Cmb_Cat_Empleados_Contratacion)
                    End If
                    
                    Lbl_Cat_Empleados_Fecha_Termino_Contrato.Visible = False
                    Dtp_Cat_Empleados_Fecha_Termino_Contrato.Visible = False
                    If Cmb_Cat_Empleados_Contratacion.ListIndex > -1 Then
                        If Cmb_Cat_Empleados_Contratacion.Text = "EVENTUAL" Then
                            Lbl_Cat_Empleados_Fecha_Termino_Contrato.Visible = True
                            Dtp_Cat_Empleados_Fecha_Termino_Contrato.Visible = True
                            If Not IsNull(.rdoColumns("Fecha_Termino_Contrato")) Then
                                Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = .rdoColumns("Fecha_Termino_Contrato")
                            End If
                        End If
                    End If
                    Txt_Cat_Empleados_Salario_Diario.Text = Val(.rdoColumns("Salario_Diario"))
                    If Not IsNull(.rdoColumns("Salario_Diario_Variable")) Then
                        Txt_Cat_Empleados_Vacaciones.Text = Val(.rdoColumns("Salario_Diario_Variable"))
                    Else
                        Txt_Cat_Empleados_Vacaciones.Text = 0
                    End If
                    If .rdoColumns("Trabaja_Domingos") = "S" Then
                        Chk_Cat_Empleados_Trabaja_Domingos.Value = 1
                    Else
                        Chk_Cat_Empleados_Trabaja_Domingos.Value = 0
                    End If
                    If .rdoColumns("Infonavit") = "S" Then
                        Chk_Cat_Empleados_Infonavit.Value = 1
                    Else
                        Chk_Cat_Empleados_Infonavit.Value = 0
                    End If
                    If Not IsNull(.rdoColumns("Cedula_Identidad_Ciudadana")) Then
                        If Trim(.rdoColumns("Cedula_Identidad_Ciudadana")) <> "" Then
                            Cmb_Subdivision.Text = .rdoColumns("Cedula_Identidad_Ciudadana")
                        Else
                            Cmb_Subdivision.ListIndex = -1
                        End If
                    Else
                        Cmb_Subdivision.ListIndex = -1
                    End If
                    If Not IsNull(.rdoColumns("En_Caso_Emergencia")) Then
                        Txt_Cat_Empleados_LLamar_A.Text = Trim(.rdoColumns("En_Caso_Emergencia"))
                    End If
                    If Not IsNull(.rdoColumns("Telefono_Emergencia1")) Then
                        Txt_Cat_Empleados_Llamar_Telefono1.Text = Trim(.rdoColumns("Telefono_Emergencia1"))
                    End If
                    If Not IsNull(.rdoColumns("Telefono_Emergencia2")) Then
                        Txt_Cat_Empleados_Llamar_Telefono2.Text = Trim(.rdoColumns("Telefono_Emergencia2"))
                    End If
                    If Not IsNull(.rdoColumns("Alergia1")) Then
                        Txt_Cat_Empleados_Alergias1.Text = Trim(.rdoColumns("Alergia1"))
                    End If
                    If Not IsNull(.rdoColumns("Alergia2")) Then
                        Txt_Cat_Empleados_Alergias2.Text = Trim(.rdoColumns("Alergia2"))
                    End If
                    If Not IsNull(.rdoColumns("Alergia3")) Then
                        Txt_Cat_Empleados_Alergias3.Text = Trim(.rdoColumns("Alergia3"))
                    End If
                    If Not IsNull(.rdoColumns("Fecha_Baja")) Then Dtp_Cat_Empleados_Fecha_Baja.Value = Format(.rdoColumns("Fecha_Baja"), "MM/dd/yyyy")
                    If Not IsNull(.rdoColumns("Motivo_Baja_ID")) Then
                        Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Motivo_Baja_ID"), Cmb_Cat_Empleados_Motivos_Baja)
                    Else
                        Cmb_Cat_Empleados_Motivos_Baja.ListIndex = -1
                    End If
                    If Not IsNull(.rdoColumns("Comentarios_Baja")) Then
                         Txt_Cat_Empleados_Observaciones_Baja.Text = Trim(.rdoColumns("Comentarios_Baja"))
                    End If
                    'Adicionales
                    If Not IsNull(.rdoColumns("Campo_1")) Then
                        Txt_Campo_1.Text = Trim(.rdoColumns("Campo_1"))
                    Else
                        Txt_Campo_1.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Campo_2")) Then
                        Txt_Campo_2.Text = Trim(.rdoColumns("Campo_2"))
                    Else
                        Txt_Campo_2.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Campo_3")) Then
                        Txt_Campo_3.Text = Trim(.rdoColumns("Campo_3"))
                    Else
                        Txt_Campo_3.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Campo_4")) Then
                        Txt_Campo_4.Text = Trim(.rdoColumns("Campo_4"))
                    Else
                        Txt_Campo_4.Text = ""
                    End If
                    If Not IsNull(.rdoColumns("Campo_5")) Then
                        Txt_Campo_5.Text = Trim(.rdoColumns("Campo_5"))
                    Else
                        Txt_Campo_5.Text = ""
                    End If
                    'Consulta los dependientes del empleado
                    Grid_Cat_Empleados_Dependientes.Rows = 0
                    Mi_SQL = "SELECT * FROM Cat_Empleados_Parentesco WHERE Empleado_ID = '" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
                    Set Rs_Consulta_Cat_Empleados_Dependientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empleados_Dependientes.EOF Then
                        'Agrega el dependiente a la lista
                        Grid_Cat_Empleados_Dependientes.Cols = 4
                        If Grid_Cat_Empleados_Dependientes.Rows = 0 Then
                            Grid_Cat_Empleados_Dependientes.AddItem "Parentesco" & Chr(9) & "Nombre" _
                                & Chr(9) & "F. Nacimiento" & Chr(9) & "Edad"
                            Grid_Cat_Empleados_Dependientes.ColWidth(0) = 1000  'Parentesco
                            Grid_Cat_Empleados_Dependientes.ColWidth(1) = 2500  'Nombre
                            Grid_Cat_Empleados_Dependientes.ColWidth(2) = 1300  'F. Nacimiento
                            Grid_Cat_Empleados_Dependientes.ColWidth(3) = 1800  'Edad
                        End If
                        While Not Rs_Consulta_Cat_Empleados_Dependientes.EOF
                            Edad_Dependiente = ""
                            Edad_Dependiente = Calcula_Edad(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Fecha_Nacimiento"))
                            Grid_Cat_Empleados_Dependientes.AddItem Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Parentesco") & Chr(9) & Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Nombre") _
                                & Chr(9) & Format(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Fecha_Nacimiento"), "dd/MMM/yyyy") & Chr(9) & Edad_Dependiente
                            Rs_Consulta_Cat_Empleados_Dependientes.MoveNext
                        Wend
                        Grid_Cat_Empleados_Dependientes.FixedRows = 1
                    End If
                    Rs_Consulta_Cat_Empleados_Dependientes.Close
                    'Consulta los cursos del empleado
                    Grid_Cursos.Rows = 0
                    Mi_SQL = "SELECT Cat_Cursos_Detalles.*,Cat_Cursos.Nombre AS Curso"
                    Mi_SQL = Mi_SQL & " FROM Cat_Cursos_Detalles,Cat_Cursos"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Cursos_Detalles.Curso_ID=Cat_Cursos.Curso_ID"
                    Mi_SQL = Mi_SQL & " AND Cat_Cursos_Detalles.Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
                    Set Rs_Consulta_Cat_Empleados_Dependientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empleados_Dependientes.EOF Then
                        Grid_Cursos.Cols = 6
                        If Grid_Cursos.Rows = 0 Then
                            Grid_Cursos.AddItem "Curso_ID" & Chr(9) & "Curso" _
                                & Chr(9) & "Comentarios" & Chr(9) & "Estatus" _
                                & Chr(9) & "Inicio" & Chr(9) & "Fin"
                            Grid_Cursos.ColWidth(0) = 0     'Curso_ID
                            Grid_Cursos.ColWidth(1) = 2500  'Curso
                            Grid_Cursos.ColWidth(2) = 1400  'Comentarios
                            Grid_Cursos.ColWidth(3) = 1200  'Estatus
                            Grid_Cursos.ColWidth(4) = 1200  'Inicio
                            Grid_Cursos.ColWidth(5) = 1200  'Fin
                        End If
                        While Not Rs_Consulta_Cat_Empleados_Dependientes.EOF
                            Grid_Cursos.AddItem Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Curso_ID") _
                                & Chr(9) & Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Curso") _
                                & Chr(9) & Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Comentarios") _
                                & Chr(9) & Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Estatus") _
                                & Chr(9) & Format(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy") _
                                & Chr(9) & Format(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Fecha_Fin"), "dd/MMM/yyyy")
                            Rs_Consulta_Cat_Empleados_Dependientes.MoveNext
                        Wend
                        Grid_Cursos.FixedRows = 1
                    End If
                    Rs_Consulta_Cat_Empleados_Dependientes.Close
                    'Consulta las evaluaciones del empleado
                    Grid_Evaluaciones.Rows = 0
                    Mi_SQL = "SELECT * FROM Cat_Empleados_Evaluaciones"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
                    Set Rs_Consulta_Cat_Empleados_Dependientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empleados_Dependientes.EOF Then
                        Grid_Evaluaciones.AddItem "Evaluacion" & Chr(9) & "Fecha" & Chr(9) & "Siguiente"
                        Grid_Evaluaciones.ColWidth(0) = 4900  'Evaluacion
                        Grid_Evaluaciones.ColWidth(1) = 1300  'Inicio
                        Grid_Evaluaciones.ColWidth(2) = 1300  'Fin
                        While Not Rs_Consulta_Cat_Empleados_Dependientes.EOF
                            Grid_Evaluaciones.AddItem Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Evaluacion") _
                                & Chr(9) & Format(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                                & Chr(9) & Format(Rs_Consulta_Cat_Empleados_Dependientes.rdoColumns("Proxima_Evaluacion"), "dd/MMM/yyyy")
                            Rs_Consulta_Cat_Empleados_Dependientes.MoveNext
                        Wend
                        Grid_Evaluaciones.FixedRows = 1
                    End If
                    Rs_Consulta_Cat_Empleados_Dependientes.Close
                End With
            End If
            Rs_Consulta_Cat_Empleados.Close
        End With
    End If
End Sub

Private Sub Grid_Cat_Empleados_EnterCell()
    Grid_Cat_Empleados_Click
End Sub

Private Sub Img_Cat_Empleados_Foto_DblClick()
    Btn_Cat_Empleados_Foto_Click
End Sub

Private Sub Txt_Campo_1_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Campo_2_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Campo_3_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Campo_4_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Campo_5_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Cat_Empleados_Alergias1_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Alergias2_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Alergias3_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Apellido_Materno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Apellido_Paterno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Ciudad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Cat_Empleados_Clave_Elector_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Colonia_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Cat_Empleados_Curp_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Dependiente_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Direccion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub
Private Sub Txt_Cat_Empleados_Estado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub
Private Sub Txt_Cat_Empleados_LLamar_A_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_Llamar_Telefono1_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_Llamar_Telefono2_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_Lugar_Nacimiento_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_No_Tarjeta_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_NSS_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empleados_Observaciones_Baja_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub
Private Sub Txt_Cat_Empleados_RFC_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empleados_Salario_Diario_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cat_Empleados_Salario_Diario.Text, True)
End Sub

Private Sub Txt_Cat_Empleados_Seccion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Cat_Empleados_Vacaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cat_Empleados_Vacaciones.Text, False)
End Sub

Private Sub Txt_Clave_SAP_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comentarios_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

'************************************************Inicio Empleados***************************************
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Empleados
'DESCRIPCION: Consulta los empleados y los muestra en el grid
'PARAMETROS : Nombre- Indica el parámetro de búsqueda
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 05-Abr-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Cat_Empleados(Nombre As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros
    Grid_Cat_Empleados.Cols = 5
    Grid_Cat_Empleados.Rows = 0
    Grid_Cat_Empleados.AddItem "Empleado ID" & Chr(9) & "Nomina" & Chr(9) & "Apellidos" & Chr(9) & "Nombre" & Chr(9) & "Estatus"
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Empleado_ID,No_Tarjeta,Nombre,Apellido_Paterno+' '+Apellido_Materno AS Apellidos,RFC,Estatus"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
    If IsNumeric(Nombre) Then
        Mi_SQL = Mi_SQL & " WHERE No_Tarjeta=" & Nombre
    Else
        Mi_SQL = Mi_SQL & " WHERE (Apellido_Paterno LIKE '%" & Nombre & "%'"
        Mi_SQL = Mi_SQL & " OR Apellido_Materno LIKE '%" & Nombre & "%'"
        Mi_SQL = Mi_SQL & " OR Nombre LIKE '%" & Nombre & "%'"
        Mi_SQL = Mi_SQL & " OR RFC LIKE '%" & Nombre & "%'"
        Mi_SQL = Mi_SQL & " OR NSS LIKE '%" & Nombre & "%')"
'        Mi_SQL = Mi_SQL & " AND Estatus <> 'E' "
    End If
    
'    If Empleado_Supervisor_ID <> "" Then    'Filtra el supervisor
'        Mi_SQL = Mi_SQL & " AND (Empleado_ID='" & Empleado_Supervisor_ID & "'"
'        Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Empleado_Supervisor_ID & "')"
'    Else
'    If Trim(Nombre) = "" _
'    And Rol_ID <> "00001" Then
'        Mi_SQL = Mi_SQL & " AND (Empleado_ID='" & Empleado_Supervisor_ID & "'"
'        Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Empleado_Supervisor_ID & "')"
'    Else
'        If Rol_ID <> "00001" Then
'            Mi_SQL = Mi_SQL & " AND Supervisor_ID IS NULL"
'            Mi_SQL = Mi_SQL & " AND Tipo <> 'S'"
'        End If
'    End If
'    End If
    
    Mi_SQL = Mi_SQL & " ORDER BY No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Empleados
        While Not .EOF
            Grid_Cat_Empleados.AddItem .rdoColumns("Empleado_ID") _
                & Chr(9) & .rdoColumns("No_Tarjeta") _
                & Chr(9) & .rdoColumns("Apellidos") _
                & Chr(9) & .rdoColumns("Nombre") _
                & Chr(9) & .rdoColumns("Estatus")
                
            Grid_Cat_Empleados.FixedRows = 1
            .MoveNext
        Wend
        
        'Configura el tamaño de las columnas del grid_usuarios
        Grid_Cat_Empleados.ColWidth(0) = 0      'Empleado_ID
        Grid_Cat_Empleados.ColWidth(1) = 1000   'No. Nomina
        Grid_Cat_Empleados.ColWidth(2) = 3000   'Apellidos
        Grid_Cat_Empleados.ColWidth(3) = 3000   'Nombre(s)
        Grid_Cat_Empleados.ColWidth(4) = 700    'Estatus
    
    End With
    Rs_Consulta_Cat_Empleados.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Cat_Empleados
'DESCRIPCION: Da de alta un registro en Cat_Empleados
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 15-Mayo-2009
'MODIFICO   : Sergio Ulises Durán Hernández
'FECHA_MODIFICO: 12-Marzo-2012
'CAUSA_MODIFICO: Adecuación al sistema de SRG
'*******************************************************************************
Private Sub Alta_Cat_Empleados()
Dim Rs_Alta_Cat_Empleados As rdoResultset 'Informacion del registro
Dim Rs_Alta_Cat_Empleados_Parentesco As rdoResultset

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Set Rs_Alta_Cat_Empleados = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados")
    'Agrega el reigstro del empleado
    With Rs_Alta_Cat_Empleados
        .AddNew
            'ID's
            Txt_Cat_Empleados_Empleado_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empleados", "Empleado_ID"), "00000")
            .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
            .rdoColumns("Empresa_ID") = Format(Cmb_Cat_Empleados_Empresa.ItemData(Cmb_Cat_Empleados_Empresa.ListIndex), "00000")
            If Cmb_Cat_Empleados_Supervisor.ListIndex > -1 Then
                .rdoColumns("Supervisor_ID") = Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000")
                'Busca la sección del supervisor
                Mi_SQL = "SELECT * FROM Cat_Secciones"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000") & "'"
                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                    .rdoColumns("Nomipaq_ID") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                    Txt_Cat_Empleados_Seccion.Text = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                Else
                    .rdoColumns("Nomipaq_ID") = Trim(Txt_Cat_Empleados_Seccion.Text)
                End If
                Rs_Alta_Cat_Empleados_Parentesco.Close
'                'Busca la gerencia UAP del supervisor
'                Mi_SQL = "SELECT Empleado_ID,Gerencia_UAP FROM Cat_Empleados"
'                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000") & "'"
'                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
'                    If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
'                        .rdoColumns("Gerencia_UAP") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Gerencia_UAP")
'                    Else
'                        .rdoColumns("Gerencia_UAP") = Null
'                    End If
'                Else
                    .rdoColumns("Gerencia_UAP") = Null
'                End If
'                Rs_Alta_Cat_Empleados_Parentesco.Close
            Else
                'Busca la sección del supervisor
                Mi_SQL = "SELECT * FROM Cat_Secciones"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Txt_Cat_Empleados_Empleado_ID.Text & "'"
                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                    .rdoColumns("Nomipaq_ID") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                Else
                    .rdoColumns("Nomipaq_ID") = Trim(Txt_Cat_Empleados_Seccion.Text)
                End If
                Rs_Alta_Cat_Empleados_Parentesco.Close
                'Busca la gerencia UAP del supervisor
                .rdoColumns("Gerencia_UAP") = Null
            End If
            .rdoColumns("Departamento_ID") = Format(Cmb_Cat_Empleados_Departamento.ItemData(Cmb_Cat_Empleados_Departamento.ListIndex), "00000")
            .rdoColumns("Puesto_ID") = Format(Cmb_Cat_Empleados_Puesto.ItemData(Cmb_Cat_Empleados_Puesto.ListIndex), "00000")
            .rdoColumns("Turno_ID") = Format(Cmb_Cat_Empleados_Turno.ItemData(Cmb_Cat_Empleados_Turno.ListIndex), "00000")
            If Cmb_Cat_Empleados_Nivel_Estudio.ListIndex > -1 Then
                .rdoColumns("Nivel_Academico_ID") = Format(Cmb_Cat_Empleados_Nivel_Estudio.ItemData(Cmb_Cat_Empleados_Nivel_Estudio.ListIndex), "00000")
            End If
            If Cmb_Transporte.ListIndex > -1 Then
                .rdoColumns("Transporte_ID") = Format(Cmb_Transporte.ItemData(Cmb_Transporte.ListIndex), "00000")
            End If
            If Cmb_Gap.ListIndex > -1 Then
                .rdoColumns("Gap_ID") = Format(Cmb_Gap.ItemData(Cmb_Gap.ListIndex), "00000")
            End If
            .rdoColumns("Motivo_Baja_ID") = "00000"
            'Datos Personales
            .rdoColumns("Estatus") = Left(Cmb_Cat_Empleados_Estatus.Text, 1)
            .rdoColumns("Tipo") = Left(Cmb_Cat_Empleados_Tipo.Text, 1) 'Indica si es supervisor
            .rdoColumns("Nombre") = Trim(UCase(Txt_Cat_Empleados_Nombre.Text))
            .rdoColumns("Apellido_Paterno") = Trim(UCase(Txt_Cat_Empleados_Apellido_Paterno.Text))
            .rdoColumns("Apellido_Materno") = Trim(UCase(Txt_Cat_Empleados_Apellido_Materno.Text))
            .rdoColumns("Lugar_Nacimiento") = Trim(UCase(Txt_Cat_Empleados_Lugar_Nacimiento.Text))
            .rdoColumns("Sexo") = Trim(UCase(Cmb_Cat_Empleados_Sexo.Text))
            .rdoColumns("Fecha_Nacimiento") = Format(Dtp_Cat_Empleados_Fecha_Nacimiento.Value, "MM/dd/yyyy")
            .rdoColumns("Clave_Elector") = Trim(UCase(Txt_Cat_Empleados_Clave_Elector.Text))
            .rdoColumns("Estado_Civil") = Trim(UCase(Cmb_Cat_Empleados_Estado_Civil.Text))
            .rdoColumns("RFC") = Trim(Txt_Cat_Empleados_RFC.Text)
            .rdoColumns("Curp") = Trim(Txt_Cat_Empleados_Curp.Text)
            .rdoColumns("Nss") = Trim(Txt_Cat_Empleados_NSS.Text)
            .rdoColumns("Email") = Trim(Txt_Cat_Empleados_Email.Text)
            'Direccion
            .rdoColumns("Direccion") = Trim(Txt_Cat_Empleados_Direccion.Text)
            .rdoColumns("Colonia") = Trim(Txt_Cat_Empleados_Colonia.Text)
            .rdoColumns("Codigo_Postal") = Trim(Txt_Cat_Empleados_CP.Text)
            .rdoColumns("Ciudad") = Trim(Txt_Cat_Empleados_Ciudad.Text)
            .rdoColumns("Estado") = Trim(Txt_Cat_Empleados_Estado.Text)
            'Foto
            .rdoColumns("Imagen_Perfil") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) & ".JPG"
            'Valida que el directorio de fotos exista, si no lo crea
            If (Txt_Cat_Empleados_Ruta_Imagen.Text) <> "" Then
                If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil", "CARPETA") = False Then
                    MkDir App.Path & "\Perfil"
                End If
                'Guarda la imagen en la carpeta
                Dim Punto As Integer
                Dim Extension As String
                Dim Nombre_Archivo As String
                Punto = InStrRev(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), ".")
                Extension = Mid(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), Punto + 1, 5)
                Nombre_Archivo = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) + "." + Extension
                'Verifica que no esxista ya el archivo de perfil,
                If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & Nombre_Archivo, "ARCHIVO") = True Then
                    Kill App.Path & "\Perfil\" & Nombre_Archivo
                End If
                FileCopy Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), App.Path & "\Perfil\" & Nombre_Archivo
                If (Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & Nombre_Archivo, "ARCHIVO")) = True Then
                    .rdoColumns("Imagen_Perfil") = Nombre_Archivo
                Else
                    'MsgBox "No se pudo agregar la imagen, intente de nuevo", vbInformation + vbOKOnly, Me.Caption
                End If
            End If
            .rdoColumns("No_Tarjeta") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text)
            .rdoColumns("Clave_SAP") = "SI" & Trim(Txt_Cat_Empleados_No_Tarjeta.Text)
            .rdoColumns("Fecha_Ingreso") = Format(Dtp_Cat_Empleados_Fecha_Ingreso.Value, "MM/dd/yyyy")
            .rdoColumns("Tipo_Empleado") = Trim(UCase(Cmb_Cat_Empleados_Tipo_Empleado.Text))
            .rdoColumns("Tipo_Contratacion") = Trim(UCase(Cmb_Cat_Empleados_Contratacion.Text))
            .rdoColumns("Fecha_Termino_Contrato") = Format(Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value, "MM/dd/yyyy")
            .rdoColumns("Salario_Diario") = Val(Txt_Cat_Empleados_Salario_Diario.Text)
            .rdoColumns("Salario_Diario_Variable") = Val(Txt_Cat_Empleados_Vacaciones.Text)
            .rdoColumns("Cedula_Identidad_Ciudadana") = Cmb_Subdivision.Text
            If Chk_Cat_Empleados_Trabaja_Domingos.Value = 1 Then
                .rdoColumns("Trabaja_Domingos") = "S"
            Else
                .rdoColumns("Trabaja_Domingos") = "N"
            End If
            If Chk_Cat_Empleados_Infonavit.Value = 1 Then
                .rdoColumns("Infonavit") = "S"
            Else
                .rdoColumns("Infonavit") = "N"
            End If
            .rdoColumns("Retardos") = 0
            .rdoColumns("Fecha_Retardo") = Format("01/01/1960", "MM/dd/yyyy")
            .rdoColumns("En_Caso_Emergencia") = Trim(Txt_Cat_Empleados_LLamar_A.Text)
            .rdoColumns("Telefono_Emergencia1") = Trim(Txt_Cat_Empleados_Llamar_Telefono1.Text)
            .rdoColumns("Telefono_Emergencia2") = Trim(Txt_Cat_Empleados_Llamar_Telefono2.Text)
            .rdoColumns("Alergia1") = Trim(Txt_Cat_Empleados_Alergias1.Text)
            .rdoColumns("Alergia2") = Trim(Txt_Cat_Empleados_Alergias2.Text)
            .rdoColumns("Alergia3") = Trim(Txt_Cat_Empleados_Alergias3.Text)
            .rdoColumns("Fecha_Baja") = Format(Dtp_Cat_Empleados_Fecha_Baja.Value, "MM/dd/yyyy")
            If Cmb_Cat_Empleados_Motivos_Baja.ListIndex > -1 Then
                .rdoColumns("Motivo_Baja_ID") = Format(Cmb_Cat_Empleados_Motivos_Baja.ItemData(Cmb_Cat_Empleados_Motivos_Baja.ListIndex), "00000")
            End If
            .rdoColumns("Comentarios_Baja") = Trim(Conectar_Ayudante.Quitar_Caracter((Conectar_Ayudante.Quitar_Caracter(Txt_Cat_Empleados_Observaciones_Baja.Text, Chr(13))), Chr(10)))
            .rdoColumns("Campo_1") = Trim(Txt_Campo_1.Text)
            .rdoColumns("Campo_2") = Trim(Txt_Campo_2.Text)
            .rdoColumns("Campo_3") = Trim(Txt_Campo_3.Text)
            .rdoColumns("Campo_4") = Trim(Txt_Campo_4.Text)
            .rdoColumns("Campo_5") = Trim(Txt_Campo_5.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Cat_Empleados = Nothing
    'Alta de Parientes
    Dim Cont_Fila As Integer
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados_Parentesco")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Cat_Empleados_Dependientes.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Parentesco") = Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 0)
                .rdoColumns("Nombre") = Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 1)
                .rdoColumns("Fecha_Nacimiento") = Format(Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy")
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    'Alta de cursos
    Mi_SQL = "DELETE FROM Cat_Cursos_Detalles WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Cursos_Detalles")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Cursos.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Curso_ID") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Comentarios") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 2))
                .rdoColumns("Estatus") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 3))
                .rdoColumns("Fecha_Inicio") = Format(Grid_Cursos.TextMatrix(Cont_Fila, 4), "MM/dd/yyyy")
                .rdoColumns("Fecha_Fin") = Format(Grid_Cursos.TextMatrix(Cont_Fila, 5), "MM/dd/yyyy")
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    'Alta de evaluaciones
    Mi_SQL = "DELETE FROM Cat_Empleados_Evaluaciones WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados_Evaluaciones")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Evaluaciones.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Evaluacion") = Trim(Grid_Evaluaciones.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Fecha") = Format(Grid_Evaluaciones.TextMatrix(Cont_Fila, 1), "MM/dd/yyyy")
                .rdoColumns("Proxima_Evaluacion") = Format(Grid_Evaluaciones.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy")
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir o modificar los valoes
    Fra_Cat_Empleados_Datos_Personales.Enabled = False
    Fra_Cat_Empleados_Datos_Dependientes.Enabled = False
    Fra_Cat_Empleados_Datos_Laborales.Enabled = False
    Fra_Cat_Empleados_Baja.Enabled = False
    Fra_Cursos.Enabled = False
    Fra_Evaluaciones.Enabled = False
    Fra_Cat_Empleados.Enabled = True
    Cmb_Cat_Empleados_Estatus.Enabled = True
    Grid_Cat_Empleados_Dependientes.Rows = 0
    Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value = Now
    Dtp_Cat_Empleados_Fecha_Baja.Value = Now
    Dtp_Cat_Empleados_Fecha_Ingreso.Value = Now
    Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Now
    Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = Now
    Cmb_Cat_Empleados_Contratacion.ListIndex = -1
    Cmb_Cat_Empleados_Estado_Civil.ListIndex = -1
    Cmb_Cat_Empleados_Parentesco.ListIndex = -1
    Cmb_Cat_Empleados_Tipo.ListIndex = -1
    Cmb_Cat_Empleados_Sexo.ListIndex = -1
    Cmb_Cat_Empleados_Tipo_Empleado.ListIndex = -1
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Imprimir.Enabled = True
    'Pone un encabezado en el grid
    With Grid_Cat_Empleados
        If .Rows = 0 Then
            .AddItem "Empleado ID" & Chr(9) & "Nomina" & Chr(9) & "Apellidos" & Chr(9) & "Nombre" & Chr(9) & "Estatus"
        End If
        'Llena el grid con los datos del nuevo usuario
        .AddItem Trim(Txt_Cat_Empleados_Empleado_ID.Text) _
            & Chr(9) & Trim(Txt_Cat_Empleados_No_Tarjeta.Text) _
            & Chr(9) & Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) & " " & Trim(Txt_Cat_Empleados_Apellido_Materno.Text) _
            & Chr(9) & Trim(Txt_Cat_Empleados_Nombre.Text) _
            & Chr(9) & Left(Cmb_Cat_Empleados_Estatus.Text, 1)
        'Configura el tamaño de las columnas del grid_usuarios
        .FixedRows = 1
        .ColWidth(0) = 0      'Empleado_ID
        .ColWidth(1) = 1000   'No. Nomina
        .ColWidth(2) = 3000   'Apellidos
        .ColWidth(3) = 3000   'Nombre
        .ColWidth(4) = 700    'Estatus
    End With
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Img_Cat_Empleados_Foto.picture = LoadPicture("")
    Grid_Cat_Empleados_Dependientes.Rows = 0
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empleados", Me)
    MsgBox "Empleado dado de alta", vbOKOnly + vbInformation, Me.Caption
    If Cmb_Cat_Empleados_Supervisor.ListIndex > -1 Then
        MsgBox "No olvide asignar Supervisor al empleado" + vbCrLf + "Para generar la lista de validacion de horas trabajadas", vbOKOnly + vbInformation, Me.Caption
    End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Cat_Empleados
'DESCRIPCION: Modifica el registro del empleado
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 15-Mayo-2009
'MODIFICO   : Sergio Ulises Durán Hernández
'FECHA_MODIFICO: 12-Marzo-2012
'CAUSA_MODIFICO: Adecuación al sistema de SRG
'*******************************************************************************
Private Sub Modifica_Cat_Empleados()
Dim Rs_Modificacion_Cat_Empleados As rdoResultset 'Informacion del registro
Dim Rs_Alta_Cat_Empleados_Parentesco As rdoResultset
Dim Rs_Adm_Cambios_Turnos As rdoResultset

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Empleado_ID ='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Empleados = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Empleados
        .Edit
            .rdoColumns("Empresa_ID") = Format(Cmb_Cat_Empleados_Empresa.ItemData(Cmb_Cat_Empleados_Empresa.ListIndex), "00000")
            .rdoColumns("Supervisor_ID") = Null
            If Cmb_Cat_Empleados_Supervisor.ListIndex > -1 Then
                .rdoColumns("Supervisor_ID") = Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000")
                'Busca la sección del supervisor
                Mi_SQL = "SELECT * FROM Cat_Secciones"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000") & "'"
                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                    .rdoColumns("Nomipaq_ID") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                    Txt_Cat_Empleados_Seccion.Text = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                Else
                    .rdoColumns("Nomipaq_ID") = Trim(Txt_Cat_Empleados_Seccion.Text)
                End If
                Rs_Alta_Cat_Empleados_Parentesco.Close
                'Busca la gerencia UAP del supervisor
                Mi_SQL = "SELECT Empleado_ID,Gerencia_UAP FROM Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Format(Cmb_Cat_Empleados_Supervisor.ItemData(Cmb_Cat_Empleados_Supervisor.ListIndex), "00000") & "'"
                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                    If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                        If Not IsNull(Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Gerencia_UAP")) Then
                            .rdoColumns("Gerencia_UAP") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Gerencia_UAP")
                        Else
                            .rdoColumns("Gerencia_UAP") = Null
                        End If
                    Else
                        .rdoColumns("Gerencia_UAP") = Null
                    End If
                Else
                    .rdoColumns("Gerencia_UAP") = Null
                End If
                Rs_Alta_Cat_Empleados_Parentesco.Close
            Else
                'Busca la sección del supervisor
                Mi_SQL = "SELECT * FROM Cat_Secciones"
                Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Txt_Cat_Empleados_Empleado_ID.Text & "'"
                Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Alta_Cat_Empleados_Parentesco.EOF Then
                    .rdoColumns("Nomipaq_ID") = Rs_Alta_Cat_Empleados_Parentesco.rdoColumns("Clave")
                Else
                    .rdoColumns("Nomipaq_ID") = Trim(Txt_Cat_Empleados_Seccion.Text)
                End If
                Rs_Alta_Cat_Empleados_Parentesco.Close
                'Busca la gerencia UAP del supervisor
                .rdoColumns("Gerencia_UAP") = Null
            End If
            .rdoColumns("Departamento_ID") = Format(Cmb_Cat_Empleados_Departamento.ItemData(Cmb_Cat_Empleados_Departamento.ListIndex), "00000")
            .rdoColumns("Puesto_ID") = Format(Cmb_Cat_Empleados_Puesto.ItemData(Cmb_Cat_Empleados_Puesto.ListIndex), "00000")
            'Valida si existe un cambio de turno para registrar el evento en cambio de turnos (porque se va areflejar para SAP)
            If .rdoColumns("Turno_ID") <> Format(Cmb_Cat_Empleados_Turno.ItemData(Cmb_Cat_Empleados_Turno.ListIndex), "00000") Then
                'Inserta el registro en la tabla de cambios de turnos
                Set Rs_Adm_Cambios_Turnos = Conectar_Ayudante.Recordset_Agregar("Adm_Cambios_Turnos")
                Rs_Adm_Cambios_Turnos.AddNew
                    Rs_Adm_Cambios_Turnos.rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                    Rs_Adm_Cambios_Turnos.rdoColumns("Turno_Anterior_ID") = .rdoColumns("Turno_ID")
                    Rs_Adm_Cambios_Turnos.rdoColumns("Turno_Nuevo_ID") = Format(Cmb_Cat_Empleados_Turno.ItemData(Cmb_Cat_Empleados_Turno.ListIndex), "00000")
                    Rs_Adm_Cambios_Turnos.rdoColumns("Fecha_Cambio") = Format(Now, "MM/dd/yyyy")
                    Rs_Adm_Cambios_Turnos.rdoColumns("Estatus") = "CAMBIADO"
                    Rs_Adm_Cambios_Turnos.rdoColumns("Usuario_Creo") = Nombre_Usuario
                    Rs_Adm_Cambios_Turnos.rdoColumns("Fecha_Creo") = Now
                Rs_Adm_Cambios_Turnos.Update
                Rs_Adm_Cambios_Turnos.Close
                'Cambia el registro en el catálogo
                .rdoColumns("Turno_ID") = Format(Cmb_Cat_Empleados_Turno.ItemData(Cmb_Cat_Empleados_Turno.ListIndex), "00000")
            End If
            If Cmb_Cat_Empleados_Nivel_Estudio.ListIndex > -1 Then
                .rdoColumns("Nivel_Academico_ID") = Format(Cmb_Cat_Empleados_Nivel_Estudio.ItemData(Cmb_Cat_Empleados_Nivel_Estudio.ListIndex), "00000")
            End If
            If Cmb_Transporte.ListIndex > -1 Then
                .rdoColumns("Transporte_ID") = Format(Cmb_Transporte.ItemData(Cmb_Transporte.ListIndex), "00000")
            Else
                .rdoColumns("Transporte_ID") = Null
            End If
            If Cmb_Gap.ListIndex > -1 Then
                .rdoColumns("Gap_ID") = Format(Cmb_Gap.ItemData(Cmb_Gap.ListIndex), "00000")
            Else
                .rdoColumns("Gap_ID") = Null
            End If
            'Datos Personales
            .rdoColumns("Estatus") = Left(Cmb_Cat_Empleados_Estatus.Text, 1)
            .rdoColumns("Tipo") = Left(Cmb_Cat_Empleados_Tipo.Text, 1) 'Indica si es supervisor
            .rdoColumns("Nombre") = Trim(UCase(Txt_Cat_Empleados_Nombre.Text))
            .rdoColumns("Apellido_Paterno") = Trim(UCase(Txt_Cat_Empleados_Apellido_Paterno.Text))
            .rdoColumns("Apellido_Materno") = Trim(UCase(Txt_Cat_Empleados_Apellido_Materno.Text))
            .rdoColumns("Lugar_Nacimiento") = Trim(UCase(Txt_Cat_Empleados_Lugar_Nacimiento.Text))
            .rdoColumns("Sexo") = Trim(UCase(Cmb_Cat_Empleados_Sexo.Text))
            .rdoColumns("Fecha_Nacimiento") = Format(Dtp_Cat_Empleados_Fecha_Nacimiento.Value, "MM/dd/yyyy")
            .rdoColumns("Clave_Elector") = Trim(UCase(Txt_Cat_Empleados_Clave_Elector.Text))
            .rdoColumns("Estado_Civil") = Trim(UCase(Cmb_Cat_Empleados_Estado_Civil.Text))
            .rdoColumns("RFC") = Trim(Txt_Cat_Empleados_RFC.Text)
            .rdoColumns("Curp") = Trim(Txt_Cat_Empleados_Curp.Text)
            .rdoColumns("Nss") = Trim(Txt_Cat_Empleados_NSS.Text)
            .rdoColumns("Email") = Trim(Txt_Cat_Empleados_Email.Text)
            'Direccion
            .rdoColumns("Direccion") = Trim(Txt_Cat_Empleados_Direccion.Text)
            .rdoColumns("Colonia") = Trim(Txt_Cat_Empleados_Colonia.Text)
            .rdoColumns("Codigo_Postal") = Trim(Txt_Cat_Empleados_CP.Text)
            .rdoColumns("Ciudad") = Trim(Txt_Cat_Empleados_Ciudad.Text)
            .rdoColumns("Estado") = Trim(Txt_Cat_Empleados_Estado.Text)
            .rdoColumns("Imagen_Perfil") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) & ".JPG"
            'Foto
            If Not IsNull(.rdoColumns("Imagen_Perfil")) Then
                Dim Punto As Integer
                Dim Extension As String
                Dim Nombre_Archivo As String
                If Txt_Cat_Empleados_Ruta_Imagen.Text <> .rdoColumns("Imagen_Perfil") Then
                    Punto = InStrRev(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), ".")
                    Extension = Mid(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), Punto + 1, 5)
                    Nombre_Archivo = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) + "." + Extension
                    'Borra la imagen anterior
                    If Trim(.rdoColumns("Imagen_Perfil")) <> "" Then
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & Trim(.rdoColumns("Imagen_Perfil")), "ARCHIVO") = True Then
                            Kill PG_Ruta_Fotos & "\" & .rdoColumns("Imagen_Perfil")
                        End If
                    End If
                    If Len(Txt_Cat_Empleados_Ruta_Imagen.Text) > 0 Then
                        FileCopy Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), PG_Ruta_Fotos & "\" & Nombre_Archivo
                        .rdoColumns("Imagen_Perfil") = Nombre_Archivo
                    End If
                Else
                    Txt_Cat_Empleados_Ruta_Imagen.Text = ""
                    .rdoColumns("Imagen_Perfil") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) & ".JPG"
                End If
            Else 'Si no se contenia imagen anterior la agrega nueva
                If Txt_Cat_Empleados_Ruta_Imagen.Text <> "" Then
                    Punto = InStrRev(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), ".")
                    Extension = Mid(Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), Punto + 1, 5)
                    Nombre_Archivo = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) + "." + Extension
                    FileCopy Trim(Txt_Cat_Empleados_Ruta_Imagen.Text), PG_Ruta_Fotos & "\" & Nombre_Archivo
                    .rdoColumns("Imagen_Perfil") = Nombre_Archivo
                Else
                    Txt_Cat_Empleados_Ruta_Imagen.Text = ""
                    .rdoColumns("Imagen_Perfil") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text) & ".JPG"
                End If
            End If
            .rdoColumns("Nomipaq_ID") = Trim(Txt_Cat_Empleados_Seccion.Text)
            .rdoColumns("No_Tarjeta") = Trim(Txt_Cat_Empleados_No_Tarjeta.Text)
            .rdoColumns("Clave_SAP") = "SI" & Trim(Txt_Cat_Empleados_No_Tarjeta.Text)
            .rdoColumns("Fecha_Ingreso") = Format(Dtp_Cat_Empleados_Fecha_Ingreso.Value, "MM/dd/yyyy")
            .rdoColumns("Tipo_Empleado") = Trim(UCase(Cmb_Cat_Empleados_Tipo_Empleado.Text))
            .rdoColumns("Tipo_Contratacion") = Trim(UCase(Cmb_Cat_Empleados_Contratacion.Text))
            .rdoColumns("Fecha_Termino_Contrato") = Format(Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value, "MM/dd/yyyy")
            .rdoColumns("Salario_Diario") = Val(Txt_Cat_Empleados_Salario_Diario.Text)
            .rdoColumns("Salario_Diario_Variable") = Val(Txt_Cat_Empleados_Vacaciones.Text)
            .rdoColumns("Cedula_Identidad_Ciudadana") = Cmb_Subdivision.Text
            If Chk_Cat_Empleados_Trabaja_Domingos.Value = 1 Then
                .rdoColumns("Trabaja_Domingos") = "S"
            Else
                .rdoColumns("Trabaja_Domingos") = "N"
            End If
            If Chk_Cat_Empleados_Infonavit.Value = 1 Then
                .rdoColumns("Infonavit") = "S"
            Else
                .rdoColumns("Infonavit") = "N"
            End If
            .rdoColumns("En_Caso_Emergencia") = Trim(Txt_Cat_Empleados_LLamar_A.Text)
            .rdoColumns("Telefono_Emergencia1") = Trim(Txt_Cat_Empleados_Llamar_Telefono1.Text)
            .rdoColumns("Telefono_Emergencia2") = Trim(Txt_Cat_Empleados_Llamar_Telefono2.Text)
            .rdoColumns("Alergia1") = Trim(Txt_Cat_Empleados_Alergias1.Text)
            .rdoColumns("Alergia2") = Trim(Txt_Cat_Empleados_Alergias2.Text)
            .rdoColumns("Alergia3") = Trim(Txt_Cat_Empleados_Alergias3.Text)
            .rdoColumns("Fecha_Baja") = Format(Dtp_Cat_Empleados_Fecha_Baja.Value, "MM/dd/yyyy")
            .rdoColumns("Motivo_Baja_ID") = "00000"
            If Cmb_Cat_Empleados_Motivos_Baja.ListIndex > -1 Then
                .rdoColumns("Motivo_Baja_ID") = Format(Cmb_Cat_Empleados_Motivos_Baja.ItemData(Cmb_Cat_Empleados_Motivos_Baja.ListIndex), "00000")
            End If
            .rdoColumns("Comentarios_Baja") = Trim(Conectar_Ayudante.Quitar_Caracter((Conectar_Ayudante.Quitar_Caracter(Txt_Cat_Empleados_Observaciones_Baja.Text, Chr(13))), Chr(10)))
            .rdoColumns("Campo_1") = Trim(Txt_Campo_1.Text)
            .rdoColumns("Campo_2") = Trim(Txt_Campo_2.Text)
            .rdoColumns("Campo_3") = Trim(Txt_Campo_3.Text)
            .rdoColumns("Campo_4") = Trim(Txt_Campo_4.Text)
            .rdoColumns("Campo_5") = Trim(Txt_Campo_5.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Empleados = Nothing
    'Alta de Parientes
    Mi_SQL = "DELETE FROM Cat_Empleados_Parentesco WHERE Empleado_ID = '" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    Dim Cont_Fila As Integer
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados_Parentesco")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Cat_Empleados_Dependientes.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Parentesco") = Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 0)
                .rdoColumns("Nombre") = Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 1)
                .rdoColumns("Fecha_Nacimiento") = Format(Grid_Cat_Empleados_Dependientes.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy")
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    'Alta de cursos
    Mi_SQL = "DELETE FROM Cat_Cursos_Detalles WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Cursos_Detalles")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Cursos.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Curso_ID") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Comentarios") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 2))
                .rdoColumns("Estatus") = Trim(Grid_Cursos.TextMatrix(Cont_Fila, 3))
                .rdoColumns("Fecha_Inicio") = Format(Grid_Cursos.TextMatrix(Cont_Fila, 4), "MM/dd/yyyy")
                .rdoColumns("Fecha_Fin") = Format(Grid_Cursos.TextMatrix(Cont_Fila, 5), "MM/dd/yyyy")
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    'Alta de evaluaciones
    Mi_SQL = "DELETE FROM Cat_Empleados_Evaluaciones WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "' "
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empleados_Parentesco = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados_Evaluaciones")
    With Rs_Alta_Cat_Empleados_Parentesco
        For Cont_Fila = 1 To Grid_Evaluaciones.Rows - 1
            .AddNew
                .rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                .rdoColumns("Evaluacion") = Trim(Grid_Evaluaciones.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Fecha") = Format(Grid_Evaluaciones.TextMatrix(Cont_Fila, 1), "MM/dd/yyyy")
                .rdoColumns("Proxima_Evaluacion") = Format(Grid_Evaluaciones.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy")
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empleados_Parentesco = Nothing
    With Grid_Cat_Empleados
        .TextMatrix(.RowSel, 1) = Trim(Txt_Cat_Empleados_No_Tarjeta.Text)
        .TextMatrix(.RowSel, 2) = Trim(Txt_Cat_Empleados_Apellido_Paterno.Text) & " " & Trim(Txt_Cat_Empleados_Apellido_Materno.Text)
        .TextMatrix(.RowSel, 3) = Trim(Txt_Cat_Empleados_Nombre.Text)
        .TextMatrix(.RowSel, 4) = Left(Cmb_Cat_Empleados_Estatus.Text, 1)
    End With
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Cat_Empleados_Datos_Personales.Enabled = False
    Fra_Cat_Empleados_Datos_Dependientes.Enabled = False
    Fra_Cat_Empleados_Datos_Laborales.Enabled = False
    Fra_Cat_Empleados_Baja.Enabled = False
    Fra_Cursos.Enabled = False
    Fra_Evaluaciones.Enabled = False
    Fra_Cat_Empleados.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Imprimir.Enabled = True
    Grid_Cat_Empleados_Dependientes.Rows = 0
    Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value = Now
    Dtp_Cat_Empleados_Fecha_Baja.Value = Now
    Dtp_Cat_Empleados_Fecha_Ingreso.Value = Now
    Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Now
    Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = Now
    Cmb_Cat_Empleados_Contratacion.ListIndex = -1
    Cmb_Cat_Empleados_Estado_Civil.ListIndex = -1
    Cmb_Cat_Empleados_Parentesco.ListIndex = -1
    Cmb_Cat_Empleados_Tipo.ListIndex = -1
    Cmb_Cat_Empleados_Sexo.ListIndex = -1
    Cmb_Cat_Empleados_Tipo_Empleado.ListIndex = -1
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Img_Cat_Empleados_Foto.picture = LoadPicture("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empleados", Me)
    MsgBox "El empleado ha sido modificado", vbInformation + vbOKOnly, Me.Caption
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Inicializa()
    Select Case Catalogo
        Case "Cat_Empleados":
            'Consulta_Cat_Empleados "" 'Consulta todos los empleados que estan dados de alta
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Cat_Empleados_Empresa, 0, "Nombre", "", False, "")
            Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Cat_Empleados_Turno, 0, "Nombre")
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND Tipo = 'S'", Cmb_Cat_Empleados_Supervisor, 0, "Nombre")
            Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Cat_Empleados_Departamento, 0, "Nombre")
            Call Conectar_Ayudante.Llena_Combo_Item("Transporte_ID,Nombre", "Cat_Transportes", Cmb_Transporte, 0, "Nombre")
            Call Conectar_Ayudante.Llena_Combo_Item("Gap_ID,Nombre", "Cat_Gaps", Cmb_Gap, 0, "Nombre")
            'Llena el combo de puestos
            Call Conectar_Ayudante.Llena_Combo_Item("Puesto_ID, Nombre", "Cat_Puestos", Cmb_Cat_Empleados_Puesto, 0, "Nombre", , False, "TODOS")
            'LLena el combo de niveles de estudio
            Call Conectar_Ayudante.Llena_Combo_Item("Nivel_Estudio_ID, Nombre", "Cat_Nivel_Estudio", Cmb_Cat_Empleados_Nivel_Estudio, 0, "Nombre", , False, "TODOS")
            'Llena los motivos de baja
            Call Conectar_Ayudante.Llena_Combo_Item("Motivo_Baja_ID,Nombre", "Cat_Motivos_Baja", Cmb_Cat_Empleados_Motivos_Baja, 0, "Nombre", , False, "TODOS")
            Dtp_Cat_Empleados_Fecha_Ingreso.Value = Now
            Dtp_Cat_Empleados_Fecha_Nacimiento.Value = Now
            Dtp_Cat_Empleados_Fecha_Termino_Contrato.Value = Now
            Dtp_Cat_Empleados_Fecha_Baja.Value = Now
            Dtp_Cat_Empleados_Dependiente_Fecha_Nacimiento.Value = Now
            Cmb_Estatus_Curso.ListIndex = 0
            Dtp_Fecha_Inicio.Value = Now
            Dtp_Fecha_Fin.Value = Now
            Dtp_Fecha_Evaluacion.Value = Now
            Dtp_Proxima_Evaluacion.Value = Now
            Dim Visible As Boolean
            If Rol_ID = 1 _
            Or Rol_ID = 10 Then
                Visible = True
            Else
                If Rol_ID = 4 Then
                    Dim Cont_Tab As Integer
                    For Cont_Tab = 0 To Tab_Cat_Empleados.Tabs - 1
                        Tab_Cat_Empleados.TabEnabled(Cont_Tab) = False
                    Next Cont_Tab
                    Tab_Cat_Empleados.TabEnabled(2) = True
                    Tab_Cat_Empleados.Tab = 2
                Else
                    Visible = False
                End If
            End If
            Lbl_Laborales(5).Visible = Visible
            Txt_Cat_Empleados_Salario_Diario.Visible = Visible
    End Select
End Sub

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
        MsgBox "El archivo que está intentando abrir no se encontró en el directorio indicado.  ", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    MsgBox Err.Description
End Sub

Private Sub Txt_Evaluacion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Imprimir_Contrato
'DESCRIPCION: Consulta los empleados para permitir o no imprimir el contrato
'PARAMETROS : Nombre- Indica el parámetro de búsqueda
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 05-Abr-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Imprimir_Contrato()
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT * FROM Cat_Empleados WHERE Empleado_ID = '" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
    
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Txt_Log.Text = Txt_Log.Text & "2. Consulta Empleados ejecutada..." & vbCrLf
    With Rs_Consulta_Cat_Empleados
        If Not .EOF Then
        Dim Cantidad_En_Letras As String
        Cantidad_En_Letras = Conectar_Ayudante.NroEnLetras(Val(.rdoColumns("Salario_Diario")))
        Txt_Log.Text = Txt_Log.Text & "3. Salario convertido a letras..." & vbCrLf
        
            If LCase(.rdoColumns("Tipo_Empleado")) = "sindicalizado" Then
                Crea_PDF_Contrato "Rpt_Contratos_Empleado_Sindicalizados", "Contrato_" & .rdoColumns("Nombre") & "_" & .rdoColumns("Apellido_Paterno") & "_" & .rdoColumns("Apellido_Materno"), Cantidad_En_Letras
            ElseIf LCase(.rdoColumns("Tipo_Empleado")) = "confianza" Then
                Crea_PDF_Contrato "Rpt_Contratos_Empleado_No_Sindicalizados", "Contrato_" & .rdoColumns("Nombre") & "_" & .rdoColumns("Apellido_Paterno") & "_" & .rdoColumns("Apellido_Materno"), Cantidad_En_Letras
            Else
                MsgBox ("Revise el tipo de empleado")
           End If
        Else
        MsgBox ("Es necesario dar de alta al empleado para poder imprimir su contrato")
        End If
    End With
    Rs_Consulta_Cat_Empleados.Close
End Sub
            
Private Function Validar_Imprimir_Contrato() As Boolean
    Validar_Imprimir_Contrato = True
    If Len(Txt_Cat_Empleados_Empleado_ID.Text) <= 0 Then
        Validar_Imprimir_Contrato = False
    End If
End Function

Public Sub Crea_PDF_Contrato(Reporte_Rpt As String, Nombre As String, Cantidad_Letras As String)
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
    Txt_Log.Text = Txt_Log.Text & "4. Creando documento Contrato..." & vbCrLf
    'Asigna el formato de la factura a la variable
    If Dir(App.Path & "\Contratos\", vbDirectory) = "" Then
    MkDir (App.Path & "\Contratos\")
    End If
     Ruta_RPT = App.Path & "\Reportes\" & Reporte_Rpt & ".rpt"
     Ruta_Salida = App.Path & "\Contratos\" & Nombre & ".doc"
     Set crxReport = crxApplication.OpenReport(Ruta_RPT)
           
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    'Asigna los datos de conexion de la base de datos
    Txt_Log.Text = Txt_Log.Text & "5. Abriendo conexión ODBC..." & vbCrLf
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo "SRG_Recursos_Humanos", Conectar_Ayudante.Base, Conectar_Ayudante.Usuario_Conexion, Conectar_Ayudante.Password
            End Select
        Next
    End With
    Txt_Log.Text = Txt_Log.Text & "6. Conexión ODBC abierta..." & vbCrLf
    'Asigna los datos a los parametros
    Txt_Log.Text = Txt_Log.Text & "7. Asignando parámetros para Contrato..." & vbCrLf
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
    Dim Fecha As Date
    Dim parametro As String
        Select Case crParamDef.ParameterFieldName
        'Cursos_Tomados_Por_Empleado
            Case "Parametro_Empleado_ID"
                  parametro = Format(Txt_Cat_Empleados_Empleado_ID.Text, "00000")
                 crParamDef.AddCurrentValue ("'" & parametro & "'")
            Case "Cantidad_En_Letras"
                  parametro = Cantidad_Letras
                  crParamDef.AddCurrentValue (parametro)
'            Case "Fecha_Inicio_Cursos_Indices_Asistencia"
'                If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
'                    Fecha = Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Inicio.Value, "MM/dd/yyyy") & " 00:00:00"
'                Else
'                    Fecha = Format("01/01/1990", "MM/dd/yyyy") & " 00:00:00"
'                End If
'                    crParamDef.AddCurrentValue (Fecha)
'
'            Case "Fecha_Fin_Cursos_Indices_Asistencia"
'                If Chk_Rpt_Cursos_Indice_Asistencias_Fechas.Value = 1 Then
'                   Fecha = Format(Dtp_Rpt_Cursos_Indice_Asistencias_Fecha_Fin.Value, "MM/dd/yyyy") & " 23:59:59"
'                Else
'                   Fecha = Format("12/31/2100", "MM/dd/yyyy") & " 23:59:59"
'                End If
'                crParamDef.AddCurrentValue (Fecha)
'
        End Select
    Next
    Txt_Log.Text = Txt_Log.Text & "8. Parámetros para el Contrato asignados..." & vbCrLf
    Txt_Log.Text = Txt_Log.Text & "9. Exportando reporte a Word..." & vbCrLf
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
   crxReport.ExportOptions.DiskFileName = Ruta_Salida

   

'    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat

    crxReport.ExportOptions.FormatType = crEFTWordForWindows
'crxReport.ExportOptions.FormatType = crEFTExcel97
'    crxReport.ExportOptions.PDFExportAllPages = True
    'Oculta el progreso de la exportacion
    crxReport.DisplayProgressDialog = False
    'Genera la exportación del documento
    On Error Resume Next
    crxReport.Export (False)
    'Destruye el documento
    Set crxReport = Nothing
    Txt_Log.Text = Txt_Log.Text & "10. Abriendo Contrato en documento de Word..." & vbCrLf
    ShellExecute Me.hwnd, "open", Ruta_Salida, "", "", 4
    Txt_Log.Text = Txt_Log.Text & "11. Contrato Abierto correctamente..." & vbCrLf
    Txt_Log.Visible = False
    Exit Sub
HANDLER:
    Printer.EndDoc
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

