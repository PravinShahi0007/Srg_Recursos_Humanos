VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Cat_Parametros_RH 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PARÁMETROS"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   690
      Left            =   5475
      Picture         =   "Frm_Cat_Parametros_RH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   690
      Left            =   75
      Picture         =   "Frm_Cat_Parametros_RH.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "A"
      Top             =   4800
      Width           =   1200
   End
   Begin TabDlg.SSTab Tab_Parametros 
      Height          =   5700
      Left            =   45
      TabIndex        =   16
      Top             =   7230
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Sistema"
      TabPicture(0)   =   "Frm_Cat_Parametros_RH.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Parametros_Sistema"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Incidencias"
      TabPicture(1)   =   "Frm_Cat_Parametros_RH.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Parametros_Incidencias"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Parametros_Incidencias 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   5250
         Left            =   -74910
         TabIndex        =   35
         Top             =   330
         Width           =   6525
         Begin VB.Frame Fra_Parametros_Incidencias_Horas_Dobles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horas Dobles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   3480
            TabIndex        =   71
            Top             =   180
            Visible         =   0   'False
            Width           =   2985
            Begin VB.TextBox Txt_Parametros_PDF_Horas_Dobles 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   73
               Top             =   569
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_Pago_horas_Dobles 
               Height          =   330
               Left            =   1935
               TabIndex        =   72
               Top             =   202
               Width           =   960
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Clave NOI"
               Height          =   195
               Left            =   135
               TabIndex        =   75
               Top             =   637
               Width           =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cantidad Horas"
               Height          =   195
               Left            =   135
               TabIndex        =   74
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.Frame Fra_Parametros_Incidencias_Horas_Triples 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horas Triples"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   3480
            TabIndex        =   66
            Top             =   1140
            Visible         =   0   'False
            Width           =   2985
            Begin VB.TextBox Txt_Parametros_Pago_horas_Triples 
               Height          =   330
               Left            =   1935
               TabIndex        =   68
               Top             =   202
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Horas_Triples 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   67
               Top             =   570
               Width           =   960
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cantidad Horas"
               Height          =   195
               Left            =   135
               TabIndex        =   70
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Clave NOI"
               Height          =   195
               Left            =   135
               TabIndex        =   69
               Top             =   630
               Width           =   735
            End
         End
         Begin VB.Frame Fra_Parametros_Incidencias_Retardos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Retardos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1890
            Left            =   60
            TabIndex        =   57
            Top             =   180
            Width           =   2985
            Begin VB.TextBox Txt_Parametros_Retardos_Periodo_Dias 
               Height          =   330
               Left            =   1935
               TabIndex        =   61
               Top             =   569
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_Retardos_Dias_Falta 
               Height          =   330
               Left            =   1935
               TabIndex        =   60
               Top             =   202
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_Clave_Falta_Retardos 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   59
               Top             =   1320
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_Retardos_Tolerancia 
               Height          =   330
               Left            =   1935
               TabIndex        =   58
               Top             =   936
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Clave NOI"
               Height          =   195
               Left            =   135
               TabIndex        =   65
               Top             =   1373
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Periodo"
               Height          =   195
               Left            =   135
               TabIndex        =   64
               Top             =   637
               Width           =   540
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cantidad Dias para falta"
               Height          =   195
               Left            =   135
               TabIndex        =   63
               Top             =   270
               Width           =   1695
            End
            Begin VB.Label Lbl_Tolerancia 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tolerancia (min.)"
               Height          =   195
               Left            =   135
               TabIndex        =   62
               Top             =   1005
               Width           =   1170
            End
         End
         Begin VB.Frame Fra_Parametros_Incidencias_Permisos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Claves Nomipaq Solicitud de Permisos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3105
            Left            =   60
            TabIndex        =   36
            Top             =   2100
            Visible         =   0   'False
            Width           =   6405
            Begin VB.TextBox Txt_Parametros_PDF_Sancion 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   84
               Top             =   2040
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Permiso_Sin_Goce 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   82
               Top             =   2430
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Permiso_Goce 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   80
               Top             =   2057
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Riesgo_Trabajo 
               Height          =   330
               Left            =   1935
               TabIndex        =   46
               Top             =   944
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Vacaciones 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   45
               Top             =   1315
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Enfermedad_General 
               Height          =   330
               Left            =   1935
               TabIndex        =   44
               Top             =   202
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Maternidad 
               Height          =   330
               Left            =   1935
               TabIndex        =   43
               Top             =   573
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Alumbramiento 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   42
               Top             =   202
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Defuncion 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   41
               Top             =   569
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Matrimonio 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   40
               Top             =   936
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Falta_Justificada 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   39
               Top             =   1305
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Permiso_Temporal 
               Height          =   330
               Left            =   1935
               MaxLength       =   4
               TabIndex        =   38
               Top             =   1686
               Width           =   960
            End
            Begin VB.TextBox Txt_Parametros_PDF_Falta_InJustificada 
               Height          =   330
               Left            =   5340
               MaxLength       =   4
               TabIndex        =   37
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Sanción"
               Height          =   195
               Left            =   3540
               TabIndex        =   85
               Top             =   2115
               Width           =   585
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Permiso SG"
               Height          =   195
               Left            =   135
               TabIndex        =   83
               Top             =   2498
               Width           =   825
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Permiso CG"
               Height          =   195
               Left            =   135
               TabIndex        =   81
               Top             =   2125
               Width           =   825
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Riesgo de Trabajo"
               Height          =   195
               Left            =   135
               TabIndex        =   56
               Top             =   1012
               Width           =   1305
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Enfermedad General"
               Height          =   195
               Left            =   135
               TabIndex        =   55
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Maternidad"
               Height          =   195
               Left            =   135
               TabIndex        =   54
               Top             =   641
               Width           =   795
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vacaciones"
               Height          =   195
               Left            =   135
               TabIndex        =   53
               Top             =   1383
               Width           =   840
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Alumbramiento"
               Height          =   195
               Left            =   3540
               TabIndex        =   52
               Top             =   270
               Width           =   1035
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Defuncion"
               Height          =   195
               Left            =   3540
               TabIndex        =   51
               Top             =   630
               Width           =   735
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Matrimonio"
               Height          =   195
               Left            =   3540
               TabIndex        =   50
               Top             =   1005
               Width           =   765
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Falta Justificada"
               Height          =   195
               Left            =   3540
               TabIndex        =   49
               Top             =   1380
               Width           =   1140
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Permiso Temporal"
               Height          =   195
               Left            =   135
               TabIndex        =   48
               Top             =   1754
               Width           =   1260
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Falta Injustificada"
               Height          =   195
               Left            =   3540
               TabIndex        =   47
               Top             =   1755
               Width           =   1230
            End
         End
      End
      Begin VB.Frame Fra_Parametros_Sistema 
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
         Height          =   5250
         Left            =   90
         TabIndex        =   25
         Top             =   390
         Width           =   6525
         Begin VB.TextBox Txt_Parametros_Aviso_Contratacion 
            Height          =   330
            Left            =   3900
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   78
            Top             =   2985
            Width           =   2400
         End
         Begin VB.TextBox Txt_Parametros_Edad_Minima_Contratacion 
            Height          =   330
            Left            =   3900
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   2610
            Width           =   2400
         End
         Begin VB.TextBox Txt_Parametros_Email_Notificacion 
            Height          =   585
            Left            =   1605
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            ToolTipText     =   "Las direcciones de correo deben estar separadas por punto y coma (;)"
            Top             =   1965
            Width           =   4695
         End
         Begin VB.TextBox Txt_Parametros_Email_Administrador 
            Height          =   330
            Left            =   1605
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            ToolTipText     =   "Las direcciones de correo deben estar separadas por punto y coma (;)"
            Top             =   600
            Width           =   4695
         End
         Begin VB.TextBox Txt_Parametros_Puerto_Smtp 
            Height          =   285
            Left            =   5145
            MaxLength       =   5
            TabIndex        =   18
            Top             =   248
            Width           =   1125
         End
         Begin VB.Frame Fra_Parametros_Importacion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Importación Automática"
            Height          =   1035
            Left            =   1605
            TabIndex        =   28
            Top             =   4050
            Width           =   4695
            Begin MSComCtl2.DTPicker Dtp_Parametros_Hora_Importacion 
               Height          =   315
               Left            =   2610
               TabIndex        =   23
               Top             =   285
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm"
               Format          =   110559235
               UpDown          =   -1  'True
               CurrentDate     =   40007
            End
            Begin MSComCtl2.DTPicker Dtp_Parametros_Hora_Importacion_Dia 
               Height          =   315
               Left            =   2610
               TabIndex        =   24
               Top             =   660
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm"
               Format          =   110559235
               UpDown          =   -1  'True
               CurrentDate     =   40007
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ejecución de Importacion Diaria"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Visible         =   0   'False
               Width           =   2250
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ejecución de Importacion"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   345
               Width           =   1800
            End
         End
         Begin VB.TextBox Txt_Parametros_Email_Sistema 
            Height          =   330
            Left            =   1605
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            ToolTipText     =   "Las direcciones de correo deben estar separadas por punto y coma (;)"
            Top             =   975
            Width           =   4695
         End
         Begin VB.TextBox Txt_Parametros_Email 
            Height          =   585
            Left            =   1605
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            ToolTipText     =   "Las direcciones de correo deben estar separadas por punto y coma (;)"
            Top             =   1335
            Width           =   4695
         End
         Begin VB.TextBox Txt_Parametros_Servidor_SMTP 
            Height          =   330
            Left            =   1605
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   225
            Width           =   2400
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vencimiento Contratación"
            Height          =   195
            Left            =   1605
            TabIndex        =   79
            Top             =   3060
            Width           =   1815
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edad Contratación"
            Height          =   195
            Left            =   1605
            TabIndex        =   77
            Top             =   2685
            Width           =   1320
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Email(s) Notificacion"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   2160
            Width           =   1425
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Email Administrador"
            Height          =   195
            Left            =   45
            TabIndex        =   32
            Top             =   668
            Width           =   1365
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Puerto"
            Height          =   195
            Left            =   4080
            TabIndex        =   31
            Top             =   293
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Servidor SMTP"
            Height          =   195
            Left            =   45
            TabIndex        =   30
            Top             =   293
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Email Sistema"
            Height          =   195
            Left            =   45
            TabIndex        =   27
            Top             =   1043
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Email(s) Validación"
            Height          =   195
            Left            =   45
            TabIndex        =   26
            Top             =   1530
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Fra_Parametros 
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
      Height          =   4620
      Left            =   120
      TabIndex        =   86
      Top             =   360
      Width           =   6645
      Begin VB.TextBox Txt_Cambio_Calzado 
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   102
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox Txt_Ruta_Huellas 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2280
         TabIndex        =   100
         Top             =   3360
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox Txt_Ruta_Fotos 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2280
         TabIndex        =   99
         Top             =   2880
         Width           =   3975
      End
      Begin VB.CommandButton Btn_Seleccionar_Ruta_Huellas 
         Caption         =   "Seleccionar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   98
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txt_Parametros_Puerto_Correos 
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Top             =   2520
         Width           =   1880
      End
      Begin VB.TextBox Txt_Parametros_Contrasenia_Correos 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   2160
         Width           =   1880
      End
      Begin VB.TextBox Txt_Parametros_Servidor_Correos 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Txt_Parametros_Email_Correos 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox Chk_Imprime_Comidas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprime Comidas"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1245
         Width           =   2625
      End
      Begin VB.TextBox Txt_Costo_Comida_Empleado 
         Height          =   330
         Left            =   5355
         TabIndex        =   10
         Top             =   1717
         Width           =   960
      End
      Begin VB.TextBox Txt_Costo_Comida_Empresa 
         Height          =   330
         Left            =   1740
         TabIndex        =   9
         Top             =   1717
         Width           =   1080
      End
      Begin VB.CheckBox Chk_Aplica_Retardos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aplica Retardos"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   360
         Width           =   2625
      End
      Begin VB.CheckBox Chk_Calcula_Horas_Extra 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcula Horas Extra"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   780
         Width           =   2625
      End
      Begin VB.CommandButton Btn_Seleccionar_Ruta_Fotos 
         Caption         =   "Seleccionar"
         Height          =   330
         Left            =   1200
         TabIndex        =   15
         Top             =   2940
         Width           =   1125
      End
      Begin VB.TextBox Txt_Comidas_Diarias 
         Height          =   330
         Left            =   5355
         TabIndex        =   8
         Top             =   1185
         Width           =   960
      End
      Begin VB.TextBox Txt_Tolerancia_Retardos 
         Height          =   330
         Left            =   5355
         TabIndex        =   4
         Top             =   307
         Width           =   960
      End
      Begin VB.TextBox Txt_Horas_Maximas_Turno 
         Height          =   330
         Left            =   5355
         TabIndex        =   6
         Top             =   727
         Width           =   960
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         Caption         =   "Cambio de Calzado (meses)"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Lbl_Ruta_Huellas 
         BackColor       =   &H8000000E&
         Caption         =   "Ruta Huellas"
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_Puerto 
         BackColor       =   &H80000014&
         Caption         =   "Puerto"
         Height          =   255
         Left            =   3540
         TabIndex        =   96
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Lbl_Servidor 
         BackColor       =   &H80000014&
         Caption         =   "Servidor"
         Height          =   255
         Left            =   210
         TabIndex        =   95
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Lbl_Contraseña 
         BackColor       =   &H80000014&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   3540
         TabIndex        =   94
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Lbl_Email 
         BackColor       =   &H80000014&
         Caption         =   "Email"
         Height          =   255
         Left            =   210
         TabIndex        =   93
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Lbl_Costo_comida_Empleado 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "$ Comida Empleado"
         Height          =   195
         Left            =   3540
         TabIndex        =   92
         Top             =   1785
         Width           =   1410
      End
      Begin VB.Label Lbl_Costo_Comida_Empresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "$ Comida Empresa"
         Height          =   195
         Left            =   225
         TabIndex        =   91
         Top             =   1785
         Width           =   1320
      End
      Begin VB.Label Lbl_Horas_Maxima_Turno 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Horas Máximas Turno"
         Height          =   195
         Left            =   3540
         TabIndex        =   90
         Top             =   795
         Width           =   1545
      End
      Begin VB.Label Lbl_Tolerancia_Retardos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tolerancia (min.)"
         Height          =   195
         Left            =   3540
         TabIndex        =   89
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Lbl_Comidas_Diarias 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comidas por Dia"
         Height          =   195
         Left            =   3540
         TabIndex        =   88
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Lbl_Ruta_Fotos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ruta Fotos"
         Height          =   195
         Left            =   210
         TabIndex        =   87
         Top             =   3015
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARÁMETROS"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6690
   End
End
Attribute VB_Name = "Frm_Cat_Parametros_RH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'Constantes
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260 ' Para Buffer de caracteres del path
'Funcion Api CoTaskMemFree
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'Funcion Api CoTaskMemFree lstrcat
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Funcion Api SHBrowseForFolder
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Funcion Api SHGetPathFromIDList
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Nuevo" Or Btn_Modificar.Caption = "Modificar" Then
        If Btn_Modificar.Caption = "Nuevo" Then
            Btn_Modificar.Caption = "Alta"
        Else
            Btn_Modificar.Caption = "Actualizar"
        End If
        Btn_Salir.Caption = "Regresar"
        Fra_Parametros_Incidencias.Enabled = True
        Fra_Parametros_Sistema.Enabled = True
        Fra_Parametros.Enabled = True
        Tab_Parametros.Tab = 0
        Txt_Parametros_Servidor_SMTP.SetFocus
    Else
        If Valida_Parametros = True Then
            Modifica_Parametros
        End If
    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Parametros
'DESCRIPCION: Modifica los valores que se tienen en los registros de la tabla Cat_Parametros
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Parametros()
Dim Rs_Alta_Cat_Parametros_Sistema As rdoResultset 'Manejo del registro de Cat_Parametros_Sistema
Dim Rs_Modifica_Cat_Parametros As rdoResultset     'Modifica los datos de los registros que tiene la tabla
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    If Btn_Modificar.Caption = "Alta" Then
        'Alta de Parametros sistema
        Set Rs_Alta_Cat_Parametros_Sistema = Conectar_Ayudante.Recordset_Agregar("Cat_Parametros")
        'Llena la tabla de Cat_Parametros sistema con los datos contenidos en las cajas de textos
        With Rs_Alta_Cat_Parametros_Sistema
            .AddNew
                .rdoColumns("Parametro_ID") = "00001"
                'Parametros
                .rdoColumns("Aplica_Retardos") = Chk_Aplica_Retardos.Value
                .rdoColumns("Tolerancia_Retardos") = Val(Txt_Tolerancia_Retardos.Text)
                .rdoColumns("Calcula_Horas_Extra") = Chk_Calcula_Horas_Extra.Value
                .rdoColumns("Horas_Maximas_Turno") = Val(Txt_Horas_Maximas_Turno.Text)
                .rdoColumns("Minutos_Tolerancia") = Chk_Imprime_Comidas.Value
                .rdoColumns("Comidas_Diarias") = Val(Txt_Comidas_Diarias.Text)
                .rdoColumns("Costo_Comida_Empresa") = Val(Txt_Costo_Comida_Empresa.Text)
                .rdoColumns("Costo_Comida_Empleado") = Val(Txt_Costo_Comida_Empleado.Text)
                .rdoColumns("Ruta_Fotos") = Trim(Txt_Ruta_Fotos.Text)
                .rdoColumns("Ruta_Huellas") = Trim(Txt_Ruta_Huellas.Text)
                .rdoColumns("Email_Correo") = Trim(Txt_Parametros_Email_Correos.Text)
                .rdoColumns("Servidor_Correos") = Trim(Txt_Parametros_Servidor_Correos.Text)
                .rdoColumns("Contrasenia_Correos") = Trim(Txt_Parametros_Contrasenia_Correos.Text)
                .rdoColumns("Puerto_Correos") = Trim(Txt_Parametros_Puerto_Correos.Text)
                
                'Otros
                .rdoColumns("Edad_Minima_Contratacion") = Trim(Txt_Parametros_Edad_Minima_Contratacion.Text)
                .rdoColumns("Aviso_Contratacion") = Trim(Txt_Parametros_Aviso_Contratacion.Text)
                .rdoColumns("Email_Validacion") = Trim(Txt_Parametros_Email.Text)
                .rdoColumns("Email_Sistema") = Trim(Txt_Parametros_Email_Sistema.Text)
                .rdoColumns("Email_Administrador") = Trim(Txt_Parametros_Email_Administrador.Text)
                .rdoColumns("Hora_Importacion") = Format(Dtp_Parametros_Hora_Importacion.Value, "HH:mm")
                .rdoColumns("Horas_Dobles") = Val(Txt_Parametros_Pago_horas_Dobles.Text)
                .rdoColumns("Horas_Triples") = Val(Txt_Parametros_Pago_horas_Triples)
                .rdoColumns("Dias_Falta") = Val(Txt_Parametros_Retardos_Dias_Falta.Text)
                .rdoColumns("Periodo_Retardos_Dias") = Val(Txt_Parametros_Retardos_Periodo_Dias.Text)
                .rdoColumns("Servidor_SMTP") = Trim(Txt_Parametros_Servidor_SMTP.Text)
                .rdoColumns("Puerto_SMTP") = Val(Txt_Parametros_Puerto_Smtp.Text)
                .rdoColumns("Email_Notificacion") = Trim(Txt_Parametros_Email_Notificacion)
                .rdoColumns("Hora_Importacion_Dia") = Format(Dtp_Parametros_Hora_Importacion_Dia.Value, "HH:mm")
                .rdoColumns("PDF_Horas_Dobles") = Trim(Txt_Parametros_PDF_Horas_Dobles.Text)
                .rdoColumns("PDF_Horas_Triples") = Trim(Txt_Parametros_PDF_Horas_Triples.Text)
                .rdoColumns("PDF_Enfermedad_General") = Trim(Txt_Parametros_PDF_Enfermedad_General.Text)
                .rdoColumns("PDF_Maternidad") = Trim(Txt_Parametros_PDF_Maternidad.Text)
                .rdoColumns("PDF_Riesgo_Trabajo") = Trim(Txt_Parametros_PDF_Riesgo_Trabajo.Text)
                .rdoColumns("PDF_Vacaciones") = Trim(Txt_Parametros_PDF_Vacaciones.Text)
                .rdoColumns("PDF_Alumbramiento") = Trim(Txt_Parametros_PDF_Alumbramiento.Text)
                .rdoColumns("PDF_Defuncion") = Trim(Txt_Parametros_PDF_Defuncion.Text)
                .rdoColumns("PDF_Matrimonio") = Trim(Txt_Parametros_PDF_Matrimonio.Text)
                .rdoColumns("PDF_Falta_Justificada") = Trim(Txt_Parametros_PDF_Falta_Justificada.Text)
                .rdoColumns("PDF_Falta_InJustificada") = Trim(Txt_Parametros_PDF_Falta_InJustificada.Text)
                .rdoColumns("PDF_Permiso_Temporal") = Trim(Txt_Parametros_PDF_Permiso_Temporal)
                .rdoColumns("PDF_Permiso_Goce") = Trim(Txt_Parametros_PDF_Permiso_Goce.Text)
                .rdoColumns("PDF_Permiso_Sin_Goce") = Trim(Txt_Parametros_PDF_Permiso_Sin_Goce.Text)
                .rdoColumns("PDF_Sancion") = Trim(Txt_Parametros_PDF_Sancion.Text)
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
        Rs_Alta_Cat_Parametros_Sistema.Close
    Else
        Mi_SQL = "SELECT * FROM Cat_Parametros"
        Set Rs_Modifica_Cat_Parametros = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        'Llena la tabla de Cat_Parametros sistema con los datos contenidos en las cajas de textos
        With Rs_Modifica_Cat_Parametros
            .Edit
                'Parametros
                .rdoColumns("Aplica_Retardos") = Chk_Aplica_Retardos.Value
                .rdoColumns("Tolerancia_Retardos") = Val(Txt_Tolerancia_Retardos.Text)
                .rdoColumns("Calcula_Horas_Extra") = Chk_Calcula_Horas_Extra.Value
                .rdoColumns("Horas_Maximas_Turno") = Val(Txt_Horas_Maximas_Turno.Text)
                .rdoColumns("Minutos_Tolerancia") = Chk_Imprime_Comidas.Value
                .rdoColumns("Comidas_Diarias") = Val(Txt_Comidas_Diarias.Text)
                .rdoColumns("Costo_Comida_Empresa") = Val(Txt_Costo_Comida_Empresa.Text)
                .rdoColumns("Costo_Comida_Empleado") = Val(Txt_Costo_Comida_Empleado.Text)
                .rdoColumns("Ruta_Fotos") = Trim(Txt_Ruta_Fotos.Text)
                .rdoColumns("Ruta_Huellas") = Trim(Txt_Ruta_Huellas.Text)
                .rdoColumns("Email_Correo") = Trim(Txt_Parametros_Email_Correos.Text)
                .rdoColumns("Servidor_Correos") = Trim(Txt_Parametros_Servidor_Correos.Text)
                .rdoColumns("Contrasenia_Correos") = Trim(Txt_Parametros_Contrasenia_Correos.Text)
                .rdoColumns("Puerto_Correos") = Trim(Txt_Parametros_Puerto_Correos.Text)
                .rdoColumns("Meses_Cambio_Calzado") = Trim(Txt_Cambio_Calzado.Text)
                
                'Otros
                .rdoColumns("Edad_Minima_Contratacion") = Trim(Txt_Parametros_Edad_Minima_Contratacion.Text)
                .rdoColumns("Aviso_Contratacion") = Trim(Txt_Parametros_Aviso_Contratacion.Text)
                .rdoColumns("Email_Validacion") = Trim(Txt_Parametros_Email.Text)
                .rdoColumns("Email_Sistema") = Trim(Txt_Parametros_Email_Sistema.Text)
                .rdoColumns("Email_Administrador") = Trim(Txt_Parametros_Email_Administrador.Text)
                .rdoColumns("Hora_Importacion") = Format(Dtp_Parametros_Hora_Importacion.Value, "HH:mm")
                .rdoColumns("Horas_Dobles") = Val(Txt_Parametros_Pago_horas_Dobles.Text)
                .rdoColumns("Horas_Triples") = Val(Txt_Parametros_Pago_horas_Triples)
                .rdoColumns("Dias_Falta") = Val(Txt_Parametros_Retardos_Dias_Falta.Text)
                .rdoColumns("Periodo_Retardos_Dias") = Val(Txt_Parametros_Retardos_Periodo_Dias.Text)
                .rdoColumns("Servidor_SMTP") = Trim(Txt_Parametros_Servidor_SMTP.Text)
                .rdoColumns("Puerto_SMTP") = Val(Txt_Parametros_Puerto_Smtp.Text)
                .rdoColumns("Email_Notificacion") = Trim(Txt_Parametros_Email_Notificacion)
                .rdoColumns("Hora_Importacion_Dia") = Format(Dtp_Parametros_Hora_Importacion_Dia.Value, "HH:mm")
                .rdoColumns("PDF_Horas_Dobles") = Trim(Txt_Parametros_PDF_Horas_Dobles.Text)
                .rdoColumns("PDF_Horas_Triples") = Trim(Txt_Parametros_PDF_Horas_Triples.Text)
                .rdoColumns("PDF_Enfermedad_General") = Trim(Txt_Parametros_PDF_Enfermedad_General.Text)
                .rdoColumns("PDF_Maternidad") = Trim(Txt_Parametros_PDF_Maternidad.Text)
                .rdoColumns("PDF_Riesgo_Trabajo") = Trim(Txt_Parametros_PDF_Riesgo_Trabajo.Text)
                .rdoColumns("PDF_Vacaciones") = Trim(Txt_Parametros_PDF_Vacaciones.Text)
                .rdoColumns("PDF_Alumbramiento") = Trim(Txt_Parametros_PDF_Alumbramiento.Text)
                .rdoColumns("PDF_Defuncion") = Trim(Txt_Parametros_PDF_Defuncion.Text)
                .rdoColumns("PDF_Matrimonio") = Trim(Txt_Parametros_PDF_Matrimonio.Text)
                .rdoColumns("PDF_Falta_Justificada") = Trim(Txt_Parametros_PDF_Falta_Justificada.Text)
                .rdoColumns("PDF_Falta_InJustificada") = Trim(Txt_Parametros_PDF_Falta_InJustificada.Text)
                .rdoColumns("PDF_Permiso_Temporal") = Trim(Txt_Parametros_PDF_Permiso_Temporal.Text)
                .rdoColumns("PDF_Permiso_Goce") = Trim(Txt_Parametros_PDF_Permiso_Goce.Text)
                .rdoColumns("PDF_Permiso_Sin_Goce") = Trim(Txt_Parametros_PDF_Permiso_Sin_Goce.Text)
                .rdoColumns("PDF_Sancion") = Trim(Txt_Parametros_PDF_Sancion.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
        Rs_Modifica_Cat_Parametros.Close
    End If
    Conexion_Base.CommitTrans
    MsgBox "Parametros capturados", vbInformation
    Fra_Parametros_Incidencias.Enabled = False
    Fra_Parametros_Sistema.Enabled = False
    Tab_Parametros.Tab = 0
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Consulta_Parametros 'Consulta los parámetros del sistema para actualizarlos de acuerdo a la modificación realizada
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Parametros", Frm_Cat_Parametros_RH)
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        If MsgBox("¿Esta seguro de cancelar la operación?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Consulta_Parametros 'Consulta los valores que tiene los parámetros
            Fra_Parametros_Incidencias.Enabled = False
            Fra_Parametros_Sistema.Enabled = False
            Fra_Parametros.Enabled = False
            If Btn_Modificar.Caption = "Alta" Then
                Btn_Modificar.Caption = "Nuevo"
            Else
                Btn_Modificar.Caption = "Modificar"
            End If
            Btn_Salir.Caption = "Salir"
        End If
    End If
End Sub

Private Sub Btn_Seleccionar_Ruta_Fotos_Click()
Dim Ruta As String
    Ruta = Selecciona_Ruta_Directorio(Me, "Indique la carpeta donde se almacenará las fotos de los empleados")
    If Ruta <> "" Then
        Txt_Ruta_Fotos.Text = Ruta
    End If
End Sub

Private Sub Btn_Seleccionar_Ruta_Huellas_Click()
Dim Ruta As String
    Ruta = Selecciona_Ruta_Directorio(Me, "Indique la carpeta donde se almacenará las fotos de los empleados")
    If Ruta <> "" Then
        Txt_Ruta_Huellas.Text = Ruta
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Consulta_Parametros 'Consulta los parametros que tiene el sistema
    Llena_Parametros
    Btn_Modificar.Caption = "Modificar"
End Sub



Private Sub Txt_Cambio_Calzado_KeyPress(KeyAscii As Integer)
     Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cambio_Calzado.Text, False)
End Sub

Private Sub Txt_Comidas_Diarias_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Comidas_Diarias.Text, False)
End Sub

Private Sub Txt_Costo_Comida_Empleado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Costo_Comida_Empleado.Text, True)
End Sub

Private Sub Txt_Costo_Comida_Empresa_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Costo_Comida_Empresa.Text, True)
End Sub

Private Sub Txt_Horas_Maximas_Turno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Horas_Maximas_Turno.Text, False)
End Sub

Private Sub Txt_Parametros_Aviso_Contratacion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Aviso_Contratacion.Text, False)
End Sub

Private Sub Txt_Parametros_Clave_Falta_Retardos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_Edad_Minima_Contratacion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Edad_Minima_Contratacion.Text, True)
End Sub





Private Sub Txt_Parametros_Pago_horas_Dobles_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Pago_horas_Dobles.Text, True)
End Sub

Private Sub Txt_Parametros_Pago_horas_Triples_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Pago_horas_Triples.Text, True)
End Sub

Private Sub Txt_Parametros_PDF_Alumbramiento_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Defuncion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Enfermedad_General_KeyPress(KeyAscii As Integer)
Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub


Private Sub Txt_Parametros_PDF_Falta_InJustificada_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Falta_Justificada_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Horas_Dobles_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Horas_Triples_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Maternidad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Matrimonio_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Permiso_Goce_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Permiso_Sin_Goce_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Permiso_Temporal_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Riesgo_Trabajo_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_PDF_Vacaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Parametros_Puerto_Correos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Puerto_Correos.Text, False)
End Sub

Private Sub Txt_Parametros_Puerto_Smtp_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Puerto_Smtp, False)
End Sub

Private Sub Txt_Parametros_Retardos_Dias_Falta_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Retardos_Dias_Falta.Text, False)
End Sub

Private Sub Txt_Parametros_Retardos_Periodo_Dias_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Retardos_Periodo_Dias.Text, False)
End Sub

Private Sub Txt_Parametros_Retardos_Tolerancia_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Parametros_Retardos_Tolerancia.Text, False)
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Llena_Parametros
'DESCRIPCION: Coloca la informacion de los parametros en las cajas de texto
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Llena_Parametros()
Dim Rs_Consulta_Cat_Parametros As rdoResultset

On Error GoTo HANDLER
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Parametros
        If Not .EOF Then
            'Parametros
            Chk_Aplica_Retardos.Value = PG_Aplica_Retardos
            Txt_Tolerancia_Retardos.Text = PG_Tolerancia_Retardos
            Chk_Calcula_Horas_Extra.Value = PG_Calcula_Horas_Extra
            Txt_Horas_Maximas_Turno.Text = PG_Horas_Maximas_Turno
            Chk_Imprime_Comidas.Value = PG_Imprime_Comidas
            Txt_Comidas_Diarias.Text = PG_Cantidad_Comidas
            Txt_Costo_Comida_Empresa.Text = PG_Costo_Comida_Empresa
            Txt_Costo_Comida_Empleado.Text = PG_Costo_Comida_Empleado
            Txt_Ruta_Fotos.Text = PG_Ruta_Fotos
            If Not IsNull(.rdoColumns("Ruta_Huellas")) Then
            Txt_Ruta_Huellas.Text = .rdoColumns("Ruta_Huellas")
            
            End If
             If Not IsNull(.rdoColumns("Meses_Cambio_Calzado")) Then
            Txt_Cambio_Calzado.Text = .rdoColumns("Meses_Cambio_Calzado")
            Else
             Txt_Cambio_Calzado.Text = 0
            End If
            'Otros
            Txt_Parametros_Edad_Minima_Contratacion.Text = Trim(.rdoColumns("Edad_Minima_Contratacion"))
            Txt_Parametros_Aviso_Contratacion.Text = Trim(.rdoColumns("Aviso_Contratacion"))
            Txt_Parametros_Email.Text = .rdoColumns("Email_validacion")
            Txt_Parametros_Email_Sistema.Text = .rdoColumns("Email_Sistema")
            Txt_Parametros_Email_Administrador.Text = .rdoColumns("Email_Administrador")
            Txt_Parametros_Pago_horas_Dobles.Text = .rdoColumns("Horas_Dobles")
            Txt_Parametros_Pago_horas_Triples = .rdoColumns("Horas_Triples")
            Txt_Parametros_Retardos_Dias_Falta.Text = .rdoColumns("Dias_Falta")
            Txt_Parametros_Retardos_Periodo_Dias.Text = .rdoColumns("Periodo_Retardos_Dias")
            Txt_Parametros_Retardos_Tolerancia.Text = .rdoColumns("Minutos_Tolerancia")
            Txt_Parametros_Servidor_SMTP.Text = .rdoColumns("Servidor_SMTP")
            Txt_Parametros_Puerto_Smtp.Text = Val(.rdoColumns("Puerto_SMTP"))
            Dtp_Parametros_Hora_Importacion.Value = .rdoColumns("Hora_Importacion")
            Txt_Parametros_Email_Notificacion.Text = .rdoColumns("Email_Notificacion")
            Dtp_Parametros_Hora_Importacion_Dia.Value = .rdoColumns("Hora_Importacion_Dia")
            Txt_Parametros_PDF_Horas_Dobles.Text = .rdoColumns("PDF_Horas_Dobles")
            Txt_Parametros_PDF_Horas_Triples.Text = .rdoColumns("PDF_Horas_Triples")
            Txt_Parametros_PDF_Enfermedad_General.Text = .rdoColumns("PDF_Enfermedad_General")
            Txt_Parametros_PDF_Maternidad.Text = .rdoColumns("PDF_Maternidad")
            Txt_Parametros_PDF_Riesgo_Trabajo.Text = .rdoColumns("PDF_Riesgo_Trabajo")
            Txt_Parametros_PDF_Vacaciones.Text = .rdoColumns("PDF_Vacaciones")
            Txt_Parametros_PDF_Alumbramiento.Text = .rdoColumns("PDF_Alumbramiento")
            Txt_Parametros_PDF_Defuncion.Text = .rdoColumns("PDF_Defuncion")
            Txt_Parametros_PDF_Matrimonio.Text = .rdoColumns("PDF_Matrimonio")
            Txt_Parametros_PDF_Falta_Justificada.Text = .rdoColumns("PDF_Falta_Justificada")
            Txt_Parametros_PDF_Falta_InJustificada.Text = .rdoColumns("PDF_Falta_InJustificada")
            Txt_Parametros_PDF_Permiso_Temporal = .rdoColumns("PDF_Permiso_Temporal")
            Txt_Parametros_PDF_Permiso_Goce.Text = Trim(.rdoColumns("PDF_Permiso_Goce"))
            Txt_Parametros_PDF_Permiso_Sin_Goce.Text = Trim(.rdoColumns("PDF_Permiso_Sin_Goce"))
            Txt_Parametros_PDF_Sancion.Text = Trim(.rdoColumns("PDF_Sancion"))
            
            If Not IsNull(.rdoColumns("Email_Correo")) Then
             Txt_Parametros_Email_Correos.Text = Trim(.rdoColumns("Email_Correo"))
            
            End If
            If Not IsNull(.rdoColumns("Contrasenia_Correos")) Then
            
            Txt_Parametros_Contrasenia_Correos.Text = (.rdoColumns("Contrasenia_Correos"))
            End If
            If Not IsNull(.rdoColumns("Servidor_Correos")) Then
            Txt_Parametros_Servidor_Correos.Text = Trim(.rdoColumns("Servidor_Correos"))
           End If
           If Not IsNull(.rdoColumns("Puerto_Correos")) Then
           Txt_Parametros_Puerto_Correos.Text = Trim(.rdoColumns("Puerto_Correos"))
        End If
        End If
    End With
    Set Rs_Consulta_Cat_Parametros = Nothing
Exit Sub
HANDLER:
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Validar_Parametros
'DESCRIPCION: Valida que los parametros se cumplan
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Function Valida_Parametros() As Boolean
    Valida_Parametros = False
    If Trim(Txt_Parametros_Servidor_SMTP) = "" Then
        MsgBox "Proporcione la direccion del servidor SMTP", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Servidor_SMTP.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Puerto_Smtp.Text) <= 0 Then
        MsgBox "Proporcione el puerto para el servidor smtp", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Puerto_Smtp.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_Email_Administrador.Text) = "" Then
        MsgBox "Proporcione la direccion de correo electronico del administrador", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Email_Administrador.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_Email_Sistema.Text) = "" Then
        MsgBox "Proporcione la direccion de correo electronico asignada al sistema", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Email_Sistema.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_Email.Text) = "" Then
        MsgBox "Proporcione la(s) direccion(es) de correo electronico" + vbCrLf + _
                       "a quien se enviara la información validación", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Email.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_Email_Notificacion.Text) = "" Then
        MsgBox "Proporcione la(s) direccion(es) de correo electronico" + vbCrLf + _
                   "a quien se enviara la notificación diaria", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Email_Notificacion.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Edad_Minima_Contratacion.Text) < 0 Then
        MsgBox "Proporcione la edad minima de contratación", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Edad_Minima_Contratacion.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Aviso_Contratacion.Text) < 0 Then
        MsgBox "Proporcione los dias para avisar antes del término del contrato", vbExclamation
        Tab_Parametros.Tab = 0
        Txt_Parametros_Aviso_Contratacion.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Pago_horas_Dobles.Text) < 0 Then
        MsgBox "Proporcione el No. de Horas para las horas Dobles", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_Pago_horas_Dobles.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Horas_Dobles.Text) = "" Then
        MsgBox "Proporcione la clave para las Horas Dobles", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Horas_Dobles.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Pago_horas_Triples.Text) < 0 Then
        MsgBox "Proporcione el No. de Horas para las horas Triples", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Horas_Triples.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Horas_Triples.Text) = "" Then
        MsgBox "Proporcione la clave para las Horas Triples", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Horas_Triples.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Retardos_Dias_Falta.Text) < 0 Then
        MsgBox "Proporcione el no. de retardos para generar una falta", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_Retardos_Dias_Falta.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_Retardos_Tolerancia.Text) = "" Then
        MsgBox "Proporcione el no. de minutos de tolerancia para retardos", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_Retardos_Tolerancia.SetFocus
        Exit Function
    End If
    If Val(Txt_Parametros_Retardos_Periodo_Dias.Text) < 0 Then
        MsgBox "Proporcione el no. de dias del periodo para contabilizar retardos", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_Retardos_Periodo_Dias.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Enfermedad_General.Text) = "" Then
        MsgBox "Proporcione la clave para Enfermedad General", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Enfermedad_General.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Maternidad.Text) = "" Then
        MsgBox "Proporcione la clave para Maternidad", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Maternidad.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Riesgo_Trabajo.Text) = "" Then
        MsgBox "Proporcione la clave para Riesgo Trabajo", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Riesgo_Trabajo.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Vacaciones.Text) = "" Then
        MsgBox "Proporcione la clave para Vacaciones", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Vacaciones.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Permiso_Temporal.Text) = "" Then
        MsgBox "Proporcione la clave para Permiso Temporal", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Permiso_Temporal.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Alumbramiento.Text) = "" Then
        MsgBox "Proporcione la clave para Alumbramiento", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Alumbramiento.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Defuncion.Text) = "" Then
        MsgBox "Proporcione la clave para Defuncion", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Defuncion.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Matrimonio.Text) = "" Then
        MsgBox "Proporcione la clave para Matrimonio", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Matrimonio.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Falta_Justificada.Text) = "" Then
        MsgBox "Proporcione la clave para Falta Justificada", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Falta_Justificada.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Falta_InJustificada.Text) = "" Then
        MsgBox "Proporcione la clave para Falta Injustificada", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Falta_InJustificada.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Permiso_Goce.Text) = "" Then
        MsgBox "Proporcione la clave para Permisos con goce de sueldo", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Permiso_Goce.SetFocus
        Exit Function
    End If
    If Trim(Txt_Parametros_PDF_Permiso_Sin_Goce.Text) = "" Then
        MsgBox "Proporcione la clave para Permisos sin goce de sueldo", vbExclamation
        Tab_Parametros.Tab = 1
        Txt_Parametros_PDF_Permiso_Sin_Goce.SetFocus
        Exit Function
    End If
    Valida_Parametros = True
End Function

Private Sub Txt_Tolerancia_Retardos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Tolerancia_Retardos.Text, False)
End Sub
