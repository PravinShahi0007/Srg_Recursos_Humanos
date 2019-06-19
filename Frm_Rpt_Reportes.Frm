VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_Rpt_Reportes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTES"
   ClientHeight    =   7695
   ClientLeft      =   2895
   ClientTop       =   1875
   ClientWidth     =   13155
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Frm_Rpt_Reportes.frx":0000
   ScaleHeight     =   7695
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox Edm_Reporte 
      Height          =   6915
      Left            =   45
      TabIndex        =   69
      Top             =   30
      Visible         =   0   'False
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   12197
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      RightMargin     =   20000
      TextRTF         =   $"Frm_Rpt_Reportes.frx":0442
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
   Begin VB.PictureBox Pic_Reportes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4005
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   9270
      TabIndex        =   4
      Top             =   -15
      Width           =   9270
      Begin VB.Frame Fra_Historico_Facturas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Historico de Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   42
         Top             =   30
         Visible         =   0   'False
         Width           =   4830
         Begin VB.CheckBox Chk_Sin_Presupuesto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sin Presupuesto"
            Height          =   195
            Left            =   135
            TabIndex        =   246
            Top             =   2340
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CommandButton Btn_Reporte_Historico_Facturas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1755
            TabIndex        =   52
            Top             =   2235
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Historico_Factura 
            Height          =   315
            Left            =   1305
            TabIndex        =   51
            Top             =   1635
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Historico_Factura 
            Height          =   315
            Left            =   1305
            TabIndex        =   50
            Top             =   1140
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin VB.Frame Fra_Cliente_Factura 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   45
            TabIndex        =   105
            Top             =   330
            Visible         =   0   'False
            Width           =   4695
            Begin VB.ComboBox Cmb_Cliente_Factura 
               Height          =   315
               Left            =   60
               TabIndex        =   49
               Top             =   345
               Width           =   4560
            End
            Begin VB.Label Lbl_Cliente_Factura 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cliente"
               Height          =   195
               Left            =   120
               TabIndex        =   106
               Top             =   15
               Width           =   480
            End
         End
         Begin VB.Frame Fra_Estatus_Facturas 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   750
            Left            =   195
            TabIndex        =   104
            Top             =   285
            Visible         =   0   'False
            Width           =   4590
            Begin VB.CheckBox Chk_Facturas_Pendientes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   3225
               TabIndex        =   47
               Top             =   75
               Width           =   1095
            End
            Begin VB.CheckBox Chk_Facturas_Canceladas 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Canceladas"
               Height          =   195
               Left            =   1500
               TabIndex        =   46
               Top             =   75
               Width           =   1230
            End
            Begin VB.CheckBox Chk_Facturas_Pagadas 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pagadas"
               Height          =   195
               Left            =   15
               TabIndex        =   45
               Top             =   75
               Width           =   990
            End
            Begin VB.ComboBox Cmb_Empresa_Facturas 
               Height          =   315
               ItemData        =   "Frm_Rpt_Reportes.frx":04C2
               Left            =   1110
               List            =   "Frm_Rpt_Reportes.frx":04D5
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   405
               Width           =   3360
            End
            Begin VB.Label Lbl_Empresa_Factura 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Empresa"
               Height          =   195
               Left            =   0
               TabIndex        =   107
               Top             =   465
               Width           =   615
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   210
            TabIndex        =   44
            Top             =   1695
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   210
            TabIndex        =   43
            Top             =   1200
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Pedidos_Facturas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Facturas por pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   45
         TabIndex        =   135
         Top             =   15
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CheckBox Chk_Pedido_Sin_Factura 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pedido sin Factura/Remision"
            Height          =   210
            Left            =   1080
            TabIndex        =   138
            Top             =   750
            Value           =   1  'Checked
            Width           =   2865
         End
         Begin VB.CheckBox Chk_Pedido_Facturado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pedido Facturado/Remisionado"
            Height          =   210
            Left            =   1080
            TabIndex        =   137
            Top             =   405
            Value           =   1  'Checked
            Width           =   2865
         End
         Begin VB.CommandButton Btn_Reporte_Pedido_Factura 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1620
            TabIndex        =   141
            Top             =   2235
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Pedido_Factura 
            Height          =   315
            Left            =   1080
            TabIndex        =   140
            Top             =   1620
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Pedido_Factura 
            Height          =   315
            Left            =   1080
            TabIndex        =   139
            Top             =   1125
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39536
         End
         Begin VB.Label Lbl_Fecha_Fin_Pedido_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   142
            Top             =   1650
            Width           =   705
         End
         Begin VB.Label Lbl_Fecha_Inicio_Pedido_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   136
            Top             =   1170
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Kardex 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kardex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   5850
         Begin VB.ComboBox Cmb_Producto 
            Height          =   315
            Left            =   90
            TabIndex        =   0
            Top             =   615
            Width           =   5610
         End
         Begin MSComCtl2.DTPicker Dtp_Kardex_Fecha_Final 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   1680
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39448
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker Dtp_Kardex_Fecha_Inicial 
            Height          =   315
            Left            =   1440
            TabIndex        =   1
            Top             =   1200
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39448
            MinDate         =   2
         End
         Begin VB.CommandButton Btn_Kardex_Reporte 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1965
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Lbl_Kardex_Fecha_Final 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   285
            TabIndex        =   32
            Top             =   1740
            Width           =   825
         End
         Begin VB.Label Lbl_Kardex_Fecha_Inicial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   285
            TabIndex        =   31
            Top             =   1275
            Width           =   900
         End
         Begin VB.Label Lbl_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   285
            TabIndex        =   30
            Top             =   330
            Width           =   645
         End
      End
      Begin VB.Frame Fra_Pedidos_Envios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Envios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   160
         Top             =   30
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CheckBox Chk_Fecha_Pedido 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Pedido"
            Height          =   225
            Left            =   90
            TabIndex        =   175
            Top             =   1950
            Width           =   1305
         End
         Begin VB.TextBox Txt_Lugar_Envio 
            Height          =   315
            Left            =   5490
            TabIndex        =   171
            Top             =   1080
            Width           =   2685
         End
         Begin VB.CheckBox Chk_Lugar_Envio 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lugar"
            Height          =   360
            Left            =   4170
            TabIndex        =   170
            Top             =   1050
            Width           =   720
         End
         Begin VB.CheckBox Chk_Fecha_Envio 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Envio"
            Height          =   270
            Left            =   90
            TabIndex        =   172
            Top             =   1507
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_Medio_Envio 
            Height          =   315
            Left            =   1395
            TabIndex        =   169
            Top             =   1080
            Width           =   2685
         End
         Begin VB.CheckBox Chk_Medio_Envio 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Medio"
            Height          =   405
            Left            =   90
            TabIndex        =   168
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox Txt_No_Factura_Envio 
            Height          =   315
            Left            =   5490
            TabIndex        =   167
            Top             =   690
            Width           =   2685
         End
         Begin VB.CheckBox Chk_No_Factura_Envio 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Factura"
            Height          =   300
            Left            =   4185
            TabIndex        =   166
            Top             =   690
            Width           =   1140
         End
         Begin VB.TextBox Txt_No_Pedido_Envio 
            Height          =   315
            Left            =   1395
            TabIndex        =   165
            Top             =   690
            Width           =   2685
         End
         Begin VB.CheckBox Chk_No_Pedido_Envio 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Pedido"
            Height          =   210
            Left            =   90
            TabIndex        =   164
            Top             =   735
            Width           =   1095
         End
         Begin VB.ComboBox Cmb_Cliente_Envio 
            Height          =   315
            Left            =   1395
            TabIndex        =   163
            Top             =   285
            Width           =   6795
         End
         Begin VB.CheckBox Chk_Cliente 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   270
            Left            =   90
            TabIndex        =   162
            Top             =   300
            Width           =   900
         End
         Begin VB.CommandButton Btn_Reporte_Envios 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   3465
            TabIndex        =   161
            Top             =   2460
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Envio 
            Height          =   315
            Left            =   1395
            TabIndex        =   173
            Top             =   1485
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39545
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Envio 
            Height          =   315
            Left            =   1395
            TabIndex        =   176
            Top             =   1905
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39543
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Envio_Fin 
            Height          =   315
            Left            =   5490
            TabIndex        =   174
            Top             =   1485
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39545
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Envio 
            Height          =   315
            Left            =   5490
            TabIndex        =   177
            Top             =   1905
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39545
         End
         Begin VB.Label Lbl_Pedido 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   4500
            TabIndex        =   191
            Top             =   1965
            Width           =   135
         End
         Begin VB.Label Lbl_Envio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   4500
            TabIndex        =   190
            Top             =   1545
            Width           =   135
         End
      End
      Begin VB.Frame Fra_Saldos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Saldo por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   143
         Top             =   30
         Visible         =   0   'False
         Width           =   4845
         Begin VB.ComboBox Cmb_Empresa_Saldo 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0516
            Left            =   1035
            List            =   "Frm_Rpt_Reportes.frx":0523
            Style           =   2  'Dropdown List
            TabIndex        =   244
            Top             =   300
            Width           =   3720
         End
         Begin VB.ComboBox Cmb_Remision_Factura 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0548
            Left            =   1545
            List            =   "Frm_Rpt_Reportes.frx":0555
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   1695
            Width           =   3195
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Saldo_Inicio 
            Height          =   315
            Left            =   1545
            TabIndex        =   148
            Top             =   2070
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd MMM yyyy"
            Format          =   60751875
            CurrentDate     =   39885
         End
         Begin VB.CheckBox Chk_Fechas_Saldos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Por Fechas"
            Height          =   195
            Left            =   120
            TabIndex        =   147
            Top             =   2130
            Width           =   1365
         End
         Begin VB.ComboBox Cmb_Vendedor_Saldo 
            Height          =   315
            Left            =   1035
            TabIndex        =   144
            Top             =   705
            Width           =   3720
         End
         Begin VB.CommandButton Btn_Reporte_Saldos 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1620
            TabIndex        =   150
            Top             =   2475
            Width           =   1350
         End
         Begin VB.ComboBox Cmb_Cliente_Saldo 
            Height          =   315
            Left            =   90
            TabIndex        =   145
            Top             =   1320
            Width           =   4665
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Saldo_Fin 
            Height          =   315
            Left            =   3165
            TabIndex        =   149
            Top             =   2070
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd MMM yyyy"
            Format          =   60751875
            CurrentDate     =   39885
         End
         Begin VB.Label Lbl_Empresa_Saldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   135
            TabIndex        =   245
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Lbl_Remision_Factura 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Documento"
            Height          =   195
            Left            =   135
            TabIndex        =   233
            Top             =   1755
            Width           =   825
         End
         Begin VB.Label Lbl_Vendedor_Saldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   135
            TabIndex        =   189
            Top             =   765
            Width           =   690
         End
         Begin VB.Label Lbl_Cliente_Saldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   135
            TabIndex        =   151
            Top             =   1065
            Width           =   480
         End
      End
      Begin VB.Frame Fra_Reporte_Notas_Credito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Notas de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   126
         Top             =   45
         Visible         =   0   'False
         Width           =   4860
         Begin VB.ComboBox Cmb_Tipo_Nota_Credito 
            Height          =   315
            Left            =   1020
            TabIndex        =   130
            Top             =   1365
            Width           =   3705
         End
         Begin VB.ComboBox Cmb_Empresa_NC 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0573
            Left            =   1020
            List            =   "Frm_Rpt_Reportes.frx":0580
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   960
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker Dtp_Nota_Credito_Fecha_Final 
            Height          =   315
            Left            =   2895
            TabIndex        =   132
            Top             =   1755
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39526
         End
         Begin MSComCtl2.DTPicker Dtp_Nota_Credito_Fecha_Inicio 
            Height          =   315
            Left            =   1020
            TabIndex        =   131
            Top             =   1755
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39526
         End
         Begin VB.ComboBox Cmb_Nota_Credito_Cliente 
            Height          =   315
            Left            =   75
            TabIndex        =   128
            Top             =   540
            Width           =   4650
         End
         Begin VB.CommandButton Btn_Reporte_Notas_Credito_Clientes 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1575
            TabIndex        =   133
            Top             =   2265
            Width           =   1590
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo NC"
            Height          =   195
            Left            =   150
            TabIndex        =   236
            Top             =   1425
            Width           =   585
         End
         Begin VB.Label Lbl_Empresa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   150
            TabIndex        =   204
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Lbl_Nota_Credito_ 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   150
            TabIndex        =   134
            Top             =   1815
            Width           =   525
         End
         Begin VB.Label Lbl_Nota_Credito_Cliente 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   150
            TabIndex        =   127
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame Fra_Reporte_Entradas_Salidas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entradas y Salidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   5835
         Begin VB.ComboBox Cmb_Proveedor_Entrada 
            Height          =   315
            Left            =   1110
            TabIndex        =   8
            Top             =   930
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ComboBox Cmb_Tipo_Entrada 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":05B1
            Left            =   1110
            List            =   "Frm_Rpt_Reportes.frx":05B3
            TabIndex        =   9
            Top             =   1350
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.TextBox Txt_Referencia 
            Height          =   300
            Left            =   1125
            TabIndex        =   7
            Top             =   510
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.CommandButton Btn_Reporte_Entradas_Salidas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2265
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   2430
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Entrada_Salida_Inicio 
            Height          =   330
            Left            =   1125
            TabIndex        =   10
            Top             =   1815
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Entrada_Salida_Fin 
            Height          =   330
            Left            =   3435
            TabIndex        =   11
            Top             =   1815
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Proveedor_Entrada 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   210
            TabIndex        =   240
            Top             =   990
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Lbl_Tipo_Entrda 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo"
            Height          =   195
            Left            =   210
            TabIndex        =   239
            Top             =   1410
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Lbl_Referencia 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Referencia"
            Height          =   195
            Left            =   210
            TabIndex        =   238
            Top             =   570
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Lbl_Fecha_Inicio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   210
            TabIndex        =   29
            Top             =   1890
            Width           =   525
         End
      End
      Begin VB.Frame Fra_Facturas_Vencidas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas Vencidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   45
         TabIndex        =   85
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.ComboBox Cmb_Empresa_Factura_Vencida 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":05B5
            Left            =   1005
            List            =   "Frm_Rpt_Reportes.frx":05C2
            TabIndex        =   86
            Top             =   660
            Width           =   3645
         End
         Begin MSComCtl2.DTPicker Dtp_Facturas_Vencidas 
            Height          =   315
            Left            =   1005
            TabIndex        =   87
            Top             =   1200
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39500
         End
         Begin VB.CommandButton Btn_Reporte_Facturas_Vencidas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1725
            TabIndex        =   88
            Top             =   1905
            Width           =   1500
         End
         Begin VB.Label Lbl_Fecha_Factura_Vencida 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   195
            TabIndex        =   179
            Top             =   1260
            Width           =   450
         End
         Begin VB.Label Lbl_Empresa_Factura_Vencida 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   195
            TabIndex        =   178
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Fra_Consulta_Cobranza 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consulta de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   108
         Top             =   30
         Visible         =   0   'False
         Width           =   4845
         Begin VB.ComboBox Cmb_Documento_Cobranza 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":05F3
            Left            =   1050
            List            =   "Frm_Rpt_Reportes.frx":0600
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   315
            Width           =   3675
         End
         Begin VB.ComboBox Cmb_Empresa_Cobranza 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":061E
            Left            =   1050
            List            =   "Frm_Rpt_Reportes.frx":0620
            TabIndex        =   110
            Top             =   780
            Width           =   3675
         End
         Begin VB.CommandButton Btn_Reporte_Cobranza 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1785
            TabIndex        =   113
            Top             =   2265
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Cobranza 
            Height          =   315
            Left            =   1050
            TabIndex        =   112
            Top             =   1785
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39519
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Cobranza 
            Height          =   315
            Left            =   1050
            TabIndex        =   111
            Top             =   1260
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39519
         End
         Begin VB.Label Lbl_Documento_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Documento"
            Height          =   195
            Left            =   90
            TabIndex        =   235
            Top             =   375
            Width           =   825
         End
         Begin VB.Label Lbl_Empresa_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   234
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Lbl_Fecha_Inicio_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   115
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Lbl_Fecha_Fin_Cobranza 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   114
            Top             =   1845
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Comisiones_Vendedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comisiones Vendedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   45
         TabIndex        =   116
         Top             =   45
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CheckBox Chk_Detalle_Vendededor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detallado"
            Height          =   195
            Left            =   135
            TabIndex        =   122
            Top             =   2490
            Width           =   1335
         End
         Begin VB.ComboBox Cmb_Tipo_Documento 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0622
            Left            =   1185
            List            =   "Frm_Rpt_Reportes.frx":062C
            TabIndex        =   119
            Top             =   1005
            Width           =   3555
         End
         Begin VB.CommandButton Btn_Reporte_Comisiones 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1590
            TabIndex        =   123
            Top             =   2400
            Width           =   1665
         End
         Begin MSComCtl2.DTPicker Dtp_Comision_Fecha_Fin 
            Height          =   315
            Left            =   1185
            TabIndex        =   121
            Top             =   1845
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39525
         End
         Begin MSComCtl2.DTPicker Dtp_Comision_Fecha_Inicio 
            Height          =   315
            Left            =   1185
            TabIndex        =   120
            Top             =   1410
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39525
         End
         Begin VB.ComboBox Cmb_Vendedor_Comision 
            Height          =   315
            Left            =   75
            TabIndex        =   118
            Top             =   600
            Width           =   4665
         End
         Begin VB.Label Lbl_Documento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Documento"
            Height          =   195
            Left            =   90
            TabIndex        =   231
            Top             =   1065
            Width           =   825
         End
         Begin VB.Label Lbl_Comision_Fecha_Final 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   125
            Top             =   1905
            Width           =   705
         End
         Begin VB.Label Lbl_Comision_Fecha_Inicio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   124
            Top             =   1470
            Width           =   870
         End
         Begin VB.Label Lbl_Vendedor_Comision 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   90
            TabIndex        =   117
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Frame Fra_Reporte_Gastos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Gastos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   45
         TabIndex        =   89
         Top             =   30
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox Cmb_Tipo_Pago 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0646
            Left            =   1260
            List            =   "Frm_Rpt_Reportes.frx":0650
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   510
            Width           =   3240
         End
         Begin VB.ComboBox Cmb_Banco 
            Height          =   315
            Left            =   1260
            TabIndex        =   91
            Top             =   1035
            Width           =   3240
         End
         Begin VB.TextBox Txt_No_Cheque 
            Height          =   315
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   92
            Top             =   1575
            Width           =   3240
         End
         Begin VB.CommandButton Btn_Reporte_Gastos 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   93
            Top             =   2220
            Width           =   1425
         End
         Begin VB.Label Lbl_Tipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo de Pago"
            Height          =   195
            Left            =   150
            TabIndex        =   211
            Top             =   570
            Width           =   960
         End
         Begin VB.Label Lbl_Banco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Banco"
            Height          =   195
            Left            =   150
            TabIndex        =   210
            Top             =   1095
            Width           =   465
         End
         Begin VB.Label Lbl_No_Cheque 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Referencia"
            Height          =   195
            Left            =   150
            TabIndex        =   209
            Top             =   1635
            Width           =   780
         End
      End
      Begin VB.Frame Fra_Comparativo_Ventas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comparativo de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   213
         Top             =   30
         Visible         =   0   'False
         Width           =   5835
         Begin VB.ComboBox Cmb_Tipo_Reporte 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":066B
            Left            =   1380
            List            =   "Frm_Rpt_Reportes.frx":0675
            Style           =   2  'Dropdown List
            TabIndex        =   221
            Top             =   1815
            Width           =   4245
         End
         Begin VB.ComboBox Cmb_Empresa 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0691
            Left            =   1380
            List            =   "Frm_Rpt_Reportes.frx":069E
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   300
            Width           =   4245
         End
         Begin MSComCtl2.DTPicker Dtp_Año_Comparativo_2 
            Height          =   315
            Left            =   3645
            TabIndex        =   223
            Top             =   2190
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            UpDown          =   -1  'True
            CurrentDate     =   39814
         End
         Begin VB.CheckBox Chk_Comparativo_Ciudad 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   195
            TabIndex        =   219
            Top             =   1485
            Width           =   1080
         End
         Begin VB.CheckBox Chk_Comparativo_Estado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estado"
            Height          =   195
            Left            =   195
            TabIndex        =   217
            Top             =   1110
            Width           =   1080
         End
         Begin VB.CheckBox Chk_Comparativo_Zona 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Zona"
            Height          =   195
            Left            =   195
            TabIndex        =   215
            Top             =   735
            Width           =   1080
         End
         Begin VB.ComboBox Cmb_Estado_Venta 
            Height          =   315
            Left            =   1380
            TabIndex        =   218
            Top             =   1050
            Width           =   4245
         End
         Begin VB.ComboBox Cmb_Zona_Venta 
            Height          =   315
            Left            =   1380
            TabIndex        =   216
            Top             =   675
            Width           =   4245
         End
         Begin VB.ComboBox Cmb_Ciudad_Venta 
            Height          =   315
            Left            =   1380
            TabIndex        =   220
            Top             =   1425
            Width           =   4245
         End
         Begin VB.CommandButton Btn_Comparativo_Ventas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2085
            TabIndex        =   224
            Top             =   2580
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker Dtp_Año_Comparativo_1 
            Height          =   315
            Left            =   1380
            TabIndex        =   222
            Top             =   2190
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            UpDown          =   -1  'True
            CurrentDate     =   39814
         End
         Begin VB.Label Lbl_Empresa_Ventas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   195
            TabIndex        =   226
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Lbl_Año_Comparativo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Reporte"
            Height          =   195
            Left            =   195
            TabIndex        =   225
            Top             =   1875
            Width           =   930
         End
      End
      Begin VB.Frame Fra_Reporte_Ventas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   45
         TabIndex        =   53
         Top             =   30
         Visible         =   0   'False
         Width           =   5835
         Begin VB.CheckBox Chk_Fechas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   135
            TabIndex        =   65
            Top             =   1530
            Width           =   975
         End
         Begin VB.CheckBox Chk_Producto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   135
            TabIndex        =   55
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Chk_Estatus_Materia_Prima 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   2910
            TabIndex        =   63
            Top             =   1155
            Width           =   975
         End
         Begin VB.ComboBox Cmb_Estatus_Materia_Prima 
            Height          =   315
            Left            =   4095
            TabIndex        =   64
            Top             =   1095
            Width           =   1515
         End
         Begin VB.CheckBox Chk_Numero_Acta 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Acta"
            Height          =   195
            Left            =   135
            TabIndex        =   61
            Top             =   1155
            Width           =   975
         End
         Begin VB.TextBox Txt_Numero_Acta 
            Height          =   315
            Left            =   1155
            TabIndex        =   62
            Top             =   1095
            Width           =   1650
         End
         Begin VB.ComboBox Cmb_Producto_Venta 
            Height          =   315
            Left            =   1155
            TabIndex        =   56
            Top             =   300
            Width           =   4455
         End
         Begin VB.CheckBox Chk_Numero_Lote 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Lote"
            Height          =   195
            Left            =   135
            TabIndex        =   57
            Top             =   765
            Width           =   975
         End
         Begin VB.TextBox Txt_Numero_Lote 
            Height          =   315
            Left            =   1155
            TabIndex        =   58
            Top             =   705
            Width           =   1665
         End
         Begin VB.CheckBox Chk_Fecha_Caducidad 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Caducidad"
            Height          =   195
            Left            =   2910
            TabIndex        =   59
            Top             =   765
            Width           =   1095
         End
         Begin VB.CommandButton Btn_Reporte_Historico_Ventas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2265
            TabIndex        =   68
            Top             =   2355
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Historico_Ventas 
            Height          =   315
            Left            =   1155
            TabIndex        =   66
            Top             =   1470
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
            MaxDate         =   73415
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Historico_Ventas 
            Height          =   315
            Left            =   1155
            TabIndex        =   67
            Top             =   1920
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Caducidad 
            Height          =   315
            Left            =   4095
            TabIndex        =   60
            Top             =   705
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   60751875
            CurrentDate     =   39798
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   135
            TabIndex        =   54
            Top             =   1980
            Width           =   825
         End
      End
      Begin VB.Frame Fra_Diario_Administrativo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   77
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CommandButton Btn_Reporte_Diario 
            Caption         =   "Reporte"
            Height          =   450
            Left            =   1620
            TabIndex        =   79
            Top             =   2310
            Width           =   1500
         End
         Begin VB.ComboBox Cmb_Cuenta_Banco 
            Height          =   315
            Left            =   1215
            TabIndex        =   78
            Top             =   570
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker Dtp_Inicio_Diario 
            Height          =   315
            Left            =   1200
            TabIndex        =   80
            Top             =   1132
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
            MaxDate         =   73415
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker Dtp_Fin_Diario 
            Height          =   315
            Left            =   1200
            TabIndex        =   81
            Top             =   1695
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Inicio_Diario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Incio"
            Height          =   195
            Left            =   135
            TabIndex        =   84
            Top             =   1192
            Width           =   840
         End
         Begin VB.Label Lbl_Fin_Diario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   135
            TabIndex        =   83
            Top             =   1755
            Width           =   825
         End
         Begin VB.Label Lbl_Cuenta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            Height          =   195
            Left            =   135
            TabIndex        =   82
            Top             =   630
            Width           =   510
         End
      End
      Begin VB.Frame Fra_Efectividad_Cobranza 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Efectividad de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   192
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.TextBox Txt_Objetivo_Efectividad_4 
            Height          =   315
            Left            =   1920
            TabIndex        =   197
            Top             =   1830
            Width           =   2625
         End
         Begin VB.TextBox Txt_Objetivo_Efectividad_3 
            Height          =   315
            Left            =   1920
            TabIndex        =   196
            Top             =   1464
            Width           =   2625
         End
         Begin VB.TextBox Txt_Objetivo_Efectividad_2 
            Height          =   315
            Left            =   1920
            TabIndex        =   195
            Top             =   1101
            Width           =   2625
         End
         Begin VB.TextBox Txt_Objetivo_Efectividad_1 
            Height          =   315
            Left            =   1920
            TabIndex        =   194
            Top             =   738
            Width           =   2625
         End
         Begin VB.CommandButton Btn_Efectividad_Cobranza 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            MousePointer    =   99  'Custom
            TabIndex        =   198
            Top             =   2400
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker Dtp_Mes_Evaluacion 
            Height          =   315
            Left            =   1920
            TabIndex        =   193
            Top             =   375
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39448
            MinDate         =   2
         End
         Begin VB.Label Lbl_Objetivo_4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Objetivo 4to Mes en %"
            Height          =   195
            Left            =   135
            TabIndex        =   203
            Top             =   1890
            Width           =   1590
         End
         Begin VB.Label Lbl_Objetivo_3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Objetivo 3er Mes en %"
            Height          =   195
            Left            =   135
            TabIndex        =   202
            Top             =   1530
            Width           =   1590
         End
         Begin VB.Label Lbl_Objetivo_2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Objetivo 2do Mes en %"
            Height          =   195
            Left            =   135
            TabIndex        =   201
            Top             =   1155
            Width           =   1635
         End
         Begin VB.Label Lbl_Objetivo_1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Objetivo 1er Mes en %"
            Height          =   195
            Left            =   135
            TabIndex        =   200
            Top             =   798
            Width           =   1590
         End
         Begin VB.Label Lbl_Mes_Evaluacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mes a Evaluar"
            Height          =   195
            Left            =   135
            TabIndex        =   199
            Top             =   435
            Width           =   1020
         End
      End
      Begin VB.Frame Fra_Estadistico_Cliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Estadistico Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   205
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.CommandButton Btn_Reporte_Cliente_Estadistico 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1665
            TabIndex        =   208
            Top             =   2205
            Width           =   1365
         End
         Begin VB.ComboBox Cmb_Cliente_Estadistico 
            Height          =   315
            Left            =   90
            TabIndex        =   207
            Top             =   1095
            Width           =   4620
         End
         Begin MSComCtl2.DTPicker Dtp_Año_Auditoria 
            Height          =   315
            Left            =   1920
            TabIndex        =   212
            Top             =   1590
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   60751875
            UpDown          =   -1  'True
            CurrentDate     =   39814
         End
         Begin VB.Label Lbl_Cliente_Estadistico 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   206
            Top             =   825
            Width           =   480
         End
      End
      Begin VB.Frame Fra_Reporte_Inventarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Inventarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   45
         TabIndex        =   17
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.ComboBox Cmb_Visibilidad 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":06C3
            Left            =   3090
            List            =   "Frm_Rpt_Reportes.frx":06D0
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2250
            Width           =   1635
         End
         Begin VB.CheckBox Chk_Natural_Health 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Natural Health"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   1950
            Value           =   1  'Checked
            Width           =   1470
         End
         Begin VB.CheckBox Chk_Sin_Existencia 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sin Existencia"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   2310
            Width           =   1470
         End
         Begin VB.CheckBox Chk_Pronacen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pronacen"
            Height          =   195
            Left            =   1680
            TabIndex        =   23
            Top             =   1950
            Value           =   1  'Checked
            Width           =   1470
         End
         Begin VB.CheckBox Chk_Maquilas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Maquilas"
            Height          =   195
            Left            =   3180
            TabIndex        =   24
            Top             =   1950
            Value           =   1  'Checked
            Width           =   1470
         End
         Begin VB.ComboBox Cmb_Almacen 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":06F2
            Left            =   1320
            List            =   "Frm_Rpt_Reportes.frx":06F4
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1515
            Width           =   3360
         End
         Begin VB.CheckBox Chk_Negativos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Negativos"
            Height          =   195
            Left            =   1680
            TabIndex        =   26
            Top             =   2310
            Width           =   1470
         End
         Begin VB.ComboBox Cmb_Tipo_Material 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":06F6
            Left            =   1320
            List            =   "Frm_Rpt_Reportes.frx":0703
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1140
            Visible         =   0   'False
            Width           =   3360
         End
         Begin VB.ComboBox Cmb_Producto_Materia_Prima 
            Height          =   315
            Left            =   105
            TabIndex        =   18
            Top             =   420
            Width           =   4575
         End
         Begin VB.ComboBox Cmb_Tipo_Producto 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   780
            Width           =   3360
         End
         Begin VB.CommandButton Btn_Inventario_General 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1545
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   2625
            Width           =   1695
         End
         Begin VB.Label Lbl_Tipo_Material 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Material          Almacen"
            Height          =   585
            Left            =   210
            TabIndex        =   243
            Top             =   1200
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Lbl_Producto_Materia_Prima 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   180
            TabIndex        =   237
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Lbl_Tipo_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo Producto"
            Height          =   195
            Left            =   210
            TabIndex        =   34
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Fra_Devoluciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte de Devoluciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   45
         TabIndex        =   94
         Top             =   15
         Visible         =   0   'False
         Width           =   4830
         Begin VB.CommandButton Btn_Reporte_Devolucion 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1635
            TabIndex        =   103
            Top             =   2385
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Devolucion_Fin 
            Height          =   315
            Left            =   1320
            TabIndex        =   102
            Top             =   1650
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39515
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Devolucion_Inicio 
            Height          =   315
            Left            =   1320
            TabIndex        =   100
            Top             =   1260
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39515
         End
         Begin VB.ComboBox Cmb_Cliente_Devoluciones 
            Height          =   315
            Left            =   105
            TabIndex        =   98
            Top             =   555
            Width           =   4575
         End
         Begin VB.TextBox Txt_No_Devolucion 
            Height          =   315
            Left            =   1320
            TabIndex        =   96
            Top             =   900
            Width           =   3345
         End
         Begin VB.Label Lbl_Fecha_Devolucion_Fin 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   90
            TabIndex        =   101
            Top             =   1725
            Width           =   705
         End
         Begin VB.Label Lbl_Fecha_Devolucion_Inicio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   90
            TabIndex        =   99
            Top             =   1335
            Width           =   870
         End
         Begin VB.Label Lbl_Cliente_Devolucion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   195
            Left            =   90
            TabIndex        =   97
            Top             =   345
            Width           =   480
         End
         Begin VB.Label Lbl_No_Devolucion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Devolución"
            Height          =   195
            Left            =   90
            TabIndex        =   95
            Top             =   975
            Width           =   1110
         End
      End
      Begin VB.Frame Fra_Facturas_Directas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas Directas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   45
         TabIndex        =   152
         Top             =   15
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox Cmb_Factura_Remision 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0729
            Left            =   1170
            List            =   "Frm_Rpt_Reportes.frx":0733
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   495
            Width           =   3510
         End
         Begin VB.CheckBox Chk_General 
            BackColor       =   &H00FFFFFF&
            Caption         =   "General"
            Height          =   195
            Left            =   1215
            TabIndex        =   158
            Top             =   1830
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.CommandButton Btn_Reporte_Facturas_Directas 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1560
            TabIndex        =   159
            Top             =   2220
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Facturas_Directas 
            Height          =   315
            Left            =   1170
            TabIndex        =   157
            Top             =   1365
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd,    dd   MMMM   yyyy"
            Format          =   60751875
            CurrentDate     =   39542
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Facturas_Directas 
            Height          =   315
            Left            =   1170
            TabIndex        =   156
            Top             =   930
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd,    dd   MMMM   yyyy"
            Format          =   60751875
            CurrentDate     =   39542
         End
         Begin VB.Label Lbl_Factura_Remision 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Documento"
            Height          =   195
            Left            =   120
            TabIndex        =   232
            Top             =   555
            Width           =   825
         End
         Begin VB.Label Lbl_Fecha_Fin_Facturas_Directas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   120
            TabIndex        =   154
            Top             =   1425
            Width           =   705
         End
         Begin VB.Label Lbl_Fecha_Inicio_Factura_Directa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   153
            Top             =   990
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Reporte_Pedidos_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedidos por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   45
         TabIndex        =   70
         Top             =   30
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox Txt_Reporte_No_Pedido 
            Height          =   315
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   72
            Top             =   1035
            Width           =   3285
         End
         Begin VB.ComboBox Cmb_Reporte_Pedido_Producto 
            Height          =   315
            Left            =   195
            TabIndex        =   71
            Top             =   600
            Width           =   4455
         End
         Begin VB.CommandButton Btn_Reporte_Pedidos_Cliente 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   2265
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Pedidos_Productos 
            Height          =   315
            Left            =   1350
            TabIndex        =   74
            Top             =   1785
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Pedido_Producto 
            Height          =   315
            Left            =   1350
            TabIndex        =   73
            Top             =   1395
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60751872
            CurrentDate     =   39513
         End
         Begin VB.Label Lbl_Reporte_Pedido_No_Pedido 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Pedido"
            Height          =   195
            Left            =   210
            TabIndex        =   230
            Top             =   1095
            Width           =   750
         End
         Begin VB.Label Lbl_Reporte_Pedido_Producto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Producto"
            Height          =   195
            Left            =   195
            TabIndex        =   229
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Lbl_Fecha_Inicio_Pedidos_Clientes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   195
            TabIndex        =   228
            Top             =   1455
            Width           =   870
         End
         Begin VB.Label Lbl_Fecha_Fin_Pedidos_Clientes 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Fin"
            Height          =   195
            Left            =   195
            TabIndex        =   227
            Top             =   1845
            Width           =   705
         End
      End
      Begin VB.Frame Fra_Anticipos_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anticipos Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   180
         Top             =   30
         Visible         =   0   'False
         Width           =   4860
         Begin VB.ComboBox Cmb_Estatus_Anticipo 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":074A
            Left            =   885
            List            =   "Frm_Rpt_Reportes.frx":0757
            Style           =   2  'Dropdown List
            TabIndex        =   183
            Top             =   1395
            Visible         =   0   'False
            Width           =   3900
         End
         Begin VB.ComboBox Cmb_Empresa_Anticipo 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":0777
            Left            =   885
            List            =   "Frm_Rpt_Reportes.frx":0784
            Style           =   2  'Dropdown List
            TabIndex        =   182
            Top             =   960
            Visible         =   0   'False
            Width           =   3900
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Fin_Anticipo_Cliente 
            Height          =   315
            Left            =   2880
            TabIndex        =   185
            Top             =   1830
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39554
         End
         Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio_Anticipo_Cliente 
            Height          =   315
            Left            =   885
            TabIndex        =   184
            Top             =   1830
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751875
            CurrentDate     =   39554
         End
         Begin VB.ComboBox Cmb_Cliente_Anticipo 
            Height          =   315
            Left            =   75
            TabIndex        =   181
            Top             =   555
            Width           =   4710
         End
         Begin VB.CommandButton Btn_Reporte_Anticipos_Clientes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1755
            TabIndex        =   186
            Top             =   2265
            Width           =   1365
         End
         Begin VB.Label Lbl_Estatus_Anticipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   90
            TabIndex        =   242
            Top             =   1455
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Lbl_Empresa_Anticipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   241
            Top             =   1020
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Lbl_Fecha_Inicio_Anticipo_Cilente 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   90
            TabIndex        =   188
            Top             =   1890
            Width           =   525
         End
         Begin VB.Label Lbl_Cliente_Anticipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cliente"
            Height          =   255
            Left            =   90
            TabIndex        =   187
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.Frame Fra_Facturas_Proveedor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturas por Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   45
         TabIndex        =   35
         Top             =   15
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox Txt_No_Factura 
            Height          =   315
            Left            =   3375
            TabIndex        =   40
            Top             =   390
            Visible         =   0   'False
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Inicio_Facturas_Proveedores 
            Height          =   315
            Left            =   1080
            TabIndex        =   38
            Top             =   1245
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin VB.ComboBox Cmb_Facturas_Proveedores 
            Height          =   315
            ItemData        =   "Frm_Rpt_Reportes.frx":07A9
            Left            =   75
            List            =   "Frm_Rpt_Reportes.frx":07AB
            TabIndex        =   37
            Top             =   750
            Width           =   4590
         End
         Begin VB.CommandButton Btn_Reporte_Productos_Facturados_Cliente 
            Caption         =   "Reporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1830
            TabIndex        =   41
            Top             =   2340
            Width           =   1230
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Final_Facturas_Proveedores 
            Height          =   315
            Left            =   1080
            TabIndex        =   39
            Top             =   1665
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   60751872
            CurrentDate     =   39448
         End
         Begin VB.Label Lbl_Fecha_Inicio_Proveedor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio                       Fecha Fin"
            Height          =   600
            Left            =   105
            TabIndex        =   76
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label Lbl_Proveedor_Reporte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   105
            TabIndex        =   36
            Top             =   450
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   555
      Left            =   1185
      Picture         =   "Frm_Rpt_Reportes.frx":07AD
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7050
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   10590
      Picture         =   "Frm_Rpt_Reportes.frx":0CDF
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7050
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Guardar_Disco 
      Caption         =   "Enviar a Disco"
      Height          =   555
      Left            =   7455
      Picture         =   "Frm_Rpt_Reportes.frx":1211
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7050
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Exportar_Excel 
      Caption         =   "Exportar a Excel"
      Height          =   555
      Left            =   4320
      Picture         =   "Frm_Rpt_Reportes.frx":1653
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7050
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Reporte 
      Height          =   270
      Left            =   3630
      TabIndex        =   33
      Top             =   5475
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   476
      _Version        =   393216
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      Appearance      =   0
   End
End
Attribute VB_Name = "Frm_Rpt_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngFormWidth As Long        'Ancho de la forma
Private lngFormHeight As Long       'Largo de la forma
Dim Ancho_Frame As Double
Dim Largo_Frame As Double

'*******************************************************************************
'NOMBRE_FUNCION: Reporte_Material_Transito
'DESCRIPCION   : Muestra el reporte de todo el material que está en tránsito
'PARAMETROS    :
'CREO          : Sergio Ulises Durán Hernández
'FECHA_CREO    : 15-Febrero-2011
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Reporte_Material_Transito()
Dim Rs_Detalles_Traspasos As rdoResultset 'Consulta todos los productos que coincidan de acuerdo a los parámetros proporcionados por el usuario

    Call Encabezado_Reportes("REPORTE DE MATERIAL EN TRANSITO", Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "No.Traspaso   Fecha       Destino     Codigo               Descripcion                         Cantidad    No. Lote     "
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem "No. Traspaso" & Chr(9) & "Fecha" & Chr(9) & "Destino" & Chr(9) & "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "No. Lote" & Chr(9) & "Caducidad"
    Grid_Reporte.AddItem ""
    Mi_SQL = "SELECT Ope_Traspasos.No_Traspaso,Ope_Traspasos.Fecha_Traspaso,Ope_Traspasos.Almacen_Destino_ID,Cat_Almacenes.Nombre AS Almacen_Destino"
    Mi_SQL = Mi_SQL & " ,Ope_Traspasos_Detalles.*,Cat_Materias_Primas.Codigo,Cat_Materias_Primas.Nombre AS Descripcion"
    Mi_SQL = Mi_SQL & " FROM Ope_Traspasos,Cat_Almacenes,Ope_Traspasos_Detalles,Cat_Materias_Primas"
    Mi_SQL = Mi_SQL & " WHERE Ope_Traspasos.Almacen_Destino_ID=Cat_Almacenes.Almacen_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Traspasos.No_Traspaso=Ope_Traspasos_Detalles.No_Traspaso"
    Mi_SQL = Mi_SQL & " AND Ope_Traspasos_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
    Mi_SQL = Mi_SQL & " AND (Ope_Traspasos.Estatus='PENDIENTE' OR Ope_Traspasos.Estatus='RECHAZADA')"
    Set Rs_Detalles_Traspasos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Detalles_Traspasos.EOF
        With Rs_Detalles_Traspasos
            Print #1, .rdoColumns("No_Traspaso"); _
                Spc(2); Format(.rdoColumns("Fecha_Traspaso"), "dd/MMM/yyyy"); _
                Spc(2); Mid(.rdoColumns("Almacen_Destino"), 1, 10); _
                Spc(12 - Len(Mid(.rdoColumns("Almacen_Destino"), 1, 10))); Mid(.rdoColumns("Codigo"), 1, 16); _
                Spc(18 - Len(Mid(.rdoColumns("Codigo"), 1, 16))); Mid(.rdoColumns("Descripcion"), 1, 37); _
                Spc(38 - Len(Mid(.rdoColumns("Descripcion"), 1, 37))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad_Surtida"), "#,##0.00"), 10); _
                Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 16)
            Grid_Reporte.AddItem .rdoColumns("No_Traspaso") _
                & Chr(9) & .rdoColumns("Fecha_Traspaso") _
                & Chr(9) & .rdoColumns("Almacen_Destino") _
                & Chr(9) & .rdoColumns("Codigo") _
                & Chr(9) & .rdoColumns("Descripcion") _
                & Chr(9) & .rdoColumns("Cantidad_Surtida") _
                & Chr(9) & .rdoColumns("Numero_Lote") _
                & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy")
            .MoveNext
        End With
    Wend
    Rs_Detalles_Traspasos.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Inventario_Producto_Detallado
'DESCRIPCION   : Muestra el inventario detallando las existencias por producto
'PARAMETROS    :
'CREO          : Sergio Ulises Durán Hernández
'FECHA_CREO    : 31-Enero-2011
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Inventario_Producto_Detallado()
Dim Mi_SQL As String
Dim Rs_Consulta_Producto As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Existencia As Long                              'Almacena el total de la mercancia disponible
Dim Total As Long                                   'Almacena el total de la existencia
Dim Costo_Almacen As Double                         'Almacena el total del costo del almacen
Dim Ubicacion_Producto As String
Dim Fecha_Libero As String

On Error GoTo HANDLER
    'Consulta los productos con existencia
    Mi_SQL = "SELECT Detalles_Entradas.Producto_ID,Cat_Productos.Codigo,Cat_Productos.Nombre,Detalles_Entradas.Faltante,ISNULL(Detalles_Entradas.Numero_Lote,'') AS Numero_Lote,ISNULL(Detalles_Entradas.Fecha_Caducidad,'') AS Fecha_Caducidad,Detalles_Entradas.Costo_Compra,Detalles_Entradas.Almacen_ID,Detalles_Entradas.Rack_ID,Detalles_Entradas.Nivel,Detalles_Entradas.Fila,Detalles_Entradas.Estatus,Detalles_Entradas.Fecha_Libero"
    Mi_SQL = Mi_SQL & " FROM Detalles_Entradas,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Detalles_Entradas.Producto_ID=Cat_Productos.Producto_ID"
    'Valida si muestra con existencia
    If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Faltante>0"
    Else
        If Chk_Sin_Existencia.Value = 1 And Chk_Negativos.Value = 0 Then
            Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Faltante>=0"
        Else
            If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Faltante<>0"
            Else
                Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Faltante>0"
            End If
        End If
    End If
    If Cmb_Tipo_Producto.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Productos.Tipo_Producto_ID='" & Format(Cmb_Tipo_Producto.ItemData(Cmb_Tipo_Producto.ListIndex), "000000") & "'"
    End If
    If Cmb_Producto_Materia_Prima.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID='" & Format(Cmb_Producto_Materia_Prima.ItemData(Cmb_Producto_Materia_Prima.ListIndex), "00000") & "'"
    End If
    'Filtro de empresas
    If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Productos.Empresa IS NOT NULL"
    Else
        If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
            Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1 OR Cat_Productos.Empresa=2)"
        Else
            If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1 OR Cat_Productos.Empresa=3)"
            Else
                If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 0 Then
                    Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1)"
                Else
                    If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
                        Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=2 OR Cat_Productos.Empresa=3)"
                    Else
                        If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                            Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=2)"
                        Else
                            If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                                Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=3)"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    'Filtra por ubicacion
    If Cmb_Almacen.ListIndex > -1 Then
        If Cmb_Almacen.Text = "SIN UBICACION" Then
            Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Almacen_ID IS NULL"
        Else
            If Cmb_Almacen.Text <> "TODAS" Then
                Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Almacen_ID='" & Format(Cmb_Almacen.ItemData(Cmb_Almacen.ListIndex), "00000") & "'"
            End If
        End If
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Productos.Codigo,Fecha_Caducidad"
    Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Imprime los productos encontrados
    If Not Rs_Consulta_Producto.EOF Then
        Call Encabezado_Reportes("INVENTARIO DETALLADO POR PRODUCTO TERMINADO", Format(Now, "dd/MMM/yyyy HH:mm:ss AM/PM"), "", True)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, " CODIGO                DESCRIPCION              EXISTENCIA    NUMERO LOTE     CADUCIDAD  ESTATUS  LIBERACION  UBICACION "
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Existencia" & Chr(9) & "Numero Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Estatus" & Chr(9) & "Fecha Liberacion" & Chr(9) & "Ubicacion"
        Grid_Reporte.AddItem ""
        Total = 0
        Costo_Almacen = 0
        With Rs_Consulta_Producto
            While Not .EOF
                'Consulta la ubicacion del producto
                If Not IsNull(.rdoColumns("Almacen_ID")) Then
                    'Consulta el nombre del almacen
                    Mi_SQL = "SELECT Almacen_ID,Nombre,Prefijo FROM Cat_Almacenes WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                    Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Ubicacion.EOF Then
                        Ubicacion_Producto = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                    Else
                        Ubicacion_Producto = "N/A"
                    End If
                    Rs_Consulta_Ubicacion.Close
                Else
                    Ubicacion_Producto = "N/A"
                End If
                'Fecha liberación
                If Not IsNull(.rdoColumns("Fecha_Libero")) Then
                    Fecha_Libero = Format(.rdoColumns("Fecha_Libero"), "dd/MMM/yy")
                Else
                    Fecha_Libero = "         "
                End If
'                Print #1, Spc(1); Mid(.rdoColumns("Codigo"), 1, 15); _
'                    Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(Trim(.rdoColumns("Nombre")), 1, 36); _
'                    Spc(38 - Len(Mid(Trim(.rdoColumns("Nombre")), 1, 36))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Faltante"), "#0"), 10); _
'                    Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 16); _
'                    Spc(18 - (Len(Mid(.rdoColumns("Numero_Lote"), 1, 16)))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
'                    Spc(2); Mid(Ubicacion_Producto, 1, 10); _
'                    Spc(11 - Len(Mid(Ubicacion_Producto, 1, 10))); .rdoColumns("Estatus")
                Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                    Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(Trim(.rdoColumns("Nombre")), 1, 28); _
                    Spc(30 - Len(Mid(Trim(.rdoColumns("Nombre")), 1, 28))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Faltante"), "#0"), 10); _
                    Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 15); _
                    Spc(17 - (Len(Mid(.rdoColumns("Numero_Lote"), 1, 15)))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yy"); _
                    Spc(2); Mid(.rdoColumns("Estatus"), 1, 8); _
                    Spc(10 - Len(Mid(.rdoColumns("Estatus"), 1, 8))); Fecha_Libero; _
                    Spc(2); Ubicacion_Producto
                Grid_Reporte.AddItem .rdoColumns("Codigo") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Faltante") _
                    & Chr(9) & .rdoColumns("Numero_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Estatus") _
                    & Chr(9) & Fecha_Libero _
                    & Chr(9) & Ubicacion_Producto
                Total = Val(Total) + Val(.rdoColumns("Faltante"))
                Costo_Almacen = Val(Costo_Almacen) + (Val(.rdoColumns("Faltante")) * Val(.rdoColumns("Costo_Compra")))
                Rs_Consulta_Producto.MoveNext
            Wend
        End With
        Print #1,
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime los totales de piezas y costo
        Print #1, "                                                                            TOTAL EXISTENCIA PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,###,##0"), 14)
        Grid_Reporte.AddItem Chr(9) & "TOTAL EXISTENCIA PIEZAS" & Chr(9) & Total
        Print #1, "                                                                            COSTO INVENTARIO PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & "COSTO INVENTARIO PIEZAS" & Chr(9) & Costo_Almacen
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "                                                                            COSTO TOTAL INVENTARIO        "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem Chr(9) & "COSTO TOTAL INVENTARIO" & Chr(9) & Costo_Almacen
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No tiene existencias en almacén", vbExclamation
    End If
    Rs_Consulta_Producto.Close
    Exit Sub
HANDLER:
    Close #1
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Reporte_Prestamos
    'DESCRIPCION: Muestra el detalle de los préstamos
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 04-Julio-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACION:
    'PARAMETROS:
'*******************************************************************************
Public Sub Reporte_Prestamos()
Dim Mi_SQL As String
Dim Rs_Consulta_Prestamos As rdoResultset
Dim Rs_Consulta_Detalles As rdoResultset
Dim No_Movimiento As String
Dim Total As Double
Dim Total_Abono As Double
Dim Tipo As String

On Error GoTo HANDLER
    'Consulta los productos con existencia
    Mi_SQL = "SELECT * FROM Adm_Movimientos_Futuros"
    Mi_SQL = Mi_SQL & " WHERE Tipo_Movimiento='PRESTAMO'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Prestamos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Imprime los productos encontrados
    If Not Rs_Consulta_Prestamos.EOF Then
        Call Encabezado_Reportes("REPORTE DE PRESTAMOS", Format(Now, "dd/MMM/yyyy HH:mm:ss AM/PM"), "", True)
        Print #1, "   Fecha        Descripcion                                  Cantidad         Abono           Saldo          Tipo       "
        Print #1, "________________________________________________________________________________________________________________________"
        Print #1,
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "Tipo"
        Grid_Reporte.AddItem ""
        With Rs_Consulta_Prestamos
            While Not .EOF
                If .rdoColumns("Tipo") = "E" Then
                    Tipo = "EGRESO"
                Else
                    Tipo = "INGRESO"
                End If
                Print #1, Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(2); Mid(.rdoColumns("Descripcion"), 1, 40); _
                    Spc(44 - Len(Mid(Trim(.rdoColumns("Descripcion")), 1, 40))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); _
                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad_Real"), "#,###,##0.00"), 13); _
                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad") - .rdoColumns("Cantidad_Real"), "#,###,##0.00"), 13); _
                    Spc(4); Tipo
                Grid_Reporte.AddItem .rdoColumns("Fecha") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Cantidad_Real") & Chr(9) & .rdoColumns("Cantidad") - .rdoColumns("Cantidad_Real") & Chr(9) & Tipo
                Total = Val(Total) + Val(.rdoColumns("Cantidad"))
                Total_Abono = Val(Total_Abono) + Val(.rdoColumns("Cantidad_Real"))
                Rs_Consulta_Prestamos.MoveNext
            Wend
        End With
        Print #1,
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime los totales de piezas y costo
        Print #1, "  T O T A L E S "; Spc(39); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,###,##0.00"), 16); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Abono, "#,###,###,##0.00"), 16); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total - Total_Abono, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem "" & Chr(9) & "TOTALES" & Chr(9) & Total & Chr(9) & Total_Abono & Chr(9) & Total - Total_Abono
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No tiene préstamos capturados", vbExclamation
        Unload Me
    End If
    Rs_Consulta_Prestamos.Close
    Exit Sub
HANDLER:
    Close #1
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Inventario_Producto_Detallado_MP
    'DESCRIPCIÓN: Muestra el inventario detallando las existencias por producto
    '             de acuerdo a los lotes y caducidades de la materia prima
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 22-Octubre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Public Sub Inventario_Producto_Detallado_MP()
Dim Mi_SQL As String
Dim Rs_Consulta_Producto As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Unidad As rdoResultset
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Existencia As Long                              'Almacena el total de la mercancia disponible
Dim Total As Long                                   'Almacena el total de la existencia
Dim Costo_Almacen As Double                         'Almacena el total del costo del almacen
Dim Costo_Materia_Prima As Double
Dim Nombre_Corto As String
Dim Fecha_Libero As String
Dim Ubicacion_Almacen As String

On Error GoTo HANDLER
    'Consulta los productos con existencia
    Mi_SQL = "SELECT Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID,Cat_Materias_Primas.Codigo,Cat_Materias_Primas.Nombre,Alm_Entradas_Materia_Prima_Detalles.Faltante,ISNULL(Alm_Entradas_Materia_Prima_Detalles.Numero_Lote,'') AS Numero_Lote,ISNULL(Alm_Entradas_Materia_Prima_Detalles.Fecha_Caducidad,'') AS Fecha_Caducidad,Alm_Entradas_Materia_Prima_Detalles.Costo_Compra,Cat_Materias_Primas.Unidad_ID,Alm_Entradas_Materia_Prima_Detalles.Estatus,Alm_Entradas_Materia_Prima_Detalles.Fecha_Libero,Alm_Entradas_Materia_Prima_Detalles.Costo_Conversion,Alm_Entradas_Materia_Prima_Detalles.Almacen_ID"
    Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima_Detalles,Cat_Materias_Primas"
    Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
    'Valida si muestra con existencia
    If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Faltante>0"
    Else
        If Chk_Sin_Existencia.Value = 1 And Chk_Negativos.Value = 0 Then
            Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Faltante>=0"
        Else
            If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Faltante<>0"
            Else
                Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Faltante>0"
            End If
        End If
    End If
    If Cmb_Tipo_Producto.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Tipo_Producto_ID='" & Format(Cmb_Tipo_Producto.ItemData(Cmb_Tipo_Producto.ListIndex), "000000") & "'"
    End If
    If Cmb_Producto_Materia_Prima.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Materia_Prima_ID='" & Format(Cmb_Producto_Materia_Prima.ItemData(Cmb_Producto_Materia_Prima.ListIndex), "00000") & "'"
    End If
    'Filtra por el tipo de material
    If Cmb_Tipo_Material.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Tipo_Material='" & Cmb_Tipo_Material.Text & "'"
    End If
    'Filtra por ubicacion
    If Cmb_Almacen.ListIndex > -1 Then
        If Cmb_Almacen.Text = "SIN UBICACION" Then
            Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Almacen_ID IS NULL"
        Else
            If Cmb_Almacen.Text <> "TODAS" Then
                Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Almacen_ID='" & Format(Cmb_Almacen.ItemData(Cmb_Almacen.ListIndex), "00000") & "'"
            End If
        End If
    End If
    'Filtro de empresas
    If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Empresa IS NOT NULL"
    Else
        If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
            Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1 OR Cat_Materias_Primas.Empresa=2)"
        Else
            If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1 OR Cat_Materias_Primas.Empresa=3)"
            Else
                If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 0 Then
                    Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1)"
                Else
                    If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
                        Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=2 OR Cat_Materias_Primas.Empresa=3)"
                    Else
                        If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                            Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=2)"
                        Else
                            If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                                Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=3)"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    'Filtro de la visibilidad
    If Cmb_Visibilidad.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Clasificacion='" & Cmb_Visibilidad.Text & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Materias_Primas.Codigo"
    Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Imprime los productos encontrados
    If Not Rs_Consulta_Producto.EOF Then
        Call Encabezado_Reportes("INVENTARIO DETALLADO POR MATERIA PRIMA", Format(Now, "dd/MMM/yyyy HH:mm:ss AM/PM"), "", True)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, " CODIGO                DESCRIPCION              EXISTENCIA    NUMERO LOTE     CADUCIDAD  ESTATUS  LIBERACION  UBICACION "
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.Cols = 11
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Existencia" & Chr(9) & "Unidad" & Chr(9) & "Numero Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Estatus" & Chr(9) & "Fecha Liberacion" & Chr(9) & "Ultmo Costo" & Chr(9) & "Total" & Chr(9) & "Ubicacion"
        Grid_Reporte.AddItem ""
        Total = 0
        Costo_Almacen = 0
        With Rs_Consulta_Producto
            While Not .EOF
                'Consulta la unidad del producto
                If Not IsNull(.rdoColumns("Unidad_ID")) Then
                    Mi_SQL = "SELECT Unidad_ID,Nombre_Corto FROM Cat_Unidades"
                    Mi_SQL = Mi_SQL & " WHERE Unidad_ID='" & .rdoColumns("Unidad_ID") & "'"
                    Set Rs_Consulta_Unidad = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Unidad.EOF Then
                        Nombre_Corto = Rs_Consulta_Unidad.rdoColumns("Nombre_Corto")
                    Else
                        Nombre_Corto = ""
                    End If
                    Rs_Consulta_Unidad.Close
                Else
                    Nombre_Corto = ""
                End If
                'Costo
                If Not IsNull(.rdoColumns("Costo_Conversion")) Then
                    Costo_Materia_Prima = .rdoColumns("Costo_Conversion")
                Else
                    If Not IsNull(.rdoColumns("Costo_Compra")) Then
                        Costo_Materia_Prima = .rdoColumns("Costo_Compra")
                    Else
                        Costo_Materia_Prima = 0
                    End If
                End If
                'Fecha liberación
                If Not IsNull(.rdoColumns("Fecha_Libero")) Then
                    Fecha_Libero = Format(.rdoColumns("Fecha_Libero"), "dd/MMM/yy")
                Else
                    Fecha_Libero = "         "
                End If
                'Consulta la ubicacion
                If Not IsNull(.rdoColumns("Almacen_ID")) Then
                    Mi_SQL = "SELECT Almacen_ID,Nombre FROM Cat_Almacenes"
                    Mi_SQL = Mi_SQL & " WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                    Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Ubicacion.EOF Then
                        Ubicacion_Almacen = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                    Else
                        Ubicacion_Almacen = "N/A"
                    End If
                    Rs_Consulta_Ubicacion.Close
                Else
                    Ubicacion_Almacen = "N/A"
                End If
                If Not IsNull(.rdoColumns("Fecha_Caducidad")) Then
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 17); _
                        Spc(19 - Len(Mid(.rdoColumns("Codigo"), 1, 17))); Mid(Trim(.rdoColumns("Nombre")), 1, 25); _
                        Spc(27 - Len(Mid(Trim(.rdoColumns("Nombre")), 1, 25))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Faltante"), "#0.00"), 10); _
                        Spc(1); Nombre_Corto; _
                        Spc(4 - Len(Nombre_Corto)); Mid(.rdoColumns("Numero_Lote"), 1, 15); _
                        Spc(17 - (Len(Mid(.rdoColumns("Numero_Lote"), 1, 15)))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yy"); _
                        Spc(2); Mid(.rdoColumns("Estatus"), 1, 8); _
                        Spc(10 - Len(Mid(.rdoColumns("Estatus"), 1, 8))); Fecha_Libero; _
                        Spc(2); Ubicacion_Almacen
                    Grid_Reporte.AddItem .rdoColumns("Codigo") _
                        & Chr(9) & .rdoColumns("Nombre") _
                        & Chr(9) & .rdoColumns("Faltante") _
                        & Chr(9) & Nombre_Corto _
                        & Chr(9) & .rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Estatus") _
                        & Chr(9) & Fecha_Libero _
                        & Chr(9) & Costo_Materia_Prima _
                        & Chr(9) & .rdoColumns("Faltante") * Costo_Materia_Prima _
                        & Chr(9) & Ubicacion_Almacen
                End If
                Total = Val(Total) + Val(.rdoColumns("Faltante"))
                Costo_Almacen = Val(Costo_Almacen) + (Val(.rdoColumns("Faltante")) * Val(Costo_Materia_Prima))
                Rs_Consulta_Producto.MoveNext
            Wend
        End With
        Print #1,
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime los totales de piezas y costo
        Print #1, "                                                                         TOTAL EXISTENCIA PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 17)
        Grid_Reporte.AddItem Chr(9) & "TOTAL EXISTENCIA PIEZAS" & Chr(9) & Total
        Print #1, "                                                                         COSTO INVENTARIO PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "#,###,##0.00"), 17)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & "COSTO INVENTARIO PIEZAS" & Chr(9) & Costo_Almacen
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "                                                                         COSTO TOTAL INVENTARIO        "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "#,###,##0.00"), 17)
        Grid_Reporte.AddItem Chr(9) & "COSTO TOTAL INVENTARIO" & Chr(9) & Costo_Almacen
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No tiene existencias en almacén", vbExclamation
    End If
    Rs_Consulta_Producto.Close
    Exit Sub
HANDLER:
    Close #1
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Reporte_Facturas_Vencer()
Dim Mi_SQL As String
Dim Proveedor As String                                     'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Total_Facturas As Double                                 'Almacena el total de las facturas
Dim Saldo_Facturas As Double
Dim Abono_Facturas As Double

    If Cmb_Empresa_Factura_Vencida.ListIndex > -1 Then
        Grid_Reporte.Cols = 8
        Grid_Reporte.Rows = 0
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
        Mi_SQL = "SELECT Adm_Facturas_Proveedores.No_Factura,Adm_Facturas_Proveedores.Proveedor_ID,Adm_Facturas_Proveedores.Fecha,Adm_Facturas_Proveedores.Fecha_Pago,Adm_Facturas_Proveedores.Importe,Adm_Facturas_Proveedores.IVA,Total,Adm_Facturas_Proveedores.Saldo,Adm_Facturas_Proveedores.Abono,Adm_Facturas_Proveedores.Empresa,Cat_Proveedores.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores LEFT JOIN Cat_Proveedores ON Adm_Facturas_Proveedores.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Proveedores.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Pagada='N'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Fecha_Pago<='" & Format(Dtp_Facturas_Vencidas.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Empresa=" & Val(Mid(Cmb_Empresa_Factura_Vencida.Text, 1, 1))
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Proveedores.Nombre,Adm_Facturas_Proveedores.Fecha"
        Set Rs_Consulta_Cat_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Facturas_Proveedores.EOF Then
            With Rs_Consulta_Cat_Facturas_Proveedores
                Call Encabezado_Reportes("REPORTE DE FACTURAS POR VENCER", Format(Dtp_Facturas_Vencidas.Value, "dd/MMM/yyyy"), "", True)
                If Cmb_Empresa_Factura_Vencida.ListIndex = 0 Then
                    Print #1, "                                             FACTURAS DE NATURAL HEALTH"
                    Grid_Reporte.AddItem Chr(9) & Chr(9) & "FACTURAS DE NATURAL HEALTH"
                Else
                    Print #1, "                                                FACTURAS DE PRONACEN"
                    Grid_Reporte.AddItem Chr(9) & Chr(9) & "FACTURAS DE PRONACEN"
                End If
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, " No Factura               Proveedor                      Fecha     Fecha Pago      Total         Saldo          Abono    "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Fecha Factura" & Chr(9) & "Fecha Pago" & Chr(9) & "Total" & Chr(9) & "Saldo" & Chr(9) & "Abono"
                Grid_Reporte.AddItem ""
                While Not .EOF
                    If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores"
                        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                        Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Proveedores.EOF Then
                            Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                        Else
                            Proveedor = ""
                        End If
                        Rs_Consulta_Cat_Proveedores.Close
                    End If
                    Print #1, Mid(Trim(.rdoColumns("No_Factura")), 1, 15); Spc(17 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 15))); Mid(Proveedor, 1, 37); _
                        Spc(39 - Len(Mid(Proveedor, 1, 37))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem Trim(.rdoColumns("No_Factura")) & Chr(9) & Proveedor & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00")
                    Total_Facturas = Total_Facturas + Val(.rdoColumns("Total"))
                    Saldo_Facturas = Saldo_Facturas + Val(.rdoColumns("Saldo"))
                    Abono_Facturas = Abono_Facturas + Val(.rdoColumns("Abono"))
                    .MoveNext
                Wend
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(60); "TOTALES : "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Facturas, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTALES :" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Saldo_Facturas, "#,###,##0.00") & Chr(9) & Format(Abono_Facturas, "#,###,##0.00")
                Call Encabezado_Reportes("", "", "", False)
            End With
        Else
            MsgBox "No hay facturas Vencidas", vbInformation
        End If
        Rs_Consulta_Cat_Facturas_Proveedores.Close
    Else
        MsgBox "Seleccione una Empresa", vbInformation
    End If
End Sub

Private Sub Reporte_Facturas_Vencidas()
Dim Mi_SQL As String
Dim Proveedor As String                                     'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Total_Facturas As Double                                'Almacena el total de las facturas
Dim Saldo_Facturas As Double
Dim Abono_Facturas As Double

    If Cmb_Empresa_Factura_Vencida.ListIndex > -1 Then
        Grid_Reporte.Cols = 8
        Grid_Reporte.Rows = 0
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
        Mi_SQL = "SELECT Adm_Facturas_Proveedores.No_Factura,Adm_Facturas_Proveedores.Proveedor_ID,Adm_Facturas_Proveedores.Fecha,Adm_Facturas_Proveedores.Fecha_Pago,Adm_Facturas_Proveedores.Importe,Adm_Facturas_Proveedores.IVA,Total,Adm_Facturas_Proveedores.Saldo,Adm_Facturas_Proveedores.Abono,Adm_Facturas_Proveedores.Empresa,Cat_Proveedores.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores LEFT JOIN Cat_Proveedores ON Adm_Facturas_Proveedores.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Proveedores.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Pagada='N'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Fecha_Pago<'" & Format(Now, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Empresa=" & Val(Mid(Cmb_Empresa_Factura_Vencida.Text, 1, 1))
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Proveedores.Nombre,Adm_Facturas_Proveedores.Fecha"
        Set Rs_Consulta_Cat_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Facturas_Proveedores.EOF Then
            With Rs_Consulta_Cat_Facturas_Proveedores
                Call Encabezado_Reportes("REPORTE DE FACTURAS VENCIDAS", Format(Dtp_Facturas_Vencidas.Value, "dd/MMM/yyyy"), "", True)
                If Cmb_Empresa_Factura_Vencida.ListIndex = 0 Then
                    Print #1, "                                             FACTURAS DE NATURAL HEALTH"
                    Grid_Reporte.AddItem Chr(9) & Chr(9) & "FACTURAS DE NATURAL HEALTH"
                Else
                    Print #1, "                                                FACTURAS DE PRONACEN"
                    Grid_Reporte.AddItem Chr(9) & Chr(9) & "FACTURAS DE PRONACEN"
                End If
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, " No Factura               Proveedor                      Fecha     Fecha Pago      Total         Saldo          Abono    "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Fecha Factura" & Chr(9) & "Fecha Pago" & Chr(9) & "Total" & Chr(9) & "Saldo" & Chr(9) & "Abono"
                Grid_Reporte.AddItem ""
                While Not .EOF
                    If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores"
                        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                        Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Proveedores.EOF Then
                            Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                        Else
                            Proveedor = ""
                        End If
                        Rs_Consulta_Cat_Proveedores.Close
                    End If
                    Print #1, Mid(Trim(.rdoColumns("No_Factura")), 1, 15); _
                        Spc(17 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 15))); Mid(Proveedor, 1, 37); _
                        Spc(39 - Len(Mid(Proveedor, 1, 37))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Proveedor & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Fecha_Pago"), "dd/MMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00")
                    Total_Facturas = Total_Facturas + Val(.rdoColumns("Total"))
                    Saldo_Facturas = Saldo_Facturas + Val(.rdoColumns("Saldo"))
                    Abono_Facturas = Abono_Facturas + Val(.rdoColumns("Abono"))
                    .MoveNext
                Wend
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(60); "TOTALES : "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Facturas, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Facturas, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTALES :" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Saldo_Facturas, "#,###,##0.00") & Chr(9) & Format(Abono_Facturas, "#,###,##0.00")
                Call Encabezado_Reportes("", "", "", False)
            End With
        Else
            MsgBox "No hay facturas Vencidas", vbInformation
        End If
        Rs_Consulta_Cat_Facturas_Proveedores.Close
    Else
        MsgBox "Seleccione una Empresa", vbInformation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Anticipos_Clientes
    'DESCRIPCIÓN: Muestra el reporte de los anticipos hechos
    'CREO: Rafael Muñoz
    'FECHA_CREO: 16/Abril/2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Anticipos_Clientes()
Dim Rs_Consulta_Anticipos As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Bancos As rdoResultset          'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Ciudad As String                            'Almacena el nombre de la ciudad
Dim Estatus As String                           'Almacena el estatus del anticipo
Dim Banco As String                             'Almacena el nombre del banco
Dim Cliente As String
Dim Empresa As String                           'Almacena el nombre de la empresa

    'Consulta los anticipos
    Mi_SQL = "SELECT DISTINCT Adm_Anticipos_Clientes.No_Anticipo,Adm_Anticipos_Clientes.Cliente_ID,Adm_Anticipos_Clientes.Total,Adm_Anticipos_Clientes.Fecha,Adm_Anticipos_Clientes.Forma_Pago,Adm_Anticipos_Clientes.Concepto,Adm_Anticipos_Clientes.Banco_ID,Adm_Anticipos_Clientes.Aplicado,Adm_Anticipos_Clientes.No_Factura,Adm_Anticipos_Clientes.Empresa,Cat_Clientes.Nombre,Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " FROM Adm_Anticipos_Clientes,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Fecha >=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Cliente_Anticipo.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Clientes.Cliente_ID= '" & Format(Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Anticipos_Clientes.Cliente_ID"
    Set Rs_Consulta_Anticipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Anticipos.EOF Then
        Call Encabezado_Reportes("REPORTE DE ANTICIPOS DE CLIENTES", Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "dd/MMM/yyyy"), True)
        Print #1,
        Grid_Reporte.AddItem ""
        With Rs_Consulta_Anticipos
            While Not .EOF
            'Imprime los anticipos
            If Cliente <> .rdoColumns("Cliente_ID") Then
                Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad"
                Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Ciudades"
                Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Clientes
                    If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                        Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                        Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                        Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                        Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                        Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ciudades.EOF Then
                            Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                        Else
                            Ciudad = ""
                        End If
                        Rs_Consulta_Ciudades.Close
                    Else
                        Ciudad = ""
                    End If
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1,
                    Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
                    Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                    Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                    Print #1, Spc(2); "Ciudad:    "; Ciudad
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                    Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                    Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
                    Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
                    Print #1,
                    If Rs_Consulta_Anticipos.rdoColumns("Aplicado") = "N" Then
                        Print #1, " Fecha        Anticipo         Banco            Forma Pago         Concepto            Estatus"
                        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Anticipo" & Chr(9) & "Banco" & Chr(9) & "Forma Pago" & Chr(9) & "Concepto" & Chr(9) & "Estatus"
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1,
                    Else
                        Print #1, " Fecha        Anticipo         Banco            Forma Pago         Concepto            Estatus        No. Factura"
                        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Anticipo" & Chr(9) & "Banco" & Chr(9) & "Forma Pago" & Chr(9) & "Concepto" & Chr(9) & "Estatus" & Chr(9) & "No. Factura" & Chr(9) & "Empresa"
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1,
                    End If
                End With
                Grid_Reporte.AddItem ""
                Cliente = .rdoColumns("Cliente_ID")
            End If
                If Not IsNull(.rdoColumns("Aplicado")) Then
                    If .rdoColumns("Aplicado") = "S" Then
                        Estatus = "Aplicado"
                    Else
                        Estatus = "No Aplicado"
                    End If
                End If
                If Not IsNull(.rdoColumns("Empresa")) Then
                    If .rdoColumns("Empresa") = "1" Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                End If
                If Not IsNull(.rdoColumns("Banco_ID")) Then
                    Mi_SQL = "SELECT Nombre,Banco_ID FROM Cat_Bancos"
                    Mi_SQL = Mi_SQL & " WHERE Banco_ID = '" & .rdoColumns("Banco_ID") & "'"
                    Set Rs_Consulta_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Bancos.EOF Then
                        Banco = Rs_Consulta_Bancos.rdoColumns("Nombre")
                    Else
                        Banco = ""
                    End If
                Else
                    Banco = ""
                End If
                If .rdoColumns("Aplicado") = "N" Then
                    Print #1, Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); Spc(3); Mid(Banco, 1, 18); Spc(23 - Len(Mid(Banco, 1, 18))); Mid(.rdoColumns("Forma_Pago"), 1, 14); Spc(15 - Len(Mid(.rdoColumns("Forma_Pago"), 1, 14))); Mid(.rdoColumns("Concepto"), 1, 20); Spc(22 - Len(Mid(.rdoColumns("Concepto"), 1, 20))); Estatus '; Spc(5); .rdoColumns("No_Factura"); Spc(1); Empresa
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Banco & Chr(9) & .rdoColumns("Forma_Pago") & Chr(9) & .rdoColumns("Concepto") & Chr(9) & Estatus
                Else
                    Print #1, Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); Spc(3); Mid(Banco, 1, 18); Spc(23 - Len(Mid(Banco, 1, 18))); Mid(.rdoColumns("Forma_Pago"), 1, 14); Spc(15 - Len(Mid(.rdoColumns("Forma_Pago"), 1, 14))); Mid(.rdoColumns("Concepto"), 1, 20); Spc(22 - Len(Mid(.rdoColumns("Concepto"), 1, 20))); Estatus; Spc(5); .rdoColumns("No_Factura"); Spc(3); Empresa
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Banco & Chr(9) & .rdoColumns("Forma_Pago") & Chr(9) & .rdoColumns("Concepto") & Chr(9) & Estatus & Chr(9) & .rdoColumns("No_Factura") & Chr(9) & Empresa
                End If
                .MoveNext
            Wend
        End With
        Rs_Consulta_Anticipos.Close
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay anticipos en las fechas seleccionadas", vbInformation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Pago_Saldos_Minimos
    'DESCRIPCIÓN: Muestra el reporte de los pagos por ajuste de facturas con
    '             saldos míminos
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 16-Octubre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Pago_Saldos_Minimos()
Dim Mi_SQL As String
Dim Rs_Consulta_Cuenta_Ajustes As rdoResultset
Dim Rs_Consulta_Cuenta_Ajustes_Detalles As rdoResultset
Dim Total_Consulta As Double
Dim Empresa As String

    'Consulta los datos de la cuenta de ajustes
    Mi_SQL = "SELECT * FROM Adm_Cuentas_Ajustes"
    Mi_SQL = Mi_SQL & " WHERE No_Cuenta='" & Format(Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.ListIndex), "00000") & "'"
    Set Rs_Consulta_Cuenta_Ajustes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cuenta_Ajustes.EOF Then
        Grid_Reporte.AddItem ""
        Call Encabezado_Reportes("REPORTE DE CUENTA DE AJUSTES", Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "dd/MMM/yyyy"), True)
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(2); "Cuenta:    "; Cmb_Cliente_Anticipo.Text
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Cuenta: " & Chr(9) & Cmb_Cliente_Anticipo.Text
        'Valida el estatus de la cuenta
        If Rs_Consulta_Cuenta_Ajustes.rdoColumns("Cancelada") = "S" Then
            Print #1, Spc(2); "Estatus:   "; "CANCELADA"
            Grid_Reporte.AddItem "Estatus: " & Chr(9) & "CANCELADA"
        Else
            If Rs_Consulta_Cuenta_Ajustes.rdoColumns("Pagada") = "S" Then
                Print #1, Spc(2); "Estatus:   "; "PAGADA"
                Grid_Reporte.AddItem "Estatus: " & Chr(9) & "PAGADA"
            Else
                Print #1, Spc(2); "Estatus:   "; "ACTIVA"
                Grid_Reporte.AddItem "Estatus: " & Chr(9) & "ACTIVA"
            End If
        End If
        Print #1, Spc(2); "Total:     "; Format(Rs_Consulta_Cuenta_Ajustes.rdoColumns("Total"), "#,###,##0.00")
        Grid_Reporte.AddItem "Total: " & Chr(9) & Format(Rs_Consulta_Cuenta_Ajustes.rdoColumns("Total"), "#,###,##0.00")
        Print #1, Spc(2); "Comentarios: "; Rs_Consulta_Cuenta_Ajustes.rdoColumns("Comentarios")
        Grid_Reporte.AddItem "Comentatios: " & Chr(9) & Rs_Consulta_Cuenta_Ajustes.rdoColumns("Comentarios")
        Grid_Reporte.AddItem ""
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "   No. Factura      Cliente                                      Cantidad        Fecha       Referencia           Tipo  "
        Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Cantidad" & Chr(9) & "Fecha" & Chr(9) & "Referencia" & Chr(9) & "Tipo"
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1,
        'Consulta los detalles de la cuenta de ajustes
        Mi_SQL = "SELECT Adm_Cuentas_Ajustes_Detalles.*,Cat_Clientes.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Cuentas_Ajustes_Detalles,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Cuentas_Ajustes_Detalles.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Cuentas_Ajustes_Detalles.No_Cuenta='" & Rs_Consulta_Cuenta_Ajustes.rdoColumns("No_Cuenta") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Cuentas_Ajustes_Detalles.Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Adm_Cuentas_Ajustes_Detalles.Fecha<= " & Par_Fecha & Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Cuentas_Ajustes_Detalles.Fecha"
        Set Rs_Consulta_Cuenta_Ajustes_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Cuenta_Ajustes_Detalles
            While Not .EOF
                'Valida la empresa
                If .rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                Print #1, Spc(2); .rdoColumns("No_Factura"); Spc(1); Empresa; _
                    Spc(2); Mid(.rdoColumns("Nombre"), 1, 45); _
                    Spc(48 - Len(Mid(.rdoColumns("Nombre"), 1, 45))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,##0.00"), 10); _
                    Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(2); Mid(.rdoColumns("Referencia"), 1, 20); _
                    Spc(25 - Len(Mid(.rdoColumns("Referencia"), 1, 20))); .rdoColumns("Tipo")
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & " - " & Empresa _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Referencia") _
                    & Chr(9) & .rdoColumns("Tipo")
                If .rdoColumns("Tipo") = "E" Then
                    Total_Consulta = Total_Consulta + .rdoColumns("Cantidad")
                Else
                    Total_Consulta = Total_Consulta - .rdoColumns("Cantidad")
                End If
                .MoveNext
            Wend
        End With
        Rs_Consulta_Cuenta_Ajustes_Detalles.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(10); "TOTAL CONSULTA: "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Consulta, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem "" & Chr(9) & "TOTAL CONSULTA: " & Chr(9) & Format(Total_Consulta, "#,###,##0.00")
    End If
    Rs_Consulta_Cuenta_Ajustes.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

Public Sub Corte_Diario()
'Dim Total_Contado As Double             'Almacena el total por concepto de ventas de contado
'Dim Total_Apartado As Double            'Almacena el total por concepto de apartados
'Dim Total_Pedido As Double              'Almacena el total por concepto de pedidos
'Dim Total_Efectivo As Double            'Almacena el total pagado en efectivo
'Dim Total_Banco As Double               'Almacena el total pagado en bancos
'Dim Total_Nota_Credito As Double        'Almacena el total pagado en nota de crédito
'Dim Total_Abono_Apartados As Double     'Almacena los pagos de apartados
'Dim Total_Abono_Pedidos As Double       'Almacena los pagos de pedidos
'Dim Total As Double                     'Almacena el gran total
'Dim Rs_Ope_Remisiones As rdoResultset   'Consulta las remisiones del día
'Dim Estatus As String
'
'    Call Encabezado_Reportes("REPORTE DE CORTE DE CAJA", Format(Dtp_Corte_Fecha_Inicial.Value, "dd/MMM/yyyy"), "", True)
'    'Ventas de contado en efectivo
'    Total_Contado = 0
'    Total_Efectivo = 0
'    Total_Banco = 0
'    Total_Nota_Credito = 0
'    Mi_SQL = "SELECT * FROM Ope_Remisiones"
'    Mi_SQL = Mi_SQL & " WHERE Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Fecha=" & Par_Fecha & Format(Dtp_Corte_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & ""
'    Mi_SQL = Mi_SQL & " ORDER BY No_Remision"
'    Set Rs_Ope_Remisiones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    If Not Rs_Ope_Remisiones.EOF Then
'        Print #1, Conectar_Ayudante.Centra("Ventas al Contado en Efectivo", Len("************************************************************************************************************************"))
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, " No Nota         Fecha        Cliente                                                                       Importe     "
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        With Rs_Ope_Remisiones
'            While Not .EOF
'                Print #1, .rdoColumns("No_Remision"); Spc(4); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); Spc(4);
'                If .rdoColumns("Cancelada") = "N" Then
'                    Print #1, "VENTAS DE MOSTRADOR"; Spc(68 - Len("VENTAS DE MOSTRADOR")); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Total"), "###,###,###.00"), 17)
'                    Total_Contado = Total_Contado + Val(.rdoColumns("Total"))
'                    Total_Efectivo = Total_Efectivo + Val(.rdoColumns("Efectivo")) - Val(.rdoColumns("Cambio"))
'                    Total_Banco = Total_Banco + Val(.rdoColumns("Banco"))
'                    Total_Nota_Credito = Total_Nota_Credito + Val(.rdoColumns("Nota_Credito"))
'                Else
'                    Print #1, "VENTAS DE MOSTRADOR *"; Spc(68 - Len("VENTAS DE MOSTRADOR *")); Spc(12); Conectar_Ayudante.Alinea_Derecha(Format(0, "###,###,##0.00"), 17)
'                End If
'                Rs_Ope_Remisiones.MoveNext
'            Wend
'        End With
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, Spc(50); "Total de ventas a contado:"; Spc(18); Conectar_Ayudante.Alinea_Derecha(Format(Total_Contado, "###,###,##0.00"), 17)
'        Print #1,
'        Print #1,
'    End If
'    Rs_Ope_Remisiones.Close
'    'Apartados de mercancia
'    Total_Apartado = 0
'    Mi_SQL = "SELECT * FROM Ope_Apartados_Clientes"
'    Mi_SQL = Mi_SQL & " WHERE Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Fecha_Apartado=" & Par_Fecha & Format(Dtp_Corte_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & ""
'    Mi_SQL = Mi_SQL & " ORDER BY No_Apartado"
'    Set Rs_Ope_Remisiones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    If Not Rs_Ope_Remisiones.EOF Then
'        Print #1, Conectar_Ayudante.Centra("Apartados", Len("************************************************************************************************************************"))
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, " No Apartado     Fecha        Cliente                                                                       Importe"
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        With Rs_Ope_Remisiones
'            While Not .EOF
'                Print #1, .rdoColumns("No_Apartado"); Spc(4); Format(.rdoColumns("Fecha_Apartado"), "dd/MMM/yyyy"); Spc(4);
'                If .rdoColumns("Cancelado") = "N" Then
'                    Print #1, Mid(.rdoColumns("Nombre"), 1, 65); Spc(68 - Len(Mid(.rdoColumns("Nombre"), 1, 65))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "###,###,##0.00"), 17)
'                    Total_Apartado = Total_Apartado + Val(.rdoColumns("Abono"))
'                Else
'                    Print #1, Mid(.rdoColumns("Nombre"), 1, 65); " *"; Spc(68 - Len(Mid(.rdoColumns("Nombre"), 1, 65))); Spc(12); Conectar_Ayudante.Alinea_Derecha(Format(0, "###,###,##0.00"), 17)
'                End If
'                Rs_Ope_Remisiones.MoveNext
'            Wend
'        End With
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, Spc(50); "Total en apartados:       "; Spc(18); Conectar_Ayudante.Alinea_Derecha(Format(Total_Apartado, "###,###,###.00"), 17)
'        Print #1,
'        Print #1,
'    End If
'    Rs_Ope_Remisiones.Close
'    'Pedidos de mercancia
'    Total_Pedido = 0
'    Mi_SQL = "SELECT * FROM Ope_Pedidos_Clientes"
'    Mi_SQL = Mi_SQL & " WHERE Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Fecha_Pedido=" & Par_Fecha & Format(Dtp_Corte_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha & ""
'    Mi_SQL = Mi_SQL & " ORDER BY No_Pedido"
'    Set Rs_Ope_Remisiones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    If Not Rs_Ope_Remisiones.EOF Then
'        Print #1, Conectar_Ayudante.Centra("Pedidos", Len("************************************************************************************************************************"))
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, " No Pedido       Fecha        Cliente                                                                       Importe"
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        With Rs_Ope_Remisiones
'            While Not .EOF
'                Print #1, .rdoColumns("No_Pedido"); Spc(4); Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(4);
'                If .rdoColumns("Cancelado") = "N" Then
'                    Print #1, Mid(.rdoColumns("Nombre"), 1, 65); Spc(68 - Len(Mid(.rdoColumns("Nombre"), 1, 65))); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Abono"), "###,###,##0.00"), 17)
'                    Total_Pedido = Total_Pedido + Val(.rdoColumns("Abono"))
'                Else
'                    Print #1, Mid(.rdoColumns("Nombre"), 1, 65); " *"; Spc(68 - Len(Mid(.rdoColumns("Nombre"), 1, 65))); Spc(12); Conectar_Ayudante.Alinea_Derecha(Format(0, "###,###,##0.00"), 17)
'                End If
'                Rs_Ope_Remisiones.MoveNext
'            Wend
'        End With
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, Spc(50); "Total en pedidos         :"; Spc(18); Conectar_Ayudante.Alinea_Derecha(Format(Total_Pedido, "###,###,###.00"), 17)
'        Print #1,
'        Print #1,
'    End If
'    Rs_Ope_Remisiones.Close
'    'Abonos de clientes para apartados
'    Mi_SQL = "SELECT Ope_Movimientos.No_Apartado,Monto,Nombre,Abono,Total"
'    Mi_SQL = Mi_SQL & " FROM Ope_Movimientos,Ope_Apartados_Clientes"
'    Mi_SQL = Mi_SQL & " WHERE Ope_Movimientos.No_Apartado=Ope_Apartados_Clientes.No_Apartado"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Ope_Apartados_Clientes.Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Fecha=" & Par_Fecha & Format(Dtp_Corte_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
'    Mi_SQL = Mi_SQL & " AND Cancelado='N'"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Estatus<>'C'"
'    Mi_SQL = Mi_SQL & " AND Descripcion<>'ABONO POR APARTADO - ENGANCHE'"
'    Set Rs_Ope_Remisiones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    If Not Rs_Ope_Remisiones.EOF Then
'        Print #1, Conectar_Ayudante.Centra("Pagos de Apartados de Clientes", Len("************************************************************************************************************************"))
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, " No Apartado   Estatus         Cliente                                                                      Importe"
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        With Rs_Ope_Remisiones
'            While Not .EOF
'                If (Val(.rdoColumns("Abono")) = Val(.rdoColumns("Total"))) Then
'                    Estatus = "Pagada     "
'                Else
'                    Estatus = "Abonada    "
'                End If
'                Print #1, .rdoColumns("No_Apartado"); Spc(4); Estatus; Spc(4); Mid(.rdoColumns("Nombre"), 1, 70); Spc(54 - Len(Mid(.rdoColumns("Nombre"), 1, 70))); Spc(14); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Monto"), "###,###,###.00"), 17)
'                Total_Abono_Apartados = Total_Abono_Apartados + Val(.rdoColumns("Monto"))
'                Rs_Ope_Remisiones.MoveNext
'            Wend
'        End With
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, Spc(50); "Total pagos de clientes:   "; Spc(20); Conectar_Ayudante.Alinea_Derecha(Format(Total_Abono_Apartados, "###,###,###.00"), 17)
'        Print #1,
'    End If
'    Rs_Ope_Remisiones.Close
'    Print #1,
'    'Abonos de clientes para pedidos
'    Mi_SQL = "SELECT Ope_Movimientos.No_Pedido,Monto,Nombre,Abono,Total"
'    Mi_SQL = Mi_SQL & " FROM Ope_Movimientos,Ope_Pedidos_Clientes"
'    Mi_SQL = Mi_SQL & " WHERE Ope_Movimientos.No_Pedido=Ope_Pedidos_Clientes.No_Pedido"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Ope_Pedidos_Clientes.Tienda_ID='" & Tienda_ID & "'"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Fecha=" & Par_Fecha & Format(Dtp_Corte_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
'    Mi_SQL = Mi_SQL & " AND Cancelado='N'"
'    Mi_SQL = Mi_SQL & " AND Ope_Movimientos.Estatus<>'C'"
'    Mi_SQL = Mi_SQL & " AND Descripcion<>'ABONO POR PEDIDO - ENGANCHE'"
'    Set Rs_Ope_Remisiones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    If Not Rs_Ope_Remisiones.EOF Then
'        Print #1, Conectar_Ayudante.Centra("Pagos de Pedidos de Clientes", Len("************************************************************************************************************************"))
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, " No Pedido     Estatus         Cliente                                                                      Importe"
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        With Rs_Ope_Remisiones
'            While Not .EOF
'                If (Val(.rdoColumns("Abono")) = Val(.rdoColumns("Total"))) Then
'                    Estatus = "Pagada     "
'                Else
'                    Estatus = "Abonada    "
'                End If
'                Print #1, .rdoColumns("No_Pedido"); Spc(4); Estatus; Spc(4); Mid(.rdoColumns("Nombre"), 1, 70); Spc(54 - Len(Mid(.rdoColumns("Nombre"), 1, 70))); Spc(14); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Monto"), "###,###,###.00"), 17)
'                Total_Abono_Pedidos = Total_Abono_Pedidos + Val(.rdoColumns("Monto"))
'                Rs_Ope_Remisiones.MoveNext
'            Wend
'        End With
'        Print #1, "------------------------------------------------------------------------------------------------------------------------"
'        Print #1, Spc(50); "Total pagos de clientes:   "; Spc(20); Conectar_Ayudante.Alinea_Derecha(Format(Total_Abono_Pedidos, "###,###,###.00"), 17)
'        Print #1,
'    End If
'    Print #1,
'    Rs_Ope_Remisiones.Close
'    Print #1,
'    Print #1, Conectar_Ayudante.Centra("Resumen de Corte Diario de Caja", Len("************************************************************************************************************************"))
'    Print #1, "------------------------------------------------------------------------------------------------------------------------"
'    Print #1,
'    Print #1, Spc(10); ""
'    Print #1, Spc(10); "______________________________________________________"
'    Print #1, Spc(10); ""
'    Print #1,
'    Print #1, Spc(10); "Efectivo                 $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Efectivo, "###,###,###.00"), 17)
'    Print #1, Spc(10); "Bancos                   $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Banco, "###,###,###.00"), 17)
'    Print #1, Spc(10); "Notas de Credito         $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Nota_Credito, "###,###,###.00"), 17)
'    Print #1, Spc(10); "                           ---------------------"
'    Print #1, Spc(10); "Ventas de Contado        $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Contado, "###,###,###.00"), 17)
'    Print #1,
'    Print #1, Spc(10); "Apartados                $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Apartado, "###,###,###.00"), 17)
'    Print #1, Spc(10); "Pedidos                  $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Pedido, "###,###,###.00"), 17)
'    Print #1,
'    Print #1, Spc(10); "Pagos de Clientes        $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Abono_Apartados + Total_Abono_Pedidos, "###,###,###.00"), 17)
'    Print #1, Spc(10); "                           ---------------------"
'    Total = Total_Efectivo + Total_Banco + Total_Nota_Credito + Total_Apartado + Total_Pedido + Total_Abono_Pedidos + Total_Abono_Apartados
'    Print #1, Spc(10); "Total Registrado         $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total, "###,###,###.00"), 17)
'    Print #1,
'    Print #1, Spc(10); "Total Deposito           $ "; Conectar_Ayudante.Alinea_Derecha(Format(Total_Efectivo + Total_Apartado + Total_Pedido + Total_Abono_Pedidos + Total_Abono_Apartados, "###,###,###.00"), 17)
'    Print #1,
'    Print #1,
'    Print #1, Spc(10); "Desgloce de de pagos, Billetes y Monedas"
'    Print #1,
'    Print #1, Spc(10); "N Credito  :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "Vouchers   :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "Billetes y Monedas"
'    Print #1,
'    Print #1, Spc(10); "1000.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); " 500.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); " 200.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); " 100.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "  50.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "  20.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "  10.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "   5.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "   2.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "   1.00    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "    .50    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "    .20    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "    .10    :__________  _________________"
'    Print #1,
'    Print #1, Spc(10); "Total del deposito  : $ _________________"
'    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Inventario
    'DESCRIPCIÓN: Muestra el inventario general
    'CREO: Sergio Ulises Druán Hernández
    'FECHA_CREO: 29-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Inventario()
Dim Rs_Consulta_Tipo_Producto As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Producto As rdoResultset            'Manejo de registro
Dim Rs_Consulta As rdoResultset                     'Manejo de registro
Dim Rs_Consulta_Detalles_Pedidos As rdoResultset    'Manejo de registro
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Existencia As Double                            'Almacena el total de la mercancia disponible
Dim Total As Long                                   'Almacena el total de la existencia
Dim Costo_Almacen As Double                         'Almacena el total del costo del almacen
Dim Ubicacion_Almacen As String

    Call Encabezado_Reportes("INVENTARIO GENERAL DE PRODUCTO TERMINADO", Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "  CODIGO             DESCRIPCION                           EXISTENCIA  DISPONIBLE    COSTO        TOTAL      UBICACION  "
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Existencia Disponible" & Chr(9) & "Disponible - Pedidos" & Chr(9) & "Ultimo Costo" & Chr(9) & "Total" & Chr(9) & "Ubicacion"
    Grid_Reporte.AddItem ""
    Print #1,
    Total = 0
    Costo_Almacen = 0
    'Consulta la lineas
    Mi_SQL = "SELECT Tipo_Producto_ID,Nombre FROM Cat_Tipos_Productos"
    'Filtra por la linea seleccionada por el usuario
    If Cmb_Tipo_Producto.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " WHERE Tipo_Producto_ID='" & Format(Cmb_Tipo_Producto.ItemData(Cmb_Tipo_Producto.ListIndex), "00000") & "'"
    Else
        Mi_SQL = Mi_SQL & " WHERE Tipo_Producto_ID<>'00016'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Tipo_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Hace el recorrido de los registros encontrados
    While Not Rs_Consulta_Tipo_Producto.EOF
        'Consulta los productos de la linea en Turno
        Mi_SQL = "SELECT Cat_Productos.Producto_ID,Cat_Productos.Nombre,Cat_Productos.Costo,Cat_Productos.Codigo,SUM(Detalles_Entradas.Faltante) AS Existencia,Detalles_Entradas.Almacen_ID"
        Mi_SQL = Mi_SQL & " FROM Cat_Productos,Detalles_Entradas"
        Mi_SQL = Mi_SQL & " WHERE Cat_Productos.Producto_ID=Detalles_Entradas.Producto_ID"
        Mi_SQL = Mi_SQL & " AND Cat_Productos.Tipo_Producto_ID='" & Rs_Consulta_Tipo_Producto.rdoColumns("Tipo_Producto_ID") & "'"
        If Cmb_Producto_Materia_Prima.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID='" & Format(Cmb_Producto_Materia_Prima.ItemData(Cmb_Producto_Materia_Prima.ListIndex), "00000") & "'"
        End If
        'Filtra por ubicacion
        If Cmb_Almacen.ListIndex > -1 Then
            If Cmb_Almacen.Text = "SIN UBICACION" Then
                Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Almacen_ID IS NULL"
            Else
                If Cmb_Almacen.Text <> "TODAS" Then
                    Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Almacen_ID='" & Format(Cmb_Almacen.ItemData(Cmb_Almacen.ListIndex), "00000") & "'"
                End If
            End If
        End If
        'Filtro de empresas
        If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Productos.Empresa IS NOT NULL"
        Else
            If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1 OR Cat_Productos.Empresa=2)"
            Else
                If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                    Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1 OR Cat_Productos.Empresa=3)"
                Else
                    If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 0 Then
                        Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=1)"
                    Else
                        If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
                            Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=2 OR Cat_Productos.Empresa=3)"
                        Else
                            If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                                Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=2)"
                            Else
                                If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                                    Mi_SQL = Mi_SQL & " AND (Cat_Productos.Empresa=3)"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Mi_SQL = Mi_SQL & " GROUP BY Cat_Productos.Producto_ID,Cat_Productos.Nombre,Cat_Productos.Costo,Cat_Productos.Codigo,Detalles_Entradas.Almacen_ID"
        'Valida si muestra con existencia
        If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 0 Then
            Mi_SQL = Mi_SQL & " HAVING SUM(Detalles_Entradas.Faltante)>0"
        Else
            If Chk_Sin_Existencia.Value = 1 And Chk_Negativos.Value = 0 Then
                Mi_SQL = Mi_SQL & " HAVING SUM(Detalles_Entradas.Faltante)>=0"
            Else
                If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 1 Then
                    Mi_SQL = Mi_SQL & " HAVING SUM(Detalles_Entradas.Faltante)<>0"
                End If
            End If
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Productos.Codigo"
        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Imprime los productos encontrados
        If Not Rs_Consulta_Producto.EOF Then
            Print #1, Spc(40); "TIPO DE PRODUCTO :     "; Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Print #1,
            Grid_Reporte.AddItem "TIPO DE PRODUCTO : " & Chr(9) & Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Grid_Reporte.AddItem ""
            With Rs_Consulta_Producto
                While Not .EOF
                    'Consulta los productos de los pedidos que tengan estatus de pendiente
                    Mi_SQL = "SELECT SUM(Cantidad) AS Pedido FROM Detalles_Pedidos"
                    Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Rs_Consulta_Producto.rdoColumns("Producto_ID") & "'"
                    Mi_SQL = Mi_SQL & " AND (Estatus='Pendiente' OR Estatus='Preautorizado' OR Estatus='Autorizado')"
                    Set Rs_Consulta_Detalles_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not IsNull(Rs_Consulta_Detalles_Pedidos.rdoColumns("Pedido")) Then
                        Existencia = Val(Rs_Consulta_Producto.rdoColumns("Existencia")) - Val(Rs_Consulta_Detalles_Pedidos.rdoColumns("Pedido"))
                    Else
                        Existencia = Val(Rs_Consulta_Producto.rdoColumns("Existencia"))
                    End If
                    Rs_Consulta_Detalles_Pedidos.Close
                    'Consulta la ubicacion
                    If Not IsNull(.rdoColumns("Almacen_ID")) Then
                        Mi_SQL = "SELECT Almacen_ID,Nombre FROM Cat_Almacenes"
                        Mi_SQL = Mi_SQL & " WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                        Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ubicacion.EOF Then
                            Ubicacion_Almacen = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                        Else
                            Ubicacion_Almacen = "N/A"
                        End If
                        Rs_Consulta_Ubicacion.Close
                    Else
                        Ubicacion_Almacen = "N/A"
                    End If
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 18); _
                        Spc(20 - Len(Mid(.rdoColumns("Codigo"), 1, 18))); Mid(.rdoColumns("Nombre"), 1, 36); _
                        Spc(38 - Len(Mid(.rdoColumns("Nombre"), 1, 36))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Existencia"), "#,##0"), 11); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Existencia, "#,##0"), 11); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo"), "#,##0.00"), 8); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Costo")) * Val(.rdoColumns("Existencia")), "#,##0.00"), 12); _
                        Spc(4); Ubicacion_Almacen
                    Grid_Reporte.AddItem .rdoColumns("Codigo") _
                        & Chr(9) & .rdoColumns("Nombre") _
                        & Chr(9) & .rdoColumns("Existencia") _
                        & Chr(9) & Existencia _
                        & Chr(9) & .rdoColumns("Costo") _
                        & Chr(9) & .rdoColumns("Costo") * .rdoColumns("Existencia") _
                        & Chr(9) & Ubicacion_Almacen
                    Total = Val(Total) + (.rdoColumns("Existencia"))
                    Costo_Almacen = Val(Costo_Almacen) + (Val(.rdoColumns("Existencia")) * Val(.rdoColumns("Costo")))
                    Rs_Consulta_Producto.MoveNext
                Wend
            End With
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
        End If
        Rs_Consulta_Producto.Close
        Rs_Consulta_Tipo_Producto.MoveNext
    Wend
    Print #1,
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    'Imprime los totales de piezas y costo
    Print #1, "                                                                            TOTAL EXISTENCIA PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,###,##0"), 13)
    Grid_Reporte.AddItem Chr(9) & "TOTAL EXISTENCIA PIEZAS" & Chr(9) & Total
    Print #1, "                                                                            COSTO INVENTARIO PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "##,###,##0.00"), 13)
    Grid_Reporte.AddItem Chr(9) & "COSTO INVENTARIO PIEZAS" & Chr(9) & Costo_Almacen
    Grid_Reporte.AddItem ""
    Print #1,
    Print #1, "                                                                            COSTO TOTAL INVENTARIO        "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(Costo_Almacen), "##,###,##0.00"), 13)
    Grid_Reporte.AddItem Chr(9) & "COSTO TOTAL INVENTARIO" & Chr(9) & Costo_Almacen
    Rs_Consulta_Tipo_Producto.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Reporte_Inventario_Materia_Prima
'DESCRIPCION   : Muestra el inventario general de materia prima
'PARAMETROS    :
'CREO          : Sergio Ulises Druán Hernández
'FECHA_CREO    : 31-Enero-201
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Reporte_Inventario_Materia_Prima()
Dim Rs_Consulta_Tipo_Producto As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Producto As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Total As Double                                 'Almacena el total de la existencia
Dim Costo_Almacen As Double                         'Almacena el total del costo del almacen
Dim Nombre_Corto As String
Dim Ubicacion_Almacen As String

    If Cmb_Tipo_Material.ListIndex > 0 Then
        Call Encabezado_Reportes("INVENTARIO GENERAL DE " & Cmb_Tipo_Material.Text, Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
    Else
        Call Encabezado_Reportes("INVENTARIO GENERAL DE INSUMOS", Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
    End If
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "  CODIGO              DESCRIPCION                                  EXISTENCIA     COSTO        TOTAL      UBICACION     "
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1,
    Grid_Reporte.AddItem "CODIGO" & Chr(9) & "DESCRIPCION" & Chr(9) & "EXISTENCIA" & Chr(9) & "UNIDAD" & Chr(9) & "COSTO" & Chr(9) & "TOTAL" & Chr(9) & "UBICACION"
    Grid_Reporte.AddItem ""
    Total = 0
    Costo_Almacen = 0
    'Consulta la lineas
    Mi_SQL = "SELECT Tipo_Producto_ID,Nombre FROM Cat_Tipos_Productos"
    'Filtra por la linea seleccionada por el usuario
    If Cmb_Tipo_Producto.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " WHERE Tipo_Producto_ID='" & Format(Cmb_Tipo_Producto.ItemData(Cmb_Tipo_Producto.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Tipo_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Hace el recorrido de los registros encontrados
    While Not Rs_Consulta_Tipo_Producto.EOF
        'Consulta los productos de la linea en Turno
        Mi_SQL = "SELECT Cat_Materias_Primas.Materia_Prima_ID,Cat_Materias_Primas.Nombre,Cat_Materias_Primas.Costo,Cat_Materias_Primas.Codigo,Cat_Unidades.Nombre_Corto,SUM(Alm_Entradas_Materia_Prima_Detalles.Faltante) AS Existencia,Alm_Entradas_Materia_Prima_Detalles.Almacen_ID"
        Mi_SQL = Mi_SQL & " FROM Cat_Materias_Primas LEFT JOIN Cat_Unidades ON Cat_Materias_Primas.Unidad_ID=Cat_Unidades.Unidad_ID"
        Mi_SQL = Mi_SQL & " INNER JOIN Alm_Entradas_Materia_Prima_Detalles ON Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " WHERE Cat_Materias_Primas.Tipo_Producto_ID='" & Rs_Consulta_Tipo_Producto.rdoColumns("Tipo_Producto_ID") & "'"
        If Cmb_Producto_Materia_Prima.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Materia_Prima_ID='" & Format(Cmb_Producto_Materia_Prima.ItemData(Cmb_Producto_Materia_Prima.ListIndex), "00000") & "'"
        End If
        'Filtra por el tipo de material
        If Cmb_Tipo_Material.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Tipo_Material='" & Cmb_Tipo_Material.Text & "'"
        End If
        'Filtra por ubicacion
        If Cmb_Almacen.ListIndex > -1 Then
            If Cmb_Almacen.Text = "SIN UBICACION" Then
                Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Almacen_ID IS NULL"
            Else
                If Cmb_Almacen.Text <> "TODAS" Then
                    Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Almacen_ID='" & Format(Cmb_Almacen.ItemData(Cmb_Almacen.ListIndex), "00000") & "'"
                End If
            End If
        End If
        'Filtro de empresas
        If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Empresa IS NOT NULL"
        Else
            If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1 OR Cat_Materias_Primas.Empresa=2)"
            Else
                If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                    Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1 OR Cat_Materias_Primas.Empresa=3)"
                Else
                    If Chk_Natural_Health.Value = 1 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 0 Then
                        Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=1)"
                    Else
                        If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 1 Then
                            Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=2 OR Cat_Materias_Primas.Empresa=3)"
                        Else
                            If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 1 And Chk_Maquilas.Value = 0 Then
                                Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=2)"
                            Else
                                If Chk_Natural_Health.Value = 0 And Chk_Pronacen.Value = 0 And Chk_Maquilas.Value = 1 Then
                                    Mi_SQL = Mi_SQL & " AND (Cat_Materias_Primas.Empresa=3)"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'Filtra la visibilidad de los productos
        If Cmb_Visibilidad.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND Cat_Materias_Primas.Clasificacion='" & Cmb_Visibilidad.Text & "'"
        End If
        Mi_SQL = Mi_SQL & " GROUP BY Cat_Materias_Primas.Materia_Prima_ID,Cat_Materias_Primas.Nombre,Cat_Materias_Primas.Existencia,Cat_Materias_Primas.Costo,Cat_Materias_Primas.Codigo,Cat_Unidades.Nombre_Corto,Alm_Entradas_Materia_Prima_Detalles.Almacen_ID"
        'Valida si muestra con existencia
        If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 0 Then
            Mi_SQL = Mi_SQL & " HAVING SUM(Alm_Entradas_Materia_Prima_Detalles.Faltante)>0"
        Else
            If Chk_Sin_Existencia.Value = 1 And Chk_Negativos.Value = 0 Then
                Mi_SQL = Mi_SQL & " HAVING SUM(Alm_Entradas_Materia_Prima_Detalles.Faltante)>=0"
            Else
                If Chk_Sin_Existencia.Value = 0 And Chk_Negativos.Value = 1 Then
                    Mi_SQL = Mi_SQL & " HAVING SUM(Alm_Entradas_Materia_Prima_Detalles.Faltante)<>0"
                End If
            End If
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Materias_Primas.Codigo"
        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Imprime los productos encontrados
        If Not Rs_Consulta_Producto.EOF Then
            Print #1, Spc(40); "TIPO DE PRODUCTO :     "; Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Print #1,
            Grid_Reporte.AddItem "TIPO DE PRODUCTO : " & Chr(9) & Rs_Consulta_Tipo_Producto.rdoColumns("Nombre")
            Grid_Reporte.AddItem ""
            With Rs_Consulta_Producto
                While Not .EOF
                    'Asigna el nombre corto de la unidad
                    If Not IsNull(.rdoColumns("Nombre_Corto")) Then
                        Nombre_Corto = .rdoColumns("Nombre_Corto")
                    Else
                        Nombre_Corto = ""
                    End If
                    'Consulta la ubicacion
                    If Not IsNull(.rdoColumns("Almacen_ID")) Then
                        Mi_SQL = "SELECT Almacen_ID,Nombre FROM Cat_Almacenes"
                        Mi_SQL = Mi_SQL & " WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                        Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ubicacion.EOF Then
                            Ubicacion_Almacen = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                        Else
                            Ubicacion_Almacen = "N/A"
                        End If
                        Rs_Consulta_Ubicacion.Close
                    Else
                        Ubicacion_Almacen = "N/A"
                    End If
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 18); _
                        Spc(20 - Len(Mid(.rdoColumns("Codigo"), 1, 18))); Mid(.rdoColumns("Nombre"), 1, 41); _
                        Spc(43 - Len(Mid(.rdoColumns("Nombre"), 1, 41))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Existencia"), "#,##0.00"), 12); _
                        Spc(1); Nombre_Corto; _
                        Spc(4 - Len(Nombre_Corto)); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo"), "#,##0.00"), 8); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Costo")) * Val(.rdoColumns("Existencia")), "#,##0.00"), 12); _
                        Spc(4); Ubicacion_Almacen
                    Grid_Reporte.AddItem .rdoColumns("Codigo") _
                        & Chr(9) & .rdoColumns("Nombre") _
                        & Chr(9) & .rdoColumns("Existencia") _
                        & Chr(9) & Nombre_Corto _
                        & Chr(9) & .rdoColumns("Costo") _
                        & Chr(9) & .rdoColumns("Costo") * .rdoColumns("Existencia") _
                        & Chr(9) & Ubicacion_Almacen
                    Total = Val(Total) + Val(.rdoColumns("Existencia"))
                    Costo_Almacen = Val(Costo_Almacen) + (Val(.rdoColumns("Existencia") * .rdoColumns("Costo")))
                    Rs_Consulta_Producto.MoveNext
                Wend
            End With
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
        End If
        Rs_Consulta_Producto.Close
        Rs_Consulta_Tipo_Producto.MoveNext
    Wend
    Rs_Consulta_Tipo_Producto.Close
    Print #1,
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    'Imprime los totales de piezas y costo
    Print #1, "                                                                       TOTAL EXISTENCIA PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "TOTAL EXISTENCIA PIEZAS" & Chr(9) & Total
    Print #1, "                                                                       COSTO INVENTARIO PIEZAS       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Costo_Almacen, "#,###,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "COSTO INVENTARIO PIEZAS" & Chr(9) & Costo_Almacen
    Grid_Reporte.AddItem ""
    Print #1,
    Print #1, "                                                                       COSTO TOTAL INVENTARIO        "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(Costo_Almacen), "##,###,##0.00"), 18)
    Grid_Reporte.AddItem Chr(9) & "COSTO TOTAL INVENTARIO" & Chr(9) & Costo_Almacen
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Entradas
    'DESCRIPCIÓN: Muestra el reporte de las entradas de mercancias
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 30-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Entradas()
Dim Mi_SQL As String
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Detalles_Entradas As rdoResultset   'Manejo de registro
Dim Proveedor As String                             'Almacena el nombre del proveedor
Dim Gran_Total As Double                            'Almacena el total de las entadas
Dim Total_Entrada As Double                         'Almacena el total parcial de la entrada consultada en turno
Dim Numero_Lote As String
Dim Fecha_Caducidad As String

    'Consulta las entradas
    Mi_SQL = "SELECT No_Entrada,Fecha,Tipo_Entrada,ISNULL(Ope_Entradas.Estatus,'ACTIVA') As Estatus,Ope_Entradas.Costo_Total,Ope_Entradas.Proveedor_ID,Cat_Proveedores.Tipo,Cat_Proveedores.Nombre AS Proveedor"
    Mi_SQL = Mi_SQL & " FROM Ope_Entradas,Cat_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Ope_Entradas.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Entradas.Fecha,Ope_Entradas.No_Entrada"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE ENTRADAS DE PRODUCTO TERMINADO", Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "dd/MMM/yyyy"), True)
    Gran_Total = 0
    Grid_Reporte.Cols = 15
    While Not Rs_Consulta_Entradas.EOF
        If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Proveedor_ID")) Then
            Proveedor = Rs_Consulta_Entradas.rdoColumns("Proveedor")
        Else
            Proveedor = ""
        End If
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "No. Entrada : "; Rs_Consulta_Entradas.rdoColumns("No_Entrada"); "   Fecha : "; Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15))); "Tipo Entrada : "; Mid(Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada"), 1, 15)
        Print #1, "Proveedor   : "; Mid(Proveedor, 1, 33); Spc(35 - Len(Mid(Proveedor, 1, 33))); "Tipo    : "; Mid(UCase(Rs_Consulta_Entradas.rdoColumns("Tipo")), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Tipo"), 1, 15))); "Costo    : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Entradas.rdoColumns("Costo_Total"), "#,###,##0.00"), 12)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("No_Entrada") & Chr(9) & "Fecha Entrada :" & Chr(9) & Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Estatus") & Chr(9) & "Tipo Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada") & Chr(9) & "Proveedor :" & Chr(9) & Proveedor & Chr(9) & "Tipo :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo") & Chr(9) & "Costo :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Costo_Total")
        'Consulta el detalle de las entradas
        Mi_SQL = "SELECT Detalles_Entradas.No_Entrada,Detalles_Entradas.Numero_Lote,Detalles_Entradas.Fecha_Caducidad,Detalles_Entradas.Cantidad,Detalles_Entradas.Costo_Compra,Cat_Productos.Codigo,Cat_Productos.Nombre"
        Mi_SQL = Mi_SQL & " FROM Detalles_Entradas,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Detalles_Entradas.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND No_Entrada='" & Rs_Consulta_Entradas.rdoColumns("No_Entrada") & "'"
        Set Rs_Consulta_Detalles_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        With Rs_Consulta_Detalles_Entradas
        Print #1, "  Codigo         Descripcion                          No. Lote              Caducidad    Cantidad    Costo       Total    "
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "No. Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        Print #1, "__________________________________________________________________________________________________________________________"
            While Not .EOF
                If Not IsNull(.rdoColumns("Numero_Lote")) Then
                    Numero_Lote = .rdoColumns("Numero_Lote")
                Else
                    Numero_Lote = ""
                End If
                If Not IsNull(.rdoColumns("Fecha_Caducidad")) Then
                    Fecha_Caducidad = Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy")
                Else
                    Fecha_Caducidad = ""
                End If
                Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                    Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 35); _
                    Spc(37 - Len(Mid(.rdoColumns("Nombre"), 1, 35))); Mid(Numero_Lote, 1, 20); _
                    Spc(22 - Len(Mid(Numero_Lote, 1, 20))); Fecha_Caducidad; _
                    Spc(13 - Len(Fecha_Caducidad)); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "####0"), 7); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Costo_Compra") * .rdoColumns("Cantidad")), "#,###,##0.00"), 12)
                Grid_Reporte.AddItem .rdoColumns("Codigo") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Numero_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo_Compra") _
                    & Chr(9) & .rdoColumns("Costo_Compra") * .rdoColumns("Cantidad")
                'Calcula el total de cada entrada
                Total_Entrada = Total_Entrada + Val(.rdoColumns("Costo_Compra") * .rdoColumns("Cantidad"))
                .MoveNext
            Wend
        End With
        'Calcula el gran total de las entradas
        Gran_Total = Gran_Total + Total_Entrada
        Rs_Consulta_Detalles_Entradas.Close
        Rs_Consulta_Entradas.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(61); "Gran Total : "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 16)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total
    Rs_Consulta_Entradas.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Entradas_Materia_Prima
    'DESCRIPCIÓN: Muestra el reporte de las entradas de materia prima
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 30-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Entradas_Materia_Prima()
Dim Mi_SQL As String
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles As rdoResultset   'Manejo de registro
Dim Proveedor As String                             'Almacena el nombre del proveedor
Dim Gran_Total As Double                            'Almacena el total de las entadas
Dim Total_Entrada As Double                         'Almacena el total parcial de la entrada consultada en turno

    'Consulta las entradas
    Mi_SQL = "SELECT No_Entrada,Fecha,Tipo_Entrada,ISNULL(Alm_Entradas_Materia_Prima.Estatus,'ACTIVA') As Estatus,Alm_Entradas_Materia_Prima.Costo_Total,Alm_Entradas_Materia_Prima.Proveedor_ID,Cat_Proveedores.Tipo,Cat_Proveedores.Nombre AS Proveedor"
    Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima,Cat_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    If Trim(Txt_Referencia.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.No_Factura='" & Trim(Txt_Referencia.Text) & "'"
    End If
    If Cmb_Proveedor_Entrada.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Proveedor_ID='" & Format(Cmb_Proveedor_Entrada.ItemData(Cmb_Proveedor_Entrada.ListIndex), "00000") & "'"
    End If
    If Trim(Cmb_Tipo_Entrada.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Tipo_Entrada='" & Trim(Cmb_Tipo_Entrada.Text) & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas_Materia_Prima.Fecha,Alm_Entradas_Materia_Prima.No_Entrada"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE ENTRADAS DE MATERIA PRIMA", Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "dd/MMM/yyyy"), True)
    Gran_Total = 0
    Grid_Reporte.Cols = 15
    While Not Rs_Consulta_Entradas.EOF
        If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Proveedor_ID")) Then
            Proveedor = Rs_Consulta_Entradas.rdoColumns("Proveedor")
        Else
            Proveedor = ""
        End If
        Print #1,
        Print #1, "__________________________________________________________________________________________________________________________"
        Print #1, "No. Entrada : "; Rs_Consulta_Entradas.rdoColumns("No_Entrada"); "   Fecha : "; Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 15))); "Tipo Entrada : "; Mid(Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada"), 1, 15)
        Print #1, "Proveedor   : "; Mid(Proveedor, 1, 33); Spc(35 - Len(Mid(Proveedor, 1, 33))); "Tipo    : "; Mid(UCase(Rs_Consulta_Entradas.rdoColumns("Tipo")), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Tipo"), 1, 15))); "Costo    : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Entradas.rdoColumns("Costo_Total"), "#,###,##0.00"), 12)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("No_Entrada") & Chr(9) & "Fecha Entrada :" & Chr(9) & Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Estatus") & Chr(9) & "Tipo Entrada :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada") & Chr(9) & "Proveedor :" & Chr(9) & Proveedor & Chr(9) & "Tipo :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo") & Chr(9) & "Costo :" & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Costo_Total")
        'Consulta el detalle de las entradas
        Mi_SQL = "SELECT Alm_Entradas_Materia_Prima_Detalles.No_Entrada,Alm_Entradas_Materia_Prima_Detalles.Numero_Lote,Alm_Entradas_Materia_Prima_Detalles.Fecha_Caducidad,Alm_Entradas_Materia_Prima_Detalles.Cantidad,Alm_Entradas_Materia_Prima_Detalles.Costo_Compra,Cat_Materias_Primas.Codigo,Cat_Materias_Primas.Nombre"
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima_Detalles,Cat_Materias_Primas"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " AND No_Entrada='" & Rs_Consulta_Entradas.rdoColumns("No_Entrada") & "'"
        Set Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        Print #1, "  Codigo         Descripcion                          No. Lote              Caducidad    Cantidad    Costo       Total    "
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "No. Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        With Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles
            While Not .EOF
                If Not IsNull(.rdoColumns("Numero_Lote")) Then
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 35); _
                        Spc(37 - Len(Mid(.rdoColumns("Nombre"), 1, 35))); Mid(.rdoColumns("Numero_Lote"), 1, 20); _
                        Spc(22 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 20))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "####0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.000"), 9); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Costo_Compra") * .rdoColumns("Cantidad")), "#,###,##0.00"), 12)
                Else
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 36); _
                        Spc(37 - Len(Mid(.rdoColumns("Nombre"), 1, 36))); ""; _
                        Spc(33); ""; _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "####0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.000"), 9); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Costo_Compra") * .rdoColumns("Cantidad")), "#,###,##0.00"), 12)
                End If
                Grid_Reporte.AddItem .rdoColumns("Codigo") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Numero_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo_Compra") _
                    & Chr(9) & .rdoColumns("Costo_Compra") * .rdoColumns("Cantidad")
                'Calcula el total de cada entrada
                Total_Entrada = Total_Entrada + Val(.rdoColumns("Costo_Compra") * .rdoColumns("Cantidad"))
                .MoveNext
            Wend
        End With
        'Calcula el gran total de las entradas
        Gran_Total = Gran_Total + Total_Entrada
        Rs_Consulta_Alm_Entradas_Materia_Prima_Detalles.Close
        Rs_Consulta_Entradas.MoveNext
    Wend
    Print #1,
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1, Spc(61); "Gran Total : "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 16)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total
    Rs_Consulta_Entradas.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Devoluciones_Cliente()
    'DESCRIPCIÓN: Muestra el reporte de las devoluciones
    'CREO: Rafael Muñoz
    'FECHA_CREO: 08-Marzo-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Devoluciones_Cliente()
Dim Rs_Consulta_Ope_Devoluciones As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Detalles_Devoluciones As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Productos As rdoResultset                   'Manejo de registro
Dim Gran_Total As Double                                    'Almacena el total de las entadas
Dim Total_Devolucion As Double                              'Almacena el total parcial de la entrada consultada en turno

    'Consulta las devoluciones
    Mi_SQL = "SELECT No_Devolucion, Fecha, Cliente_ID,No_Factura,Empresa_Factura,Tipo,Importe FROM Ope_Devoluciones"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Devolucion_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Devolucion_Fin.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Cliente_Devoluciones.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cliente_ID = '" & Format(Cmb_Cliente_Devoluciones.ItemData(Cmb_Cliente_Devoluciones.ListIndex), "00000") & "'"
    End If
    If Trim(Txt_No_Devolucion.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND No_Devolucion = '" & Format(Trim(Txt_No_Devolucion.Text), "0000000000") & "'"
    End If

    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Ope_Devoluciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE DEVOLUCIONES CLIENTES", Format(Dtp_Fecha_Devolucion_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Devolucion_Fin.Value, "dd/MMM/yyyy"), True)
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1,
    Gran_Total = 0
    While Not Rs_Consulta_Ope_Devoluciones.EOF
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "Fecha Devolucion : "; Format(Rs_Consulta_Ope_Devoluciones.rdoColumns("Fecha"), "dd/MMM/yyyy") '; Spc(23); "Tipo: "; Mid(Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada"), 1, 19)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Fecha Devolucion : " & Chr(9) & Format(Rs_Consulta_Ope_Devoluciones.rdoColumns("Fecha"), "dd/MMM/yyyy") ' & Chr(9) & "Tipo Entrada : " & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Tipo_Entrada")
        'Consulta el detalle de las entradas
        Mi_SQL = "SELECT * FROM Detalles_Devoluciones"
        Mi_SQL = Mi_SQL & " WHERE No_Devolucion='" & Rs_Consulta_Ope_Devoluciones.rdoColumns("No_Devolucion") & "'"
        Set Rs_Consulta_Detalles_Devoluciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Devolucion = 0
        With Rs_Consulta_Detalles_Devoluciones
        Print #1, "Devolucion  Codigo            Descripcion                             Cantidad    Precio         Total"
        Grid_Reporte.AddItem "No Devolucion" & Chr(9) & "Codigo Barras" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Total"
        Print #1, ""
        Print #1, "________________________________________________________________________________________________________________________"
            While Not .EOF
                'Consulta la descripcion del producto
                Mi_SQL = "SELECT Producto_ID, Codigo, Nombre FROM Cat_Productos"
                Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & .rdoColumns("Producto_ID") & "'"
                Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1, .rdoColumns("No_Devolucion"); Spc(2); Mid(Rs_Consulta_Productos.rdoColumns("Codigo"), 1, 17); Spc(18 - Len(Mid(Rs_Consulta_Productos.rdoColumns("Codigo"), 1, 17))); Mid(Rs_Consulta_Productos.rdoColumns("Nombre"), 1, 36); Spc(37 - Len(Mid(Rs_Consulta_Productos.rdoColumns("Nombre"), 1, 36))); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "####"), 5); Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Precio"), "###.00"), 6); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Precio") * .rdoColumns("Cantidad")), "#,###,###.00"), 12)
                Grid_Reporte.AddItem .rdoColumns("No_Devolucion") & Chr(9) & Rs_Consulta_Productos.rdoColumns("Codigo") & Chr(9) & Rs_Consulta_Productos.rdoColumns("Nombre") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Precio") & Chr(9) & .rdoColumns("Precio") * .rdoColumns("Cantidad")
                'Calcula el total de cada entrada
                Total_Devolucion = Total_Devolucion + Val(.rdoColumns("Precio") * .rdoColumns("Cantidad"))
                Rs_Consulta_Productos.Close
                .MoveNext
            Wend
        End With
        Print #1, Spc(50); "                                   __________________ "
        Print #1, Spc(53); "Total de la Devolución: "; Spc(13); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Devolucion, "#,###,###.00"), 12)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total de la Devolucion : " & Chr(9) & Total_Devolucion
        'Calcula el gran total de las entradas
        Gran_Total = Gran_Total + Total_Devolucion
        Rs_Consulta_Detalles_Devoluciones.Close
        Rs_Consulta_Ope_Devoluciones.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(61); "Gran Total : "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,###.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "         Gran Total :       " & Chr(9) & Gran_Total
    Rs_Consulta_Ope_Devoluciones.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Saldos_Cliente
    'DESCRIPCIÓN: Muestra el global de saldos del cliente con vencimientos
    '             30, 60, 90 y mas días en facturas y/o remisiones
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Saldos_Cliente()
Dim Cliente As String                                       'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Cat_Ciudades As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Remisiones_Clientes As rdoResultset      'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_Factura As Double                                 'Almacena el total de la facura
Dim Total_Saldo As Double                                   'Almacena el total del saldo
Dim Total_Abono As Double                                   'Almacena el total de los abonos
Dim Fecha_Vencimiento As Date                               'Almacena la fecha de vencimiento de la factura
Dim Vencimiento As Integer                                  'Almacena los dias de vnecimiento de la factura
Dim Total_Vencimiento As Integer                            'Almacena los espacios para colocar la información
Dim Espacio As String
Dim Total_0_30 As Double
Dim Total_30_60 As Double
Dim Total_60_90 As Double
Dim Total_90 As Double
Dim Gran_Total As Double
Dim Total_Cliente As Double
Dim Abono_Cliente As Double
Dim Saldo_Cliente As Double
Dim Cliente_ID As String
Dim Tipo_Pago As String

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    If Cmb_Vendedor_Saldo.ListIndex > -1 Or Cmb_Cliente_Saldo.ListIndex > -1 Then
        If Chk_Fechas_Saldos.Value = 0 Then
            Call Encabezado_Reportes("REPORTE DE SALDO POR CLIENTE", Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
        Else
            Call Encabezado_Reportes("REPORTE DE SALDO POR CLIENTE", Format(Dtp_Fecha_Saldo_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Saldo_Fin.Value, "dd/MMM/yyyy"), True)
        End If
        If Cmb_Vendedor_Saldo.ListIndex > -1 Then
            Print #1,
            Print #1, Conectar_Ayudante.Centra("VENDEDOR: " & Cmb_Vendedor_Saldo.Text, Len("-------------------------------------------------------------------------------------------------------------------------"))
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "VENDEDOR" & Chr(9) & Cmb_Vendedor_Saldo.Text
        End If
        If Cmb_Remision_Factura.ListIndex = 0 Or Cmb_Remision_Factura.ListIndex = 2 Then
            'Consulta las facturas que fueron emitidas al cliente seleccionado por el usuario
            Mi_SQL = "SELECT No_Factura,Empresa,Adm_Facturas_Clientes.Cliente_ID,Fecha,Total,Saldo,Abono,Cancelada,Cat_Clientes.Vendedor_ID,Tipo_Pago"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes,Cat_Vendedores"
            Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
            Mi_SQL = Mi_SQL & " AND Pagada='N'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            If Cmb_Empresa_Saldo.ListIndex = 0 Then
                Mi_SQL = Mi_SQL & " AND Empresa=1"
            Else
                If Cmb_Empresa_Saldo.ListIndex = 1 Then
                    Mi_SQL = Mi_SQL & " AND Empresa=2"
                End If
            End If
            If Cmb_Vendedor_Saldo.ListIndex > -1 Then
                Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID='" & Format(Cmb_Vendedor_Saldo.ItemData(Cmb_Vendedor_Saldo.ListIndex), "00000") & "'"
            End If
            If Cmb_Cliente_Saldo.ListIndex > -1 Then
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
            End If
            If Chk_Fechas_Saldos.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
            End If
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Facturas_Clientes.Cliente_ID,Empresa,Fecha"
            Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Adm_Remisiones_Clientes.EOF
                Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Clientes.Telefono,Cat_Clientes.Vendedor_ID,Cat_Vendedores.Nombre AS Vendedor"
                Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Vendedores"
                Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
                Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID='" & Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    With Rs_Consulta_Cat_Clientes
                        If Cliente_ID <> Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") Then
                            'Consulta la ciudad del cliente
                            If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID")) Then
                                Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                                Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                                Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                                Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID") & "'"
                                Set Rs_Consulta_Cat_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                If Not Rs_Consulta_Cat_Ciudades.EOF Then
                                    Ciudad = Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre_Corto")
                                Else
                                    Ciudad = ""
                                End If
                                Rs_Consulta_Cat_Ciudades.Close
                            Else
                                Ciudad = ""
                            End If
                            If Cliente_ID <> "" Then
                                Print #1, "_________________________________________________________________________________________________________________________"
                                Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
                                          Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
                                          Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
                                Print #1,
                                Grid_Reporte.AddItem ""
                                Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                                Grid_Reporte.AddItem ""
                            End If
                            Print #1,
                            Print #1, "_________________________________________________________________________________________________________________________"
                            Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
                            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion"); " "; .rdoColumns("Colonia")
                            Print #1, Spc(2); "Ciudad:    "; Ciudad; Spc(2); "CP: "; .rdoColumns("CP"); Spc(2); "Telefono: "; .rdoColumns("Telefono")
                            Print #1, Spc(2); "(*) Facturas de contado"
                            Grid_Reporte.AddItem ""
                            Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                            Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion") & Chr(9) & .rdoColumns("Colonia")
                            Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP") & Chr(9) & "Telefono" & Chr(9) & .rdoColumns("Telefono")
                            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                            Print #1, " Factura     Fecha         Total        Abono          Saldo           0-30         30-60         60-90          +90     "
                            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                            Grid_Reporte.AddItem ""
                            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "0-30" & Chr(9) & "30-60" & Chr(9) & "60-90" & Chr(9) & "'+90"
                            Cliente_ID = Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")
                            Total_Cliente = 0
                            Saldo_Cliente = 0
                            Abono_Cliente = 0
                        End If
                    End With
                End If
                With Rs_Consulta_Adm_Remisiones_Clientes
                    If Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                    If Trim(.rdoColumns("Tipo_Pago")) = "CONTADO" Then
                        Tipo_Pago = "*"
                    Else
                        Tipo_Pago = " "
                    End If
                    Fecha_Vencimiento = Format(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy")
                    Vencimiento = DateDiff("d", Format(Fecha_Vencimiento, "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                    If Vencimiento <= 30 Then
                        Total_0_30 = Total_0_30 + .rdoColumns("Saldo")
                        Total_Vencimiento = 2
                        Espacio = ""
                    Else
                        If Vencimiento >= 31 And Vencimiento <= 60 Then
                            Total_30_60 = Total_30_60 + .rdoColumns("Saldo")
                            Total_Vencimiento = 16
                            Espacio = Chr(9)
                        Else
                            If Vencimiento >= 61 And Vencimiento <= 90 Then
                                Total_60_90 = Total_60_90 + .rdoColumns("Saldo")
                                Total_Vencimiento = 30
                                Espacio = Chr(9) & Chr(9)
                            Else
                                If Vencimiento >= 91 Then
                                    Total_90 = Total_90 + .rdoColumns("Saldo")
                                    Total_Vencimiento = 44
                                    Espacio = Chr(9) & Chr(9) & Chr(9)
                                End If
                            End If
                        End If
                    End If
                    Print #1, Format(.rdoColumns("No_Factura"), "00000"); Spc(1); Empresa; Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                              Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                              Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                              Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                              Spc(Total_Vencimiento); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Espacio & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00")
                    Total_Factura = Total_Factura + .rdoColumns("Total")
                    Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                    Total_Abono = Total_Abono + .rdoColumns("Abono")
                    Total_Cliente = Total_Cliente + .rdoColumns("Total")
                    Saldo_Cliente = Saldo_Cliente + .rdoColumns("Saldo")
                    Abono_Cliente = Abono_Cliente + .rdoColumns("Abono")
                End With
                Rs_Consulta_Adm_Remisiones_Clientes.MoveNext
            Wend
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
                      Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
            Grid_Reporte.AddItem ""
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, " TOTALES FACTURAS  "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Factura, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Abono, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_0_30, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_30_60, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_60_90, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_90, "#,###,##0.00"), 12)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES FACTURAS" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Factura, "#,###,##0.00") & Chr(9) & Format(Total_Abono, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00") & Chr(9) & Format(Total_0_30, "#,###,##0.00") & Chr(9) & _
                        Format(Total_30_60, "#,###,##0.00") & Chr(9) & Format(Total_60_90, "#,###,##0.00") & Chr(9) & Format(Total_90, "#,###,##0.00")
            Rs_Consulta_Adm_Remisiones_Clientes.Close
        End If
        If Cmb_Remision_Factura.ListIndex = 1 Or Cmb_Remision_Factura.ListIndex = 2 Then
            Cliente_ID = ""
            Total_Factura = 0
            Total_Abono = 0
            Total_Saldo = 0
            Total_0_30 = 0
            Total_30_60 = 0
            Total_60_90 = 0
            Total_90 = 0
            'Consulta las remisiones que fueron emitidas al cliente seleccionado por el usuario
            Mi_SQL = "SELECT No_Remision,Empresa,Adm_Remisiones_Clientes.Cliente_ID,Fecha,Total,Saldo,Abono,Cancelada,Cat_Clientes.Vendedor_ID,Tipo_Pago"
            Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes,Cat_Vendedores"
            Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
            Mi_SQL = Mi_SQL & " AND Pagada='N'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            If Cmb_Empresa_Saldo.ListIndex = 0 Then
                Mi_SQL = Mi_SQL & " AND Empresa=1"
            Else
                If Cmb_Empresa_Saldo.ListIndex = 1 Then
                    Mi_SQL = Mi_SQL & " AND Empresa=2"
                End If
            End If
            If Cmb_Vendedor_Saldo.ListIndex > -1 Then
                Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID='" & Format(Cmb_Vendedor_Saldo.ItemData(Cmb_Vendedor_Saldo.ListIndex), "00000") & "'"
            End If
            If Cmb_Cliente_Saldo.ListIndex > -1 Then
                Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cliente_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
            End If
            If Chk_Fechas_Saldos.Value = 1 Then
                Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
            End If
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Remisiones_Clientes.Cliente_ID,Empresa,Fecha"
            Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Adm_Remisiones_Clientes.EOF
                Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Clientes.Telefono,Cat_Clientes.Vendedor_ID,Cat_Vendedores.Nombre AS Vendedor"
                Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Vendedores"
                Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
                Mi_SQL = Mi_SQL & " AND Cat_Clientes.Cliente_ID='" & Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    With Rs_Consulta_Cat_Clientes
                        If Cliente_ID <> Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID") Then
                            'Consulta la ciudad del cliente
                            If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID")) Then
                                Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                                Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                                Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                                Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID") & "'"
                                Set Rs_Consulta_Cat_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                If Not Rs_Consulta_Cat_Ciudades.EOF Then
                                    Ciudad = Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre_Corto")
                                Else
                                    Ciudad = ""
                                End If
                                Rs_Consulta_Cat_Ciudades.Close
                            Else
                                Ciudad = ""
                            End If
                            If Cliente_ID <> "" Then
                                Print #1, "_________________________________________________________________________________________________________________________"
                                Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
                                          Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
                                          Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
                                Print #1,
                                Grid_Reporte.AddItem ""
                                Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                                Grid_Reporte.AddItem ""
                            End If
                            Print #1,
                            Print #1, "_________________________________________________________________________________________________________________________"
                            Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
                            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion"); " "; .rdoColumns("Colonia")
                            Print #1, Spc(2); "Ciudad:    "; Ciudad; Spc(2); "CP: "; .rdoColumns("CP"); Spc(2); "Telefono: "; .rdoColumns("Telefono")
                            Print #1, Spc(2); "(*) Facturas de contado"
                            Grid_Reporte.AddItem ""
                            Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                            Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion") & Chr(9) & .rdoColumns("Colonia")
                            Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP") & Chr(9) & "Telefono" & Chr(9) & .rdoColumns("Telefono")
                            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                            Print #1, " Remision    Fecha         Total        Abono          Saldo           0-30         30-60         60-90          +90     "
                            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                            Grid_Reporte.AddItem ""
                            Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "0-30" & Chr(9) & "30-60" & Chr(9) & "60-90" & Chr(9) & "'+90"
                            Cliente_ID = Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Cliente_ID")
                            Total_Cliente = 0
                            Saldo_Cliente = 0
                            Abono_Cliente = 0
                        End If
                    End With
                End If
                With Rs_Consulta_Adm_Remisiones_Clientes
                    If Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                    If Trim(.rdoColumns("Tipo_Pago")) = "CONTADO" Then
                        Tipo_Pago = "*"
                    Else
                        Tipo_Pago = " "
                    End If
                    Fecha_Vencimiento = Format(Rs_Consulta_Adm_Remisiones_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy")
                    Vencimiento = DateDiff("d", Format(Fecha_Vencimiento, "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                    If Vencimiento <= 30 Then
                        Total_0_30 = Total_0_30 + .rdoColumns("Saldo")
                        Total_Vencimiento = 2
                        Espacio = ""
                    Else
                        If Vencimiento >= 31 And Vencimiento <= 60 Then
                            Total_30_60 = Total_30_60 + .rdoColumns("Saldo")
                            Total_Vencimiento = 16
                            Espacio = Chr(9)
                        Else
                            If Vencimiento >= 61 And Vencimiento <= 90 Then
                                Total_60_90 = Total_60_90 + .rdoColumns("Saldo")
                                Total_Vencimiento = 30
                                Espacio = Chr(9) & Chr(9)
                            Else
                                If Vencimiento >= 91 Then
                                    Total_90 = Total_90 + .rdoColumns("Saldo")
                                    Total_Vencimiento = 44
                                    Espacio = Chr(9) & Chr(9) & Chr(9)
                                End If
                            End If
                        End If
                    End If
                    Print #1, Format(.rdoColumns("No_Remision"), "00000"); Spc(1); Empresa; Tipo_Pago; Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                              Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 13); _
                              Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                              Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 13); _
                              Spc(Total_Vencimiento); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Espacio & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00")
                    Total_Factura = Total_Factura + .rdoColumns("Total")
                    Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                    Total_Abono = Total_Abono + .rdoColumns("Abono")
                    Total_Cliente = Total_Cliente + .rdoColumns("Total")
                    Saldo_Cliente = Saldo_Cliente + .rdoColumns("Saldo")
                    Abono_Cliente = Abono_Cliente + .rdoColumns("Abono")
                End With
                Rs_Consulta_Adm_Remisiones_Clientes.MoveNext
            Wend
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, " TOTALES CLIENTE "; Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Abono_Cliente, "#,###,##0.00"), 12); _
                      Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 12)
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00") & Chr(9) & Format(Abono_Cliente, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
            Grid_Reporte.AddItem ""
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, " TOTALES REMISIONES"; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Factura, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Abono, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 13); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_0_30, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_30_60, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_60_90, "#,###,##0.00"), 12); _
                      Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_90, "#,###,##0.00"), 12)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES REMISIONES" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Factura, "#,###,##0.00") & Chr(9) & Format(Total_Abono, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00") & Chr(9) & Format(Total_0_30, "#,###,##0.00") & Chr(9) & _
                        Format(Total_30_60, "#,###,##0.00") & Chr(9) & Format(Total_60_90, "#,###,##0.00") & Chr(9) & Format(Total_90, "#,###,##0.00")
            Rs_Consulta_Adm_Remisiones_Clientes.Close
        End If
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "Seleccione un vendedor o un cliente de la lista", vbExclamation
        Cmb_Vendedor_Saldo.SetFocus
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Salidas
    'DESCRIPCIÓN: Muestra el reporte de las salidas a su costo de compra, así
    '             como su precio de venta y totales
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 30-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Salidas()
Dim Mi_SQL As String
Dim Rs_Consulta_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Detalles_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset    'Manejo de registro
Dim Cliente As String                       'Almacena el nombre del cliente
Dim Total_Costos As Double                  'Almacena el total de los costos por salida
Dim Total_Venta As Double                   'Almacena el total de la venta por salida
Dim Gran_Total_Costos As Double             'Almacena el gran total de los costos de las salidas
Dim Gran_Total_Venta As Double              'Almacena el gran total de la venta por las salidas
Dim Gran_Total_Utilidad As Double           'Almacena el gran total de la utilidades de las salidas

    'Consulta las salidas de acuerdo a las fechas proporcionadas
    Mi_SQL = "SELECT DISTINCT No_Salida,Fecha,Tipo,ISNULL(No_Pedido,'') AS No_Pedido,Cliente_ID,ISNULL(Estatus,'ACTIVA') AS Estatus,Nombre_Solicito"
    Mi_SQL = Mi_SQL & " FROM Ope_Salidas"
    Mi_SQL = Mi_SQL & " WHERE No_Salida IS NOT NULL"
    'Si seleccionó algún tipo de salida
    If Trim(Cmb_Tipo_Entrada.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Tipo='" & Trim(Cmb_Tipo_Entrada.Text) & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Salida"
    Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE SALIDAS DE PRODUCTO TERMINADO", Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "dd/MMM/yyyy"), True)
    Gran_Total_Costos = 0
    Gran_Total_Venta = 0
    Gran_Total_Utilidad = 0
    While Not Rs_Consulta_Salidas.EOF
        If Not IsNull(Rs_Consulta_Salidas.rdoColumns("Cliente_ID")) Then
            Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Salidas.rdoColumns("Cliente_ID") & "'"
            Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Salidas.EOF Then
                Cliente = Rs_Consulta_Clientes.rdoColumns("Nombre")
            Else
                Cliente = ""
            End If
            Rs_Consulta_Clientes.Close
        Else
            Cliente = ""
        End If
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "No. Salida  : "; Rs_Consulta_Salidas.rdoColumns("No_Salida"); "   Fecha : "; Format(Rs_Consulta_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Consulta_Salidas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Consulta_Salidas.rdoColumns("Estatus"), 1, 15))); "Tipo Salida : "; Mid(Rs_Consulta_Salidas.rdoColumns("Tipo"), 1, 15)
        Print #1, "Cliente     : "; Mid(Cliente, 1, 33); Spc(35 - Len(Mid(Cliente, 1, 33))); "No. Pedido : "; Rs_Consulta_Salidas.rdoColumns("No_Pedido"); "    Solicito  : "; Mid(Rs_Consulta_Salidas.rdoColumns("Nombre_Solicito"), 1, 30)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Salida : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("No_Salida") & Chr(9) & "Fecha : " & Chr(9) & Format(Rs_Consulta_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Estatus") & Chr(9) & "Tipo Salida : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Tipo") & Chr(9) & "Cliente : " & Chr(9) & Cliente & Chr(9) & "No. Pedido : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("No_Pedido") & Chr(9) & "Solicito : " & Chr(9) & Rs_Consulta_Salidas.rdoColumns("Nombre_Solicito")
        'Consulta la descripcion de las salidas
        Mi_SQL = "SELECT Detalles_Salidas.No_Salida,Detalles_Salidas.Producto_ID,Detalles_Salidas.Numero_Lote,Detalles_Salidas.Fecha_Caducidad,Detalles_Salidas.Cantidad,Detalles_Salidas.Costo_Compra,Detalles_Salidas.Precio_Venta,Cat_Productos.Codigo,Cat_Productos.Nombre"
        Mi_SQL = Mi_SQL & " FROM Detalles_Salidas,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Detalles_Salidas.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND No_Salida='" & Rs_Consulta_Salidas.rdoColumns("No_Salida") & "'"
        Set Rs_Detalles_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Costos = 0
        Total_Venta = 0
        Print #1, "  Codigo          Descripcion                   No. Lote             Caducidad   Cantidad    Costo      Total     Venta "
        Print #1, "________________________________________________________________________________________________________________________"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Numero_Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total Compra" & Chr(9) & "Precio Venta" & Chr(9) & "Total Venta" & Chr(9) & "Diferencia"
        With Rs_Detalles_Salidas
            While Not .EOF
                If Not IsNull(.rdoColumns("Numero_Lote")) Then
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                        Spc(16 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 29); _
                        Spc(31 - Len(Mid(.rdoColumns("Nombre"), 1, 29))); Mid(.rdoColumns("Numero_Lote"), 1, 20); _
                        Spc(21 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 20))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Precio_Venta"), "#,##0.00"), 8)
                Else
                    Print #1, Mid(.rdoColumns("Codigo"), 1, 15); _
                        Spc(16 - Len(Mid(.rdoColumns("Codigo"), 1, 15))); Mid(.rdoColumns("Nombre"), 1, 29); _
                        Spc(31 - Len(Mid(.rdoColumns("Nombre"), 1, 29))); ""; _
                        Spc(32); ""; _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Precio_Venta"), "#,##0.00"), 8)
                End If
                Grid_Reporte.AddItem .rdoColumns("Codigo") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Numero_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo_Compra") _
                    & Chr(9) & Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")) _
                    & Chr(9) & .rdoColumns("Precio_Venta") _
                    & Chr(9) & Val(.rdoColumns("Cantidad") * .rdoColumns("Precio_Venta")) _
                    & Chr(9) & Val(.rdoColumns("Cantidad") * (.rdoColumns("Precio_Venta") - .rdoColumns("Costo_Compra")))
                Total_Costos = Total_Costos + Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra"))
                Total_Venta = Total_Venta + Val(.rdoColumns("Cantidad") * .rdoColumns("Precio_Venta"))
                .MoveNext
            Wend
        End With
        Print #1,
        'Calcula el gran total de todas las salidas
        Gran_Total_Costos = Gran_Total_Costos + Total_Costos
        Gran_Total_Venta = Gran_Total_Venta + Total_Venta
        Rs_Detalles_Salidas.Close
        Rs_Consulta_Salidas.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(44); "Gran Total   : "; Spc(16); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total_Costos, "#,###,##0.00"), 12); Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total_Venta, "#,###,##0.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total_Costos & Chr(9) & Gran_Total_Venta
    Rs_Consulta_Salidas.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Reporte_Salidas_Materia_Prima
'DESCRIPCION   : Muestra el reporte de las salidas de acuerdo al tipo en referencia
'PARAMETROS    :
'CREO          : Sergio Ulises Durán Hernández
'FECHA_CREO    : 02-Octubre-2008
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Reporte_Salidas_Materia_Prima()
Dim Mi_SQL As String
Dim Rs_Alm_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Alm_Salidas_Detalles As rdoResultset     'Manejo de registro
Dim Rs_Pro_Ordenes_Produccion As rdoResultset    'Manejo de registro
Dim Cliente As String                       'Almacena el nombre del cliente
Dim Total_Costos As Double                  'Almacena el total de los costos por salida
Dim Gran_Total_Costos As Double             'Almacena el gran total de los costos de las salidas

    'Consulta las salidas de acuerdo a las fechas proporcionadas
    Mi_SQL = "SELECT No_Salida,Fecha,Tipo_Salida,Nombre_Solicito,Estatus,Orden_Produccion_ID"
    Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Materia_Prima"
    Mi_SQL = Mi_SQL & " WHERE Fecha BETWEEN '" & Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "MM/dd/yyyy") & "'"
    If Cmb_Tipo_Entrada.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND Tipo_Salida='" & Cmb_Tipo_Entrada.Text & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Salida"
    Set Rs_Alm_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE SALIDAS DE MATERIA PRIMA", Format(Dtp_Fecha_Entrada_Salida_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Entrada_Salida_Fin.Value, "dd/MMM/yyyy"), True)
    Gran_Total_Costos = 0
    While Not Rs_Alm_Salidas.EOF
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "No. Salida  : "; Rs_Alm_Salidas.rdoColumns("No_Salida"); "   Fecha : "; Format(Rs_Alm_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy"); "   Estatus : "; Mid(Rs_Alm_Salidas.rdoColumns("Estatus"), 1, 15); Spc(17 - Len(Mid(Rs_Alm_Salidas.rdoColumns("Estatus"), 1, 15))); "Tipo Salida  : "; Mid(Rs_Alm_Salidas.rdoColumns("Tipo_Salida"), 1, 15)
        If Not IsNull(Rs_Alm_Salidas.rdoColumns("Orden_Produccion_ID")) Then
            'Consulta el número de orden de producción
            Mi_SQL = "SELECT Orden_Produccion_ID,Orden_Produccion FROM Pro_Orden_Produccion"
            Mi_SQL = Mi_SQL & " WHERE Orden_Produccion_ID='" & Rs_Alm_Salidas.rdoColumns("Orden_Produccion_ID") & "'"
            Set Rs_Pro_Ordenes_Produccion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Pro_Ordenes_Produccion.EOF Then
                Print #1, "Solicito    : "; Mid(Rs_Alm_Salidas.rdoColumns("Nombre_Solicito"), 1, 50); Spc(62 - Len(Mid(Rs_Alm_Salidas.rdoColumns("Nombre_Solicito"), 1, 50))); "No. Orden    : "; Rs_Pro_Ordenes_Produccion.rdoColumns("Orden_Produccion")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("No_Salida") & Chr(9) & "Fecha : " & Chr(9) & Format(Rs_Alm_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Estatus") & Chr(9) & "Tipo Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Tipo_Salida") & Chr(9) & "Solicito : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Nombre_Solicito") & Chr(9) & "No. Orden : " & Chr(9) & Rs_Pro_Ordenes_Produccion.rdoColumns("orden_Produccion")
            Else
                Print #1, "Solicito    : "; Mid(Rs_Alm_Salidas.rdoColumns("Nombre_Solicito"), 1, 50)
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("No_Salida") & Chr(9) & "Fecha : " & Chr(9) & Format(Rs_Alm_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Estatus") & Chr(9) & "Tipo Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Tipo_Salida") & Chr(9) & "Solicito : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Nombre_Solicito")
            End If
            Rs_Pro_Ordenes_Produccion.Close
        Else
            Print #1, "Solicito    : "; Mid(Rs_Alm_Salidas.rdoColumns("Nombre_Solicito"), 1, 50)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("No_Salida") & Chr(9) & "Fecha : " & Chr(9) & Format(Rs_Alm_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Estatus : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Estatus") & Chr(9) & "Tipo Salida : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Tipo_Salida") & Chr(9) & "Solicito : " & Chr(9) & Rs_Alm_Salidas.rdoColumns("Nombre_Solicito")
        End If
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Consulta la descripcion de las salidas
        Mi_SQL = "SELECT Alm_Salidas_Materia_Prima_Detalles.No_Salida,Alm_Salidas_Materia_Prima_Detalles.Materia_Prima_ID,Alm_Salidas_Materia_Prima_Detalles.Numero_Lote,Alm_Salidas_Materia_Prima_Detalles.Fecha_Caducidad,Alm_Salidas_Materia_Prima_Detalles.Cantidad,Alm_Salidas_Materia_Prima_Detalles.Costo_Compra,Cat_Materias_Primas.Codigo,Cat_Materias_Primas.Nombre"
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Materia_Prima_Detalles,Cat_Materias_Primas"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " AND No_Salida='" & Rs_Alm_Salidas.rdoColumns("No_Salida") & "'"
        Set Rs_Alm_Salidas_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Costos = 0
        Print #1, "  Codigo          Descripcion                   No. Lote        Caducidad      Cantidad    Costo      Total           "
        Print #1, "________________________________________________________________________________________________________________________"
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Numero_Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & "Total Compra"
        With Rs_Alm_Salidas_Detalles
            While Not .EOF
                Print #1, Mid(.rdoColumns("Codigo"), 1, 16); _
                    Spc(17 - Len(Mid(.rdoColumns("Codigo"), 1, 16))); Mid(.rdoColumns("Nombre"), 1, 28); _
                    Spc(30 - Len(Mid(.rdoColumns("Nombre"), 1, 28))); Mid(.rdoColumns("Numero_Lote"), 1, 15); _
                    Spc(17 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 20))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yy"); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "####0.00"), 12); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,##0.00"), 10)
                Grid_Reporte.AddItem .rdoColumns("Codigo") _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Numero_Lote") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Cantidad") _
                    & Chr(9) & .rdoColumns("Costo_Compra") _
                    & Chr(9) & Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra"))
                Total_Costos = Total_Costos + Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra"))
                .MoveNext
            Wend
        End With
        Print #1,
        'Calcula el gran total de todas las salidas
        Gran_Total_Costos = Gran_Total_Costos + Total_Costos
        Rs_Alm_Salidas_Detalles.Close
        Rs_Alm_Salidas.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(44); "Gran Total   : "; Spc(16); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total_Costos, "#,###,##0.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total_Costos
    Rs_Alm_Salidas.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Pedidos_Facturas
    'DESCRIPCIÓN: Muestra el reporte de las pedidos con las respectivas facturas
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 24-Octubre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Pedidos_Facturas()
Dim Mi_SQL As String
Dim Rs_Consulta_Ope_Pedidos As rdoResultset
Dim Rs_Consulta_Salidas_Facturas As rdoResultset
Dim Empresa As String
Dim Total_General As Double

    'Coloca el encabezado de los reportes
    Call Encabezado_Reportes("REPORTE DE PEDIDOS POR FACTURA/REMISION", Format(Dtp_Fecha_Inicio_Pedido_Factura.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedido_Factura.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    'Consulta los pedidos de acuerdo al rango de fechas proporcionadas
    Mi_SQL = "SELECT No_Pedido,Fecha_Pedido,Ope_Pedidos.Estatus,Ope_Pedidos.Cliente_ID,Cat_Clientes.Nombre AS Cliente"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido BETWEEN " & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Factura.Value, "MM/dd/yyyy") & Par_Fecha & " AND " & Par_Fecha & Format(Dtp_Fecha_Fin_Pedido_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY No_Pedido"
    Set Rs_Consulta_Ope_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Ope_Pedidos.EOF
        Mi_SQL = "SELECT DISTINCT Ope_Salidas.No_Salida,Adm_Facturas_Clientes.No_Factura,Adm_Facturas_Clientes.Empresa,Adm_Facturas_Clientes.Total"
        Mi_SQL = Mi_SQL & " FROM Ope_Salidas,Detalles_Salidas,Adm_Facturas_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Ope_Salidas.No_Salida=Detalles_Salidas.No_Salida"
        Mi_SQL = Mi_SQL & " AND Detalles_Salidas.No_Factura=Adm_Facturas_Clientes.No_Factura"
        Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Empresa_Factura=Adm_Facturas_Clientes.Empresa"
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.No_Pedido='" & Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido") & "'"
        Set Rs_Consulta_Salidas_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Salidas_Facturas.EOF Then
            If Chk_Pedido_Facturado.Value = 1 Then
                'Imprime el encabezado del pedido
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10); Spc(12 - Len(Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente"), 1, 30)
                Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus") & Chr(9) & "Cliente:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "    No. Factura          Empresa                    Importe"
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Importe"
                Print #1, "________________________________________________________________________________________________________________________"
                With Rs_Consulta_Salidas_Facturas
                    While Not .EOF
                        If .rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                        Print #1, Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(.rdoColumns("No_Factura"), 7); Spc(3), Empresa; Spc(20); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Total")), "#,###,##0.00"), 7)
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                        Total_General = Total_General + .rdoColumns("Total")
                        .MoveNext
                    Wend
                End With
                Rs_Consulta_Salidas_Facturas.Close
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(30); "GRAN TOTAL:       "; Format(Total_General, "#,###,##0.00")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_General, "#,###,##0.00")
            End If
        Else    'Checa si lo tiene en remisiones
            Mi_SQL = "SELECT DISTINCT Ope_Salidas.No_Salida,Adm_Remisiones_Clientes.No_Remision,Adm_Remisiones_Clientes.Empresa,Adm_Remisiones_Clientes.Total"
            Mi_SQL = Mi_SQL & " FROM Ope_Salidas,Detalles_Salidas,Adm_Remisiones_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Ope_Salidas.No_Salida=Detalles_Salidas.No_Salida"
            Mi_SQL = Mi_SQL & " AND Detalles_Salidas.No_Remision=Adm_Remisiones_Clientes.No_Remision"
            Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Empresa_Factura=Adm_Remisiones_Clientes.Empresa"
            Mi_SQL = Mi_SQL & " AND Ope_Salidas.No_Pedido='" & Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido") & "'"
            Set Rs_Consulta_Salidas_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Salidas_Facturas.EOF Then
                If Chk_Pedido_Facturado.Value = 1 Then
                    'Imprime el encabezado del pedido
                    Print #1,
                    Grid_Reporte.AddItem ""
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10); Spc(12 - Len(Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente"), 1, 30)
                    Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus") & Chr(9) & "Cliente:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente")
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "   No. Remisión          Empresa                    Importe"
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Importe"
                    Print #1, "________________________________________________________________________________________________________________________"
                    With Rs_Consulta_Salidas_Facturas
                        While Not .EOF
                            If .rdoColumns("Empresa") = 1 Then
                                Empresa = "NHE"
                            Else
                                Empresa = "PNC"
                            End If
                            Print #1, Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(.rdoColumns("No_Remision"), 7); Spc(3), Empresa; Spc(20); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Total")), "#,###,##0.00"), 7)
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                            Total_General = Total_General + .rdoColumns("Total")
                            .MoveNext
                        Wend
                    End With
                    Rs_Consulta_Salidas_Facturas.Close
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Spc(30); "GRAN TOTAL:       "; Format(Total_General, "#,###,##0.00")
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_General, "#,###,##0.00")
                End If
            Else
                If Chk_Pedido_Sin_Factura.Value = 1 Then
                    'Imprime el encabezado del pedido
                    Print #1,
                    Grid_Reporte.AddItem ""
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "No. Pedido: "; Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido"); Spc(5); "Fecha Pedido: "; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(2); "Estatus: "; Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10); Spc(12 - Len(Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus"), 1, 10))); "Cliente: "; Spc(1); Mid(Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente"), 1, 30)
                    Grid_Reporte.AddItem "No. Pedido:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("No_Pedido") & Chr(9) & "Fecha Pedido:" & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Estatus:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Estatus") & Chr(9) & "Cliente:" & Chr(9) & Rs_Consulta_Ope_Pedidos.rdoColumns("Cliente")
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "    No. Factura          Empresa                    Importe"
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Importe"
                    Print #1, "________________________________________________________________________________________________________________________"
                    Print #1, Spc(2); "SIN FACTURAS NI REMISIONES"
                    Grid_Reporte.AddItem "SIN FACTURAS"
                End If
            End If
        End If
        Rs_Consulta_Ope_Pedidos.MoveNext
    Wend
    Rs_Consulta_Ope_Pedidos.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

Private Sub Btn_Comparativo_Ventas_Click()
    If Cmb_Tipo_Reporte.Text = "COMPARATIVO" Then
        Reporte_Comparativo_Ventas
    Else
        Reporte_Detallado_Ventas
    End If
End Sub

Private Sub Btn_Efectividad_Cobranza_Click()
    Reporte_Efectividad
End Sub

Private Sub Btn_Exportar_Excel_Click()
Dim Columna As Integer     'Indica que columna del grid es la que se esta consultando
Dim Fila As Long           'Indica que fila del grid es la que se esta consultando
Dim Ruta_Archivo As String 'Obtiene la rura del archivo en donde se desea guardar
Dim CABECERA As String     'Indica el texto a exportar

On Error GoTo ErrHandler
    'Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    'Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    'Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.XLS|"
    'Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    'Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    'Display name of selected file
    Ruta_Archivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    Open Ruta_Archivo For Output As #1
        For Fila = 0 To Grid_Reporte.Rows - 1
            For Columna = 0 To Grid_Reporte.Cols - 1
                If Columna > 0 Then
                    CABECERA = CABECERA & Chr(9) & Grid_Reporte.TextMatrix(Fila, Columna)
                Else
                    CABECERA = Grid_Reporte.TextMatrix(Fila, Columna)
                End If
            Next Columna
            Print #1, CABECERA
            CABECERA = ""
        Next Fila
    Close #1
    MsgBox "Reporte Exportado a Excel", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Ha ocurrido un error en la exportación, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub

Private Sub Btn_Guardar_Disco_Click()
Dim RutaArchivo As String
Dim linea As String
On Error GoTo ErrHandler
    ' Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    ' Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Texto |*.TXT|"
    ' Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    ' Display name of selected file
    RutaArchivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    SHCopyFile "C:\REP\Reportes.txt", RutaArchivo
    MsgBox "Reporte Guardado en " & RutaArchivo, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Ocurrio un error al copiar el archivo, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub

Private Sub Btn_Imprimir_Click()
Dim linea As String 'Obtiene el texto a imprimir

    'Printer.PaperSize = vbPRPSLetter
    'Printer.PaperBin = 8
    'Printer.Orientation = vbPRORPortrait
    Printer.FontSize = 8
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 11
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 8
    Printer.Font = "Courier New"
    'Valida que exista la ruta si no la crea
    If Len(Dir("C:\REP", vbDirectory)) = 0 Then
        MkDir "C:\REP"
    End If
    Open "C:\REP\Reportes.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, linea
        Printer.Print linea
    Loop
    Printer.EndDoc
    Close #1
    MsgBox "Reporte enviado a impresora", vbInformation
End Sub

Private Sub Btn_Inventario_General_Click()
    Select Case Fra_Reporte_Inventarios.Caption
        Case "Inventario de Producto Terminado"
            Reporte_Inventario                  'Reporte de inventario por ultimo costo de producto terminado
        Case "Inventario Detallado de Producto Terminado"
            Inventario_Producto_Detallado
        Case "Inventario de Materia Prima"
            Reporte_Inventario_Materia_Prima    'Reporte de inventario por ultimo costo de materia prima
        Case "Inventario Detallado de Materia Prima"
            Inventario_Producto_Detallado_MP
    End Select
End Sub

Private Sub Btn_Kardex_Reporte_Click()
    If Fra_Kardex.Caption = "Kardex de Producto Terminado" Then
        Kardex_Producto_Terminado
    Else
        Kardex_Materia_Prima
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Kardex_Producto_Terminado
    'DESCRIPCIÓN: Muestra el kardex del producto
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 30-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Kardex_Producto_Terminado()
Dim Mi_SQL As String
Dim Rs_Consulta_Productos As rdoResultset           'Manejo de registro
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Salidas As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Total_Entrada As Double                         'Almacena el costo total de las entradas
Dim Total_Salida As Double                          'Almacena el costo total de las salidas
Dim Total_Productos As Double                       'Almacena la cantidad total de piezas
Dim Ubicacion_Almacen As String

    If Cmb_Producto.ListIndex > -1 Then
        'Kardex de productos
        Call Encabezado_Reportes("REPORTE DE KARDEX DE PRODUCTO TERMINADO", Format(Dtp_Kardex_Fecha_Inicial.Value, "dd/MMM/yyyy"), Format(Dtp_Kardex_Fecha_Final.Value, "dd/MMM/yyyy"), True)
        'Consulta los productos
        Mi_SQL = "SELECT * FROM Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(30); "Codigo      : "; Rs_Consulta_Productos.rdoColumns("Codigo")
            Grid_Reporte.AddItem "Codigo :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Codigo")
            Print #1, Spc(30); "Nombre      : "; Mid(Rs_Consulta_Productos.rdoColumns("Nombre"), 1, 70)
            Grid_Reporte.AddItem "Nombre :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Nombre")
            Print #1, Spc(30); "Minimo           : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Minimo"), 3)
            Grid_Reporte.AddItem "Minimo :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Minimo")
            Print #1, Spc(30); "Maximo           : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Maximo"), 3)
            Grid_Reporte.AddItem "Maximo" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Maximo")
            Print #1, Spc(30); "Reorden          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Punto_Reorden"), 3)
            Grid_Reporte.AddItem "Reorden" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Punto_Reorden")
            Grid_Reporte.AddItem ""
            Print #1,
        End If
        Print #1, "________________________________________________________________________________________________________________________"
        Print #1,
        Print #1, "   Fecha    Referencia  Movim.      Cantidad  No.Lote             Caducidad     Precio Venta   Costo Compra      Total  "
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Referencia" & Chr(9) & "Movimiento" & Chr(9) & "No. Lote" & Chr(9) & "Fecha_Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Precio Venta" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        Grid_Reporte.AddItem ""
        Print #1, "________________________________________________________________________________________________________________________"
        'Consulta los movimientos del producto en entradas
        Mi_SQL = "SELECT Ope_Entradas.No_Entrada,Ope_Entradas.Fecha,Ope_Entradas.Tipo_Entrada,Ope_Entradas.Estatus"
        Mi_SQL = Mi_SQL & " ,Detalles_Entradas.Producto_ID,Detalles_Entradas.Cantidad,Detalles_Entradas.Costo_Compra,Detalles_Entradas.Numero_Lote,Detalles_Entradas.Fecha_Caducidad,Detalles_Entradas.No_Partida,Detalles_Entradas.Almacen_ID"
        Mi_SQL = Mi_SQL & " FROM Ope_Entradas,Detalles_Entradas"
        Mi_SQL = Mi_SQL & " WHERE Ope_Entradas.No_Entrada=Detalles_Entradas.No_Entrada"
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Ope_Entradas.Fecha>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Ope_Entradas.Fecha<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND (Ope_Entradas.Estatus<>'CANCELADA')"
        Mi_SQL = Mi_SQL & " ORDER BY Ope_Entradas.Fecha"
        Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        Total_Productos = 0
        With Rs_Consulta_Entradas
            If Not .EOF Then
                Print #1,
                Print #1, Spc(60); "ENTRADAS"
                Grid_Reporte.AddItem "ENTRADAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las entradas y suma el costo de las entradas y el total de piezas y guarda la información en el grid para su exportación a Excel
                While Not .EOF
                    If Not IsNull(.rdoColumns("Almacen_ID")) Then
                        Mi_SQL = "SELECT Almacen_ID,Nombre FROM Cat_Almacenes"
                        Mi_SQL = Mi_SQL & " WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                        Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ubicacion.EOF Then
                            Ubicacion_Almacen = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                        Else
                            Ubicacion_Almacen = "N/A"
                        End If
                        Rs_Consulta_Ubicacion.Close
                    Else
                        Ubicacion_Almacen = "N/A"
                    End If
                    If Not IsNull(.rdoColumns("Numero_Lote")) Then
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(1); .rdoColumns("No_Entrada"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Entrada"), 1, 10); _
                            Spc(12 - Len(Mid(.rdoColumns("Tipo_Entrada"), 1, 10))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                            Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 20); _
                            Spc(22 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 20))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
                            Spc(2); Mid(Ubicacion_Almacen, 1, 17); _
                            Spc(19 - Len(Mid(Ubicacion_Almacen, 1, 17))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Entradas("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    Else
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(1); .rdoColumns("No_Entrada"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Entrada"), 1, 10); _
                            Spc(12 - Len(Mid(.rdoColumns("Tipo_Entrada"), 1, 10))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                            Spc(2); ""; _
                            Spc(22); ""; _
                            Spc(32); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Entradas("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    End If
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Entrada") _
                        & Chr(9) & .rdoColumns("Tipo_Entrada") _
                        & Chr(9) & .rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & "" _
                        & Chr(9) & .rdoColumns("Costo_Compra") _
                        & Chr(9) & .rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")
                    Total_Entrada = Total_Entrada + Val(.rdoColumns("Cantidad") * (.rdoColumns("Costo_Compra")))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Entradas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Entradas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(32); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Productos, "#######0"), 10); Spc(61); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Entrada, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL ENTRADAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Entrada
        'Consulta los movimientos del producto en salidas
        Mi_SQL = "SELECT Ope_Salidas.No_Salida,Ope_Salidas.Fecha,Ope_Salidas.Tipo,Ope_Salidas.Estatus,Detalles_Salidas.No_Salida,Detalles_Salidas.Producto_ID"
        Mi_SQL = Mi_SQL & " ,Detalles_Salidas.Cantidad,Detalles_Salidas.Precio_Venta,Detalles_Salidas.Costo_Compra,Detalles_Salidas.Numero_Lote,Detalles_Salidas.Fecha_Caducidad"
        Mi_SQL = Mi_SQL & " FROM Ope_Salidas,Detalles_Salidas,Cat_Productos"
        Mi_SQL = Mi_SQL & " WHERE Ope_Salidas.No_Salida=Detalles_Salidas.No_Salida"
        Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Producto_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.Fecha>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.Fecha<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND (Ope_Salidas.Estatus='ACTIVA' OR Ope_Salidas.Estatus IS NULL)"
        Mi_SQL = Mi_SQL & " ORDER BY Ope_Salidas.Fecha"
        Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Salida = 0
        Total_Productos = 0
        With Rs_Consulta_Salidas
            If Not .EOF Then
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(60); "SALIDAS"
                Grid_Reporte.AddItem "SALIDAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las salidas y suma el costo de las salidas y el total de piezas
                While Not .EOF
                    If Not IsNull(.rdoColumns("Numero_Lote")) Then
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(1); .rdoColumns("No_Salida"); _
                            Spc(1); Mid(.rdoColumns("Tipo"), 1, 10); _
                            Spc(12 - Len(Mid(.rdoColumns("Tipo"), 1, 10))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                            Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 20); _
                            Spc(22 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 20))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy"); _
                            Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Precio_Venta"), "#,##0.00"), 8); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    Else
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(1); .rdoColumns("No_Salida"); _
                            Spc(1); Mid(.rdoColumns("Tipo"), 1, 10); _
                            Spc(12 - Len(Mid(.rdoColumns("Tipo"), 1, 10))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0"), 7); _
                            Spc(2); ""; _
                            Spc(22); ""; _
                            Spc(17); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Precio_Venta"), "#,##0.00"), 8); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    End If
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Salida") _
                        & Chr(9) & .rdoColumns("Tipo") _
                        & Chr(9) & .rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & .rdoColumns("Precio_Venta") _
                        & Chr(9) & .rdoColumns("Costo_Compra") _
                        & Chr(9) & .rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")
                    Total_Salida = Total_Salida + Val(.rdoColumns("Cantidad") * (.rdoColumns("Costo_Compra")))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Salidas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Salidas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(32); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Productos, "#######0"), 10); Spc(61); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Salida, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL SALIDAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Salida
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime la existencia de los productos y saca la diferencia de costos de las entradas y salidas
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(36); "Existencia          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Existencia"), 10); Spc(25); "Total    :"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Entrada - Total_Salida, "#,###,###,##0.00"), 16)
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Existencia : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Existencia") & Chr(9) & Chr(9) & Chr(9) & Total_Entrada - Total_Salida
            Print #1, Spc(36); "Cuarentena          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Cuarentena"), 10)
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Cuarentena : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Cuarentena")
            'Print #1, Spc(36); "Reacondicionamiento : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Reacondicionamiento"), 7)
            'Grid_Reporte.AddItem Chr(9) & Chr(9) & "Reacondicionamiento : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Reacondicionamiento")
            Print #1, Spc(36); "Rechazo             : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Rechazo"), 10)
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Rechazo : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Rechazo")
        End If
        Call Encabezado_Reportes("", "", "", False)
        Rs_Consulta_Productos.Close
    Else
        MsgBox "Seleccione el producto a mostrar", vbExclamation
        Dtp_Kardex_Fecha_Inicial.Value = Now
        Dtp_Kardex_Fecha_Final.Value = Now
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Kardex_Materia_Prima
    'DESCRIPCIÓN: Muestra el kardex de materia prima
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 30-Septiembre-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Kardex_Materia_Prima()
Dim Mi_SQL As String
Dim Rs_Consulta_Productos As rdoResultset           'Manejo de registro
Dim Rs_Consulta_Entradas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Salidas As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Unidad As rdoResultset
Dim Rs_Consulta_Ubicacion As rdoResultset
Dim Total_Entrada As Double                         'Almacena el costo total de las entradas
Dim Total_Salida As Double                          'Almacena el costo total de las salidas
Dim Total_Productos As Double                       'Almacena la cantidad total de piezas
Dim Ubicacion_Almacen As String

    If Cmb_Producto.ListIndex > -1 Then
        'Kardex de productos
        Call Encabezado_Reportes("REPORTE DE KARDEX DE MATERIA PRIMA", Format(Dtp_Kardex_Fecha_Inicial.Value, "dd/MMM/yyyy"), Format(Dtp_Kardex_Fecha_Final.Value, "dd/MMM/yyyy"), True)
        'Consulta los productos
        Mi_SQL = "SELECT * FROM Cat_Materias_Primas"
        Mi_SQL = Mi_SQL & " WHERE Materia_Prima_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(30); "Codigo      : "; Rs_Consulta_Productos.rdoColumns("Codigo")
            Grid_Reporte.AddItem "Codigo :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Codigo")
            Print #1, Spc(30); "Nombre      : "; Mid(Rs_Consulta_Productos.rdoColumns("Nombre"), 1, 70)
            Grid_Reporte.AddItem "Nombre :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Nombre")
            Print #1, Spc(30); "Minimo           : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Minimo"), 3)
            Grid_Reporte.AddItem "Minimo :" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Minimo")
            Print #1, Spc(30); "Maximo           : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Maximo"), 3)
            Grid_Reporte.AddItem "Maximo" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Maximo")
            Print #1, Spc(30); "Reorden          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Punto_Reorden"), 3)
            Grid_Reporte.AddItem "Reorden" & Chr(9) & Rs_Consulta_Productos.rdoColumns("Punto_Reorden")
            Grid_Reporte.AddItem ""
            If Not IsNull(Rs_Consulta_Productos.rdoColumns("Unidad_ID")) Then
                Mi_SQL = "SELECT Unidad_ID,Nombre FROM Cat_Unidades"
                Mi_SQL = Mi_SQL & " WHERE Unidad_ID='" & Rs_Consulta_Productos.rdoColumns("Unidad_ID") & "'"
                Set Rs_Consulta_Unidad = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Unidad.EOF Then
                    Print #1, Spc(30); "Unidad           : "; Rs_Consulta_Unidad.rdoColumns("Nombre")
                    Grid_Reporte.AddItem "Unidad" & Chr(9) & Rs_Consulta_Unidad.rdoColumns("Nombre")
                    Grid_Reporte.AddItem ""
                End If
            End If
            Print #1,
        End If
        Print #1, "________________________________________________________________________________________________________________________"
        Print #1,
        Print #1, " Fecha    Referencia  Movimiento   Estatus   Cantidad  No.Lote      Caducidad      Ubicacion     Costo Compra    Total  "
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Referencia" & Chr(9) & "Movimiento" & Chr(9) & "No. Lote" & Chr(9) & "Fecha Caducidad" & Chr(9) & "Cantidad" & Chr(9) & "Ubiacion" & Chr(9) & "Costo Compra" & Chr(9) & "Total"
        Grid_Reporte.AddItem ""
        Print #1, "________________________________________________________________________________________________________________________"
        'Consulta los movimientos del producto en entradas
        Mi_SQL = "SELECT Alm_Entradas_Materia_Prima.No_Entrada,Alm_Entradas_Materia_Prima.Fecha,Alm_Entradas_Materia_Prima.Tipo_Entrada,Alm_Entradas_Materia_Prima_Detalles.No_Entrada,Alm_Entradas_Materia_Prima_Detalles.Estatus"
        Mi_SQL = Mi_SQL & " ,Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID,Alm_Entradas_Materia_Prima_Detalles.Cantidad,Costo_Compra,Alm_Entradas_Materia_Prima_Detalles.Numero_Lote,Alm_Entradas_Materia_Prima_Detalles.Fecha_Caducidad,Alm_Entradas_Materia_Prima_Detalles.Almacen_ID"
        Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima,Alm_Entradas_Materia_Prima_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima.No_Entrada=Alm_Entradas_Materia_Prima_Detalles.No_Entrada"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Fecha>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Fecha<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Estatus<>'CANCELADA'"
        Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas_Materia_Prima.Fecha,Alm_Entradas_Materia_Prima.No_Entrada"
        Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Entrada = 0
        Total_Productos = 0
        With Rs_Consulta_Entradas
            If Not .EOF Then
                Print #1,
                Print #1, Spc(60); "ENTRADAS"
                Grid_Reporte.AddItem "ENTRADAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las entradas y suma el costo de las entradas y el total de piezas y guarda la información en el grid para su exportación a Excel
                While Not .EOF
                    'Consulta la ubicacion
                    If Not IsNull(.rdoColumns("Almacen_ID")) Then
                        Mi_SQL = "SELECT Almacen_ID,Nombre FROM Cat_Almacenes"
                        Mi_SQL = Mi_SQL & " WHERE Almacen_ID='" & .rdoColumns("Almacen_ID") & "'"
                        Set Rs_Consulta_Ubicacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ubicacion.EOF Then
                            Ubicacion_Almacen = Rs_Consulta_Ubicacion.rdoColumns("Nombre")
                        Else
                            Ubicacion_Almacen = "N/A"
                        End If
                        Rs_Consulta_Ubicacion.Close
                    Else
                        Ubicacion_Almacen = "N/A"
                    End If
                    Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(1); .rdoColumns("No_Entrada"); _
                        Spc(1); Mid(.rdoColumns("Tipo_Entrada"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Tipo_Entrada"), 1, 15))); Mid(.rdoColumns("Estatus"), 1, 3); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#0.00"), 10); _
                        Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 15); _
                        Spc(16 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 15))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yy"); _
                        Spc(2); Mid(Ubicacion_Almacen, 1, 15); _
                        Spc(17 - Len(Mid(Ubicacion_Almacen, 1, 15))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Entradas("Costo_Compra"), "#,##0.00"), 8); _
                        Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Entrada") _
                        & Chr(9) & .rdoColumns("Tipo_Entrada") _
                        & Chr(9) & .rdoColumns("Estatus") _
                        & Chr(9) & .rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & Ubicacion_Almacen _
                        & Chr(9) & .rdoColumns("Costo_Compra") _
                        & Chr(9) & .rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")
                    Total_Entrada = Total_Entrada + Val(.rdoColumns("Cantidad") * (.rdoColumns("Costo_Compra")))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Entradas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Entradas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(29); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Productos, "#######0.00"), 13); Spc(61); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Entrada, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL ENTRADAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Entrada
        'Consulta los movimientos del producto en salidas
        Mi_SQL = "SELECT DISTINCT Alm_Salidas_Materia_Prima.No_Salida,Alm_Salidas_Materia_Prima.Fecha,Alm_Salidas_Materia_Prima.Tipo_Salida,Alm_Salidas_Materia_Prima.Estatus,Alm_Salidas_Materia_Prima_Detalles.No_Salida,Alm_Salidas_Materia_Prima_Detalles.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " ,Alm_Salidas_Materia_Prima_Detalles.Cantidad,Alm_Salidas_Materia_Prima_Detalles.Costo_Compra,Alm_Salidas_Materia_Prima_Detalles.Numero_Lote,Alm_Salidas_Materia_Prima_Detalles.Fecha_Caducidad"
        Mi_SQL = Mi_SQL & " FROM Alm_Salidas_Materia_Prima,Alm_Salidas_Materia_Prima_Detalles,Cat_Materias_Primas"
        Mi_SQL = Mi_SQL & " WHERE Alm_Salidas_Materia_Prima.No_Salida=Alm_Salidas_Materia_Prima_Detalles.No_Salida"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Materia_Prima_Detalles.Materia_Prima_ID='" & Format(Cmb_Producto.ItemData(Cmb_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Materia_Prima.Fecha>=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Inicial.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Alm_Salidas_Materia_Prima.Fecha<=" & Par_Fecha & Format(Dtp_Kardex_Fecha_Final.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " ORDER BY Alm_Salidas_Materia_Prima.Fecha,Alm_Salidas_Materia_Prima.No_Salida"
        Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Salida = 0
        Total_Productos = 0
        With Rs_Consulta_Salidas
            If Not .EOF Then
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(60); "SALIDAS"
                Grid_Reporte.AddItem "SALIDAS"
                Grid_Reporte.AddItem ""
                Print #1,
                'Imprime las salidas y suma el costo de las salidas y el total de piezas
                While Not .EOF
                    If Not IsNull(.rdoColumns("Numero_Lote")) Then
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); .rdoColumns("No_Salida"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Salida"), 1, 20); _
                            Spc(22 - Len(Mid(.rdoColumns("Tipo_Salida"), 1, 20))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#0.00"), 10); _
                            Spc(2); Mid(.rdoColumns("Numero_Lote"), 1, 15); _
                            Spc(16 - Len(Mid(.rdoColumns("Numero_Lote"), 1, 15))); Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yy"); _
                            Spc(12); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    Else
                        Print #1, Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); .rdoColumns("No_Salida"); _
                            Spc(1); Mid(.rdoColumns("Tipo_Salida"), 1, 20); _
                            Spc(22 - Len(Mid(.rdoColumns("Tipo"), 1, 20))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "######0.00"), 10); _
                            Spc(2); ""; _
                            Spc(22); ""; _
                            Spc(23); _
                            Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Costo_Compra"), "#,##0.00"), 8); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")), "#,###,##0.00"), 12)
                    End If
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Salida") _
                        & Chr(9) & .rdoColumns("Tipo_Salida") _
                        & Chr(9) & .rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("Cantidad") _
                        & Chr(9) & "" _
                        & Chr(9) & .rdoColumns("Costo_Compra") _
                        & Chr(9) & .rdoColumns("Cantidad") * .rdoColumns("Costo_Compra")
                    Total_Salida = Total_Salida + Val(.rdoColumns("Cantidad") * (.rdoColumns("Costo_Compra")))
                    Total_Productos = Total_Productos + Val(.rdoColumns("Cantidad"))
                    Rs_Consulta_Salidas.MoveNext
                Wend
            End If
        End With
        Rs_Consulta_Salidas.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(29); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Productos, "#######0.00"), 13); Spc(61); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Salida, "#,###,###,##0.00"), 16)
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "TOTAL SALIDAS" & Chr(9) & Total_Productos & Chr(9) & Chr(9) & Chr(9) & Total_Salida
        Grid_Reporte.AddItem ""
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        'Imprime la existencia de los productos y saca la diferencia de costos de las entradas y salidas
        If Not Rs_Consulta_Productos.EOF Then
            Print #1, Spc(34); "Existencia          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Productos.rdoColumns("Existencia"), "#0.00"), 13); Spc(25); "Total    :"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Entrada - Total_Salida, "#,###,###,##0.00"), 16)
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Existencia : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Existencia") & Chr(9) & Chr(9) & Chr(9) & Total_Entrada - Total_Salida
            'Print #1, Spc(36); "Cuarentena          : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Cuarentena"), 10)
            'Grid_Reporte.AddItem Chr(9) & Chr(9) & "Cuarentena : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Cuarentena")
            'Print #1, Spc(36); "Reacondicionamiento : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Reacondicionamiento"), 7)
            'Grid_Reporte.AddItem Chr(9) & Chr(9) & "Reacondicionamiento : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Reacondicionamiento")
            'Print #1, Spc(36); "Rechazo             : "; Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Productos.rdoColumns("Rechazo"), 10)
            'Grid_Reporte.AddItem Chr(9) & Chr(9) & "Rechazo : " & Chr(9) & Rs_Consulta_Productos.rdoColumns("Rechazo")
        End If
        Call Encabezado_Reportes("", "", "", False)
        Rs_Consulta_Productos.Close
    Else
        MsgBox "Seleccione el producto a mostrar", vbExclamation
        Dtp_Kardex_Fecha_Inicial.Value = Now
        Dtp_Kardex_Fecha_Final.Value = Now
    End If
End Sub

Public Sub Encabezado_Reportes(NombreReporte As String, FechaInicio As String, FechaFinal As String, Abrir As Boolean)
    If Abrir = True Then
        Grid_Reporte.Rows = 0
        'Valida que exista la ruta si no la crea
        If Len(Dir("C:\REP", vbDirectory)) = 0 Then
            MkDir "C:\REP"
        End If
        Open "C:\REP\Reportes.txt" For Output As #1
        Print #1, "************************************************************************************************************************"
        Print #1,
        Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
        Print #1,
        Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
        Grid_Reporte.AddItem ""
        Print #1, Conectar_Ayudante.Centra(NombreReporte, Len("************************************************************************************************************************"))
        Grid_Reporte.AddItem NombreReporte
        Grid_Reporte.AddItem ""
        Print #1,
        If FechaInicio <> "" And FechaFinal <> "" Then
            Print #1, Conectar_Ayudante.Centra("DEL " & FechaInicio & " AL " & FechaFinal, Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & FechaInicio & " AL " & FechaFinal
        ElseIf FechaInicio <> "" Then
            Print #1, Conectar_Ayudante.Centra("FECHA GENERACION " & FechaInicio, Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "FECHA GENERACION " & FechaInicio
            Grid_Reporte.AddItem ""
        Else
            Print #1,
        End If
        Print #1, "************************************************************************************************************************"
        Print #1,
    Else
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Close #1
        Me.Edm_Reporte.Font.Size = 8
        Me.Edm_Reporte.Font = "Courier New"
        Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
        Me.Edm_Reporte.Visible = True
        Ancho_Frame = Me.Width
        Largo_Frame = Me.Height
        Me.Width = 13250
        Me.Height = 8200
        Me.Left = 0
        Me.Top = 0
    End If
End Sub

Private Sub Btn_Reporte_Anticipos_Clientes_Click()
    Select Case Fra_Anticipos_Clientes.Caption
        Case "Anticipos Clientes"
            Reporte_Anticipos_Clientes      'Reporte de anticipos clientes
        Case "Anticipos Proveedores"
            Reporte_Anticipos_Proveedores   'Reporte de anticipos proveedores
        Case "Reporte de Pago de Saldos Pequeños"
            If Cmb_Cliente_Anticipo.ListIndex > -1 Then
                Reporte_Pago_Saldos_Minimos 'Reporte de pagos mínimos (cuenta de ajustes)
            Else
                MsgBox "Seleccione una cuenta de la lista", vbExclamation
                Cmb_Cliente_Anticipo.SetFocus
            End If
    End Select
End Sub

Private Sub Reporte_Cliente_Estadistico()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Fecha_Factura As rdoResultset
Dim Rs_Adm_Facturas_Clientes As rdoResultset
Dim Rs_Adm_Notas_Credtio_Clientes As rdoResultset
Dim Rs_Consulta_Clientes As rdoResultset
Dim Rs_Consulta_Ciudades As rdoResultset
Dim Ciudad As String
Dim Total_Nota_Credito As Double
Dim Total_Pagos As Double
Dim Compra_Total As Double
Dim Compra_Maxima As Double
Dim Compra_Minima As Double
Dim Cheque_Devuelto As Double
Dim Meses_Trasncurridos As Integer
Dim Puntos_Gestion As Integer
Dim Nivel_Gestion As String

    If Cmb_Cliente_Estadistico.ListIndex > -1 Then
        Grid_Reporte.Cols = 7
        Grid_Reporte.Rows = 0
        Call Encabezado_Reportes("ESTADISTICO DE CLIENTES", Format(Now, "dd/MMM/yyyy"), "", True)
        Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Ciudades"
        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Nombre='" & Cmb_Cliente_Estadistico.Text & "'"
        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Clientes
            If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Ciudades.EOF Then
                    Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                Else
                    Ciudad = ""
                End If
                Rs_Consulta_Ciudades.Close
            Else
                Ciudad = ""
            End If
            Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Estadistico.Text
            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
            Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
            Print #1, Spc(2); "Ciudad:    "; Ciudad
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Estadistico.Text
            Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
            Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
            Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
        End With
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, "     Estadístico         Limite            Notas          Compra           Compra           Cargos          Nivel  "
        Print #1, "     Pago X Mes          Crédito          Crédito         Máxima           Mínima          Bancarios       Gestión "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Estadístico Pago X Mes" & Chr(9) & "Límte de Crédito" & Chr(9) & "Notas de Crédito" & Chr(9) & "Compra Máxima" & Chr(9) & "Compra Mínima" & Chr(9) & "Cargos Bancarios" & Chr(9) & "Nivel de Gestión"
        Grid_Reporte.AddItem ""
        'Consulta los clientes
        Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes "
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & Format(Cmb_Cliente_Estadistico.ItemData(Cmb_Cliente_Estadistico.ListIndex), "00000") & "' ORDER BY Nombre"
        Set Rs_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Cat_Clientes.EOF
            Puntos_Gestion = 0
            'Consulta la fecha de la primera compra del cliente y calcula el total de meses transcurridos
            Mi_SQL = "SELECT TOP 1 Fecha FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Fecha_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Fecha_Factura.EOF Then
                Meses_Trasncurridos = DateDiff("M", Format(Rs_Consulta_Fecha_Factura.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
            Else
                Meses_Trasncurridos = 0
            End If
            Rs_Consulta_Fecha_Factura.Close
            'Consulta el total pagado, el máximo y el mínimo de compra del cliente
            Mi_SQL = "SELECT SUM(Abono) AS Total_Pagos,MAX(Total) AS Maxima_Compra,MIN(Total) AS Minima_Compra,SUM(Total) AS Total_Compras"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")) Then
                Total_Pagos = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")
                Compra_Maxima = Rs_Adm_Facturas_Clientes.rdoColumns("Maxima_Compra")
                Compra_Minima = Rs_Adm_Facturas_Clientes.rdoColumns("Minima_Compra")
                Compra_Total = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Compras")
                'Valida que el límite supere la cantidad adeudada para sumar un punto a nivel de gestion
                If Meses_Trasncurridos > 0 Then
                    If Val(Compra_Total - Total_Pagos) > (Total_Pagos / Meses_Trasncurridos * 2) Then
                        Puntos_Gestion = Puntos_Gestion + 1
                    End If
                End If
            Else
                Total_Pagos = 0
                Compra_Maxima = 0
                Compra_Minima = 0
            End If
            Rs_Adm_Facturas_Clientes.Close
            'Consulta las notas de crédito
            Mi_SQL = "SELECT SUM(Importe) FROM Adm_Notas_Credito_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Autorizado='SI' AND Aplicada='SI'"
            Set Rs_Adm_Notas_Credtio_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)) Then
                Total_Nota_Credito = Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)
                'Valida el 20 % de compras superen el total de notas para agregar un punto a nivel de gestion
                If Val(Total_Nota_Credito / Compra_Total * 100) > 20 Then
                    Puntos_Gestion = Puntos_Gestion + 1
                End If
            Else
                Total_Nota_Credito = 0
            End If
            Rs_Adm_Notas_Credtio_Clientes.Close
            'Consulta las comisiones por cheques devueltos
            Mi_SQL = "SELECT SUM(Adm_Detalles_Facturas.Importe) AS Importe,SUM(Adm_Detalles_Facturas.IVA) AS Iva"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Producto='COMISION CHEQUE DEVUELTO'"
            Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Importe")) Then
                Cheque_Devuelto = Rs_Adm_Facturas_Clientes.rdoColumns("Importe") + Rs_Adm_Facturas_Clientes.rdoColumns("Iva")
                Puntos_Gestion = Puntos_Gestion + 1
            Else
                Cheque_Devuelto = 0
            End If
            Rs_Adm_Facturas_Clientes.Close
            'Valida el nivel de gestión
            Select Case Puntos_Gestion
                Case 0
                    Nivel_Gestion = "BAJO"
                Case 1
                    Nivel_Gestion = "BAJO"
                Case 2
                    Nivel_Gestion = "MEDIO"
                Case 3
                    Nivel_Gestion = "ALTO"
            End Select
            If Meses_Trasncurridos > 0 Then
                Print #1, Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12); _
                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12); _
                    Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
                    Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "#,###,##0.00"), 12); _
                    Spc(10); Nivel_Gestion
                Grid_Reporte.AddItem Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "###,##0.00"), 9) _
                    & Chr(9) & Nivel_Gestion
            Else
                Print #1, Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                    Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
                    Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
                    Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
                    Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                    Spc(10); Nivel_Gestion
                Grid_Reporte.AddItem Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
                    & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "###,##0.00"), 9) _
                    & Chr(9) & Nivel_Gestion
            End If
            Rs_Cat_Clientes.MoveNext
        Wend
        Rs_Cat_Clientes.Close
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "Seleccione un cliente de la lista", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub Btn_Reporte_Cliente_Estadistico_Click()
    If Fra_Estadistico_Cliente.Caption = "Estadistico Cliente" Then
        Reporte_Cliente_Estadistico
    Else    'Reporte general de clientes
        Reporte_Clientes
    End If
End Sub

Private Sub Btn_Reporte_Cobranza_Click()
    Select Case Fra_Consulta_Cobranza.Caption
        Case "Reporte de Cobranza"
            If Cmb_Documento_Cobranza.ListIndex = 0 Then
                Reporte_Cobranza
            Else
                Reporte_Cobranza_Remisiones
            End If
        Case "Estado de Cuenta de Cobranza"
            Estado_Cuenta_Cobranza
        Case "Estado de Cuenta Detallado"
            If Cmb_Empresa_Cobranza.ListIndex > -1 Then
                Estado_Cuenta_Detallado
            Else
                MsgBox "Seleccione el cliente de la lista", vbExclamation
            End If
    End Select
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Cobranza
    'DESCRIPCIÓN: Muestra la cobranza aplicada
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 09-Abril-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Cobranza()
Dim Cliente As String                                   'Almacena el nombre del cliente
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Cuenta As String                                    'Almacena el numero de cuenta
Dim Estatus As String
Dim Total_Abonos As Double                              'Almacena el total de los abonos en las fechas especificadas
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Movimientos As rdoResultset         'Manejo de registro
Dim Aplica_IVA As String

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT Adm_Movimientos.No_Movimiento,Adm_Movimientos.Banco_ID,Adm_Movimientos.Cliente_ID,Adm_Movimientos.No_Cheque,Adm_Movimientos.No_Factura,Adm_Movimientos.Empresa,Adm_Movimientos.Fecha,Adm_Movimientos.Cuenta,Adm_Movimientos.Cantidad,Adm_Movimientos.Estatus,Adm_Movimientos.Forma_Pago,Adm_Facturas_Clientes.Saldo AS Restante,Adm_Facturas_Clientes.IVA"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos,Adm_Facturas_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Factura_Remision='F'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo='I'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.No_Factura=Adm_Facturas_Clientes.No_Factura"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Empresa=Adm_Facturas_Clientes.Empresa"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo>='" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & " 00:00:00" & "'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo<='" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & " 23:59:59'"
    If Cmb_Empresa_Cobranza.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Empresa=" & Cmb_Empresa_Cobranza.ListIndex + 1
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Movimientos.Fecha"
    Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Movimientos.EOF Then
        With Rs_Consulta_Adm_Movimientos
            Call Encabezado_Reportes("REPORTE DE COBRANZA DE FACTURACION", Format(Dtp_Fecha_Inicio_Cobranza.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Cobranza.Value, "dd/MMM/yyyy"), True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   No. Factura     Cliente                             Fecha           Abono        Restante     Cuenta        Estatus   "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Restante" & Chr(9) & "Cuenta Banco" & Chr(9) & "Estatus" & Chr(9) & "Aplica IVA"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    Else
                        Cliente = ""
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                Else
                    Cliente = ""
                End If
                'Consulta el nombre del banco
                If .rdoColumns("Estatus") = "A" Then
                    Estatus = "Activo"
                Else
                    Estatus = "Cancel"
                End If
                If Not IsNull(.rdoColumns("Empresa")) Then
                    If .rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                Else
                    Empresa = "   "
                End If
                If Not IsNull(.rdoColumns("Cuenta")) Then
                    Cuenta = .rdoColumns("Cuenta")
                Else
                    If .rdoColumns("Forma_Pago") = "NOTAS CREDITO" Then
                        Cuenta = "NOTA CREDITO"
                    Else
                        Cuenta = ""
                    End If
                End If
                If .rdoColumns("IVA") > 0 Then
                    Aplica_IVA = "S"
                Else
                    Aplica_IVA = "N"
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); _
                    Spc(1); Empresa; _
                    Spc(2); Mid(Cliente, 1, 35); _
                    Spc(37 - Len(Mid(Cliente, 1, 35))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Restante"), "#,###,##0.00"), 13); _
                    Spc(2); Mid(Cuenta, 1, 14); _
                    Spc(15 - Len(Mid(Cuenta, 1, 14))); Estatus; _
                    Spc(2); Aplica_IVA
                Grid_Reporte.AddItem .rdoColumns("No_Factura") _
                    & Chr(9) & Empresa _
                    & Chr(9) & Cliente _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Restante"), "#,###,##0.00") _
                    & Chr(9) & .rdoColumns("Cuenta") _
                    & Chr(9) & Estatus _
                    & Chr(9) & Aplica_IVA
                Total_Abonos = Total_Abonos + .rdoColumns("Cantidad")
                .MoveNext
            Wend
            End With
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(55); "Total:     "; Format(Total_Abonos, "#,###,##0.00"); Spc(27); "Aplica IVA S/N"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_Abonos, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay movimientos en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Adm_Movimientos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Cobranza_Remisiones
    'DESCRIPCIÓN: Muestra la cobranza aplicada para remisiones
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 19-Mayo-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Cobranza_Remisiones()
Dim Cliente As String                                   'Almacena el nombre del cliente
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Cuenta As String                                    'Almacena el numero de cuenta
Dim Estatus As String
Dim Total_Abonos As Double                              'Almacena el total de los abonos en las fechas especificadas
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Movimientos As rdoResultset         'Manejo de registro

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT Adm_Movimientos.No_Movimiento,Adm_Movimientos.Banco_ID,Adm_Movimientos.Cliente_ID,Adm_Movimientos.No_Cheque,Adm_Movimientos.No_Factura,Adm_Movimientos.Empresa,Adm_Movimientos.Fecha,Adm_Movimientos.Cuenta,Adm_Movimientos.Cantidad,Adm_Movimientos.Estatus,Adm_Movimientos.Forma_Pago,Adm_Remisiones_Clientes.Saldo AS Restante"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos,Adm_Remisiones_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Factura_Remision='R'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo='I'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.No_Factura=Adm_Remisiones_Clientes.No_Remision"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Empresa=Adm_Remisiones_Clientes.Empresa"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo>='" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & " 00:00:00" & "'"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha_Creo<='" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & " 23:59:59'"
    If Cmb_Empresa_Cobranza.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Empresa=" & Cmb_Empresa_Cobranza.ListIndex + 1
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Movimientos.Fecha"
    Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Movimientos.EOF Then
        With Rs_Consulta_Adm_Movimientos
            Call Encabezado_Reportes("REPORTE DE COBRANZA DE NOTAS DE REMISION", Format(Dtp_Fecha_Inicio_Cobranza.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Cobranza.Value, "dd/MMM/yyyy"), True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Remision     Cliente                             Fecha           Abono        Restante     Cuenta        Estatus   "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Restante" & Chr(9) & "Cuenta Banco" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    Else
                        Cliente = ""
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                Else
                    Cliente = ""
                End If
                'Consulta el nombre del banco
                If .rdoColumns("Estatus") = "A" Then
                    Estatus = "Activo"
                Else
                    Estatus = "Cancelado"
                End If
                If Not IsNull(.rdoColumns("Empresa")) Then
                    If .rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                Else
                    Empresa = "   "
                End If
                If Not IsNull(.rdoColumns("Cuenta")) Then
                    Cuenta = .rdoColumns("Cuenta")
                Else
                    If .rdoColumns("Forma_Pago") = "NOTAS CREDITO" Then
                        Cuenta = "NOTA CREDITO"
                    Else
                        Cuenta = ""
                    End If
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(1); Empresa; Spc(2); Mid(Cliente, 1, 35); _
                    Spc(37 - Len(Mid(Cliente, 1, 35))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Restante"), "#,###,##0.00"), 13); _
                    Spc(2); Mid(Cuenta, 1, 14); Spc(15 - Len(Mid(Cuenta, 1, 14))); Estatus
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Cliente & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Cantidad"), "#,###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Restante"), "#,###,##0.00") & Chr(9) & .rdoColumns("Cuenta") & Chr(9) & Estatus
                Total_Abonos = Total_Abonos + .rdoColumns("Cantidad")
                .MoveNext
            Wend
            End With
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(55); "Total:     "; Format(Total_Abonos, "#,###,##0.00")
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_Abonos, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay movimientos en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Adm_Movimientos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Estado_Cuenta_Cobranza()
    'DESCRIPCIÓN: Muestra el estado de cuenta de la cobranza aplicada
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 09-Abril-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Estado_Cuenta_Cobranza()
Dim Cliente As String                                   'Almacena el nombre del cliente
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Estatus As String
Dim Total_Abonos As Double                              'Almacena el total de los abonos en las fechas especificadas
Dim Gran_Total As Double
Dim Rs_Consulta_Cat_Bancos As rdoResultset
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Movimientos As rdoResultset         'Manejo de registro

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    Call Encabezado_Reportes("ESTADO DE CUENTA DE COBRANZA", Format(Dtp_Fecha_Inicio_Cobranza.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Cobranza.Value, "dd/MMM/yyyy"), True)
    'Consulta los bancos
    Mi_SQL = "SELECT Banco_ID,Nombre,No_Cuenta FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Cat_Bancos.EOF
        Print #1,
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, "BANCO : "; Rs_Consulta_Cat_Bancos.rdoColumns("Nombre"); Spc(35 - Len("BANCO : " & Rs_Consulta_Cat_Bancos.rdoColumns("Nombre"))); "CUENTA: "; Rs_Consulta_Cat_Bancos.rdoColumns("No_Cuenta")
        Print #1,
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "BANCO:" & Chr(9) & Rs_Consulta_Cat_Bancos.rdoColumns("Nombre") & Chr(9) & "" & Chr(9) & "CUENTA" & Chr(9) & Rs_Consulta_Cat_Bancos.rdoColumns("No_Cuenta")
        Grid_Reporte.AddItem ""
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
        Mi_SQL = "SELECT No_Movimiento,Banco_ID,Cliente_ID,No_Cheque,No_Factura,Empresa,Fecha,Cuenta,Cantidad,Estatus"
        Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
        Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.Banco_ID='" & Rs_Consulta_Cat_Bancos.rdoColumns("Banco_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo='I'"
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Movimientos.Cuenta,Adm_Movimientos.Fecha"
        Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Movimientos.EOF Then
            With Rs_Consulta_Adm_Movimientos
                Print #1, "   No. Factura     Cliente                            Fecha Pago            Cantidad   Estatus                           "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Cliente" & Chr(9) & "Fecha Pago" & Chr(9) & "Cantidad" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
                Total_Abonos = 0
                While Not .EOF
                    If Not IsNull(.rdoColumns("Cliente_ID")) Then
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Clientes.EOF Then
                            Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                        Else
                            Cliente = ""
                        End If
                        Rs_Consulta_Cat_Clientes.Close
                    Else
                        Cliente = ""
                    End If
                    If .rdoColumns("Estatus") = "A" Then
                        Estatus = "Activo"
                    Else
                        Estatus = "Cancelado"
                    End If
                    If Not IsNull(.rdoColumns("Empresa")) Then
                        If .rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                    Else
                        Empresa = "   "
                    End If
                    If Not IsNull(.rdoColumns("No_Factura")) Then
                        Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(1); Empresa; Spc(2); Mid(Cliente, 1, 35); _
                            Spc(37 - Len(Mid(Cliente, 1, 35))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0.00") & Chr(9) & Estatus
                    Else
                        Print #1, Spc(17); Mid(Cliente, 1, 35); Spc(37 - Len(Mid(Cliente, 1, 35))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); Spc(3); Estatus
                        Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0.00") & Chr(9) & Estatus
                    End If
                    Total_Abonos = Total_Abonos + .rdoColumns("Cantidad")
                    Gran_Total = Gran_Total + .rdoColumns("Cantidad")
                    .MoveNext
                Wend
                End With
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(65); "TOTAL:     "; Format(Total_Abonos, "#,###,##0.00")
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTAL:" & Chr(9) & Format(Total_Abonos, "#,###,##0.00")
        Else
            Print #1, "   SIN MOVIMIENTOS"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem "SIN MOVIMIENTOS"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem ""
        End If
        Rs_Consulta_Cat_Bancos.MoveNext
    Wend
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(65); "GRAN TOTAL: "; Format(Gran_Total, "#,###,##0.00")
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL:" & Chr(9) & Format(Gran_Total, "#,###,##0.00")
    Call Encabezado_Reportes("", "", "", False)
    Rs_Consulta_Cat_Bancos.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Estado_Cuenta_Detallado
'DESCRIPCION: Muestra el estado de cuenta de la cobranza aplicada por cliente
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 09-Abril-2010
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICACION:
'*******************************************************************************
Private Sub Estado_Cuenta_Detallado()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Facturas As rdoResultset
Dim Rs_Consulta_Adm_Movimientos As rdoResultset
Dim No_Factura As String
Dim Empresa As String
Dim Estatus As String
Dim Total_Saldo As Double
Dim Saldo_Factura As Double

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    Call Encabezado_Reportes("HISTORIAL DE MOVIMIENTOS EN ESTADO DE CUENTA", Format(Dtp_Fecha_Inicio_Cobranza.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Cobranza.Value, "dd/MMM/yyyy"), True)
    'Consulta el nombre del cliente
    Mi_SQL = "SELECT Cliente_ID,Cat_Clientes.Nombre AS Cliente,RFC,Direccion,Colonia,CP,Telefono,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre AS Estado"
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes LEFT JOIN Cat_Ciudades ON Cat_Clientes.Ciudad_ID=Cat_Ciudades.Ciudad_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Estados ON Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Empresa_Cobranza.ItemData(Cmb_Empresa_Cobranza.ListIndex), "00000") & "'"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        Print #1,
        Print #1, "_________________________________________________________________________________________________________________________"
        Print #1, "CLIENTE  : "; Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"); Spc(72 - Len("CLIENTE  : " & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"))); "R.F.C.   : "; Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
        Print #1, "DOMICILIO: "; Rs_Consulta_Cat_Clientes.rdoColumns("Direccion"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
        Print #1, "TELEFONO : "; Rs_Consulta_Cat_Clientes.rdoColumns("Telefono"); "   C.P.: "; Rs_Consulta_Cat_Clientes.rdoColumns("CP")
        If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad")) Then
            Print #1, "CIUDAD   : "; Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Estado")
        Else
            Print #1, "CIUDAD   : "
        End If
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, "   Fecha     No. Factura     Concepto                 Cargo          Abono       Saldo                                   "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "Concepto" & Chr(9) & "Cargo" & Chr(9) & "Abono" & Chr(9) & "Saldo"
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado
        Mi_SQL = "SELECT No_Factura,Empresa,Fecha,Cliente_ID,Total,Saldo"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Fecha>='" & Format(Dtp_Fecha_Inicio_Cobranza.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Fecha<='" & Format(Dtp_Fecha_Fin_Cobranza.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Empresa,No_Factura"
        Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Facturas.EOF
            If No_Factura <> Rs_Consulta_Facturas.rdoColumns("No_Factura") Then
                'Consulta la empresa
                If Rs_Consulta_Facturas.rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(1); Format(Rs_Consulta_Facturas.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(2); Rs_Consulta_Facturas.rdoColumns("No_Factura"); _
                    Spc(1); Empresa; _
                    Spc(2); "FACTURA"; _
                    Spc(17 - Len("FACTURA")); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Facturas.rdoColumns("Total"), "#,###,##0.00"), 13); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 13); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Facturas.rdoColumns("Total"), "#,###,##0.00"), 13)
                Grid_Reporte.AddItem Format(Rs_Consulta_Facturas.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & Rs_Consulta_Facturas.rdoColumns("No_Factura") _
                    & Chr(9) & Empresa _
                    & Chr(9) & "FACTURA" _
                    & Chr(9) & Format(Rs_Consulta_Facturas.rdoColumns("Total"), "#,###,##0.00") _
                    & Chr(9) & "0.00" _
                    & Chr(9) & Format(Rs_Consulta_Facturas.rdoColumns("Total"), "#,###,##0.00")
                No_Factura = Rs_Consulta_Facturas.rdoColumns("No_Factura")
                Total_Saldo = Total_Saldo + Val(Rs_Consulta_Facturas.rdoColumns("Saldo"))
                Saldo_Factura = Val(Rs_Consulta_Facturas.rdoColumns("Total"))
            End If
            'Consulta los movimientos
            Mi_SQL = "SELECT No_Movimiento,No_Factura,Empresa,Fecha,Cantidad,Forma_Pago"
            Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Rs_Consulta_Facturas.rdoColumns("No_Factura") & "'"
            Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Facturas.rdoColumns("Empresa")
            Mi_SQL = Mi_SQL & " AND Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Factura_Remision='F'"
            Mi_SQL = Mi_SQL & " AND Tipo='I'"
            Mi_SQL = Mi_SQL & " AND Estatus='A'"
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Movimientos.Fecha"
            Set Rs_Consulta_Adm_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Adm_Movimientos
                While Not .EOF
                    'Consulta la empresa
                    If .rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                    Saldo_Factura = Saldo_Factura - Val(.rdoColumns("Cantidad"))
                    Print #1, Spc(1); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); .rdoColumns("No_Factura"); _
                        Spc(1); Empresa; _
                        Spc(2); Mid(.rdoColumns("Forma_Pago"), 1, 15); _
                        Spc(17 - Len(Mid(.rdoColumns("Forma_Pago"), 1, 15))); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 13); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 13); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Factura, "#,###,##0.00"), 13)
                    Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & .rdoColumns("No_Factura") _
                        & Chr(9) & Empresa _
                        & Chr(9) & .rdoColumns("Forma_Pago") _
                        & Chr(9) & "0.00" _
                        & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0.00") _
                        & Chr(9) & Format(Saldo_Factura, "#,###,##0.00")
                    .MoveNext
                Wend
            End With
            Rs_Consulta_Adm_Movimientos.Close
            Rs_Consulta_Facturas.MoveNext
        Wend
    End If
    Rs_Consulta_Cat_Clientes.Close
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(55); "SALDO ACUMULADO:     "; Format(Total_Saldo, "#,###,##0.00")
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "SALDO ACUMULADO:" & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
    Call Encabezado_Reportes("", "", "", False)
End Sub

Private Sub Btn_Reporte_Comisiones_Click()
    If Fra_Comisiones_Vendedores.Caption = "Comisiones por Vendedor" Then
        Comisiones_Vendedores
    Else
        If Fra_Comisiones_Vendedores.Caption = "Ventas por Vendedor" Then
            Ventas_Vendedor
        Else
            If Cmb_Tipo_Documento.ListIndex = 0 Then
                Estado_Cuenta_General
            Else
                Estado_Cuenta_General_Remisiones
            End If
        End If
    End If
End Sub

Private Sub Btn_Reporte_Devolucion_Click()
    Reporte_Devoluciones_Cliente
End Sub

Private Sub Btn_Reporte_Diario_Click()
Dim Rs_Movimientos As rdoResultset
Dim Deposito As Double
Dim Cargos As Double
Dim No_Factura As String
Dim No_Cheque As String

On Error GoTo HANDLER
    If Cmb_Cuenta_Banco.ListIndex > -1 Then
        Mi_SQL = "SELECT * FROM Adm_Movimientos"
        Mi_SQL = Mi_SQL & " WHERE Fecha BETWEEN '" & Format(Dtp_Inicio_Diario.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fin_Diario.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Banco_ID='" & Format(Cmb_Cuenta_Banco.ItemData(Cmb_Cuenta_Banco.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Movimiento"
        Set Rs_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Movimientos.EOF Then
            Call Encabezado_Reportes("REPORTE DIARIO DE MOVIMIENTOS", Format(Dtp_Inicio_Diario.Value, "dd/MMM/yyyy"), Format(Dtp_Fin_Diario.Value, "dd/MMM/yyyy"), True)
            Print #1, "     MOVIMIENTOS DEL BANCO "; Cmb_Cuenta_Banco.Text
            Grid_Reporte.AddItem "MOVIMIENTOS DEL BANCO " & Cmb_Cuenta_Banco.Text
            Print #1, "     CUENTA   "; Rs_Movimientos.rdoColumns("Cuenta")
            Grid_Reporte.AddItem "MOVIMIENTOS DEL BANCO " & Cmb_Cuenta_Banco.Text
            Print #1,
            Print #1, "_______________________________________________________________________________________________________________________"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem ""
            Print #1, "FECHA EXP.  BENEFICIARIO            DESCRIPCION         REFERENCIA   No.CHEQUE       DEPOSITO      PAGO          SALDO"
            Grid_Reporte.AddItem "FECHA EXP" & Chr(9) & "BENEFICIARIO" & Chr(9) & "DESCRIPCION" & Chr(9) & "REFERENCIA" & Chr(9) & "No.CHEQUE" & Chr(9) & "DEPOSITO" & Chr(9) & "PAGO" & Chr(9) & "SALDO"
            Print #1, "_______________________________________________________________________________________________________________________"
            Print #1,
            Grid_Reporte.AddItem ""
            Deposito = 0
            Cargos = 0
            While Not Rs_Movimientos.EOF
                If Not IsNull(Rs_Movimientos.rdoColumns("No_Factura")) Then
                    No_Factura = Rs_Movimientos.rdoColumns("No_Factura")
                Else
                    No_Factura = ""
                End If
                If Not IsNull(Rs_Movimientos.rdoColumns("No_Cheque")) Then
                    No_Cheque = Rs_Movimientos.rdoColumns("No_Cheque")
                Else
                    No_Cheque = ""
                End If
                If Rs_Movimientos.rdoColumns("Tipo") = "I" Then
                    Print #1, Format(Rs_Movimientos.rdoColumns("Fecha"), "dd/MMM/yyyy"); Spc(1); Mid(Rs_Movimientos.rdoColumns("Beneficiario"), 1, 20); Spc(22 - Len(Mid(Rs_Movimientos("Beneficiario"), 1, 20))); Mid(Rs_Movimientos.rdoColumns("Concepto"), 1, 20); Spc(22 - Len(Mid(Rs_Movimientos.rdoColumns("Concepto"), 1, 20))); Mid(No_Factura, 1, 12); Spc(13 - Len(Mid(No_Factura, 1, 12))); No_Cheque; Spc(13 - Len(No_Cheque)); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 11); Spc(12); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00"), 13)
                    Grid_Reporte.AddItem Format(Rs_Movimientos.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Rs_Movimientos.rdoColumns("Beneficiario") & Chr(9) & Rs_Movimientos.rdoColumns("Concepto") & Chr(9) & No_Factura & Chr(9) & No_Cheque & Chr(9) & Format(Rs_Movimientos.rdoColumns("Cantidad"), "#,##,##0.00") & Chr(9) & Chr(9) & Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00")
                    Deposito = Deposito + Rs_Movimientos.rdoColumns("Cantidad")
                Else
                    Print #1, Format(Rs_Movimientos.rdoColumns("Fecha"), "dd/MMM/yyyy"); Spc(1); Mid(Rs_Movimientos.rdoColumns("Beneficiario"), 1, 20); Spc(22 - Len(Mid(Rs_Movimientos("Beneficiario"), 1, 20))); Mid(Rs_Movimientos.rdoColumns("Concepto"), 1, 20); Spc(22 - Len(Mid(Rs_Movimientos.rdoColumns("Concepto"), 1, 20))); Mid(No_Factura, 1, 12); Spc(13 - Len(Mid(No_Factura, 1, 12))); No_Cheque; Spc(13 - Len(No_Cheque)); Spc(11); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Movimientos.rdoColumns("Cantidad"), "#,##,##0.00"), 11); Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00"), 13)
                    Grid_Reporte.AddItem Format(Rs_Movimientos.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Rs_Movimientos.rdoColumns("Beneficiario") & Chr(9) & Rs_Movimientos.rdoColumns("Concepto") & Chr(9) & No_Factura & Chr(9) & No_Cheque & Chr(9) & Chr(9) & Format(Rs_Movimientos.rdoColumns("Cantidad"), "#,##,##0.00") & Chr(9) & Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00")
                    Cargos = Cargos + Rs_Movimientos.rdoColumns("Cantidad")
                End If
                Rs_Movimientos.MoveNext
            Wend
            Print #1, "-----------------------------------------------------------------------------------------------------------------------"
            Rs_Movimientos.MovePrevious
            Print #1, Spc(80); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Deposito, "#,###,##0.00"), 13)
            Print #1, Spc(91); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cargos, "#,###,##0.00"), 13)
            Print #1,
            Print #1, Spc(90); "SALDO ACTUAL = "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00"), 13)
            Print #1, "-----------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Deposito, "#,##,##0.00")
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Cargos, "#,##,##0.00")
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "SALDO ACTUAL" & Chr(9) & Format(Rs_Movimientos.rdoColumns("Saldo"), "#,###,##0.00")
            Call Encabezado_Reportes("REPORTE DIARIO DE MOVIMIENTOS", "", "", False)
            Fra_Diario_Administrativo.Visible = False
        Else
            MsgBox "No se encontraron movimientos a la fecha seleccionada", vbExclamation
        End If
        Rs_Movimientos.Close
    Else
        MsgBox "Seleccione el banco de la lista", vbExclamation
    End If
    Exit Sub
HANDLER:
    Close #1
    MsgBox Err.Description
End Sub

Private Sub Btn_Reporte_Entradas_Salidas_Click()
    Select Case Fra_Reporte_Entradas_Salidas.Caption
        Case "Reporte de Entradas Producto Terminado"
            Reporte_Entradas                    'Reporte de Entradas de mercancia
        Case "Reporte de Salidas Producto Terminado"
            Reporte_Salidas                     'Reporte de Salidas de mercancia
        Case "Reporte de Entradas Materia Prima"
            Reporte_Entradas_Materia_Prima
        Case "Reporte de Salidas Materia Prima"
            Reporte_Salidas_Materia_Prima
    End Select
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Concentrado_Pedidos
    'DESCRIPCIÓN: Muestra el reporte de los pedidos hechos, con su detalle
    'CREO: Rafael Muñoz
    'FECHA_CREO: 28/Enero/2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Concentrado_Pedidos()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Gran_Total As Double                        'Almacena total del pedido
Dim Cliente As String                           'Almacena el nombre de Cliente

    'Coloca el encabezado de los reportes
    Call Encabezado_Reportes("REPORTE DE PEDIDOS", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    Print #1, "                                                                                  Cantidad     Cantidad                 "
    Print #1, "Codigo           Descripcion                                                       Pedida      Surtida     Estatus      "
'    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Cliente" & Chr(9) & "No. Pedido" & Chr(9) & "Fecha" & Chr(9) & "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad Pedida" & Chr(9) & "Cantidad Surtida" & Chr(9) & "Estatus"
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1,
    'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
    No_Pedido = ""
    Mi_SQL = "SELECT Ope_Pedidos.No_Pedido,Ope_Pedidos.Cliente_ID,Fecha_Pedido,Detalles_Pedidos.Producto_ID,Cantidad,Detalles_Pedidos.Cantidad_Entregada,Detalles_Pedidos.Estatus,Cat_Productos.Nombre,Cat_Productos.Codigo"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Detalles_Pedidos,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.No_Pedido=Detalles_Pedidos.No_Pedido"
    Mi_SQL = Mi_SQL & " AND Detalles_Pedidos.Producto_ID=Cat_Productos.Producto_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Fecha_Pedido>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Fecha_Pedido<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.Fecha_Pedido,Ope_Pedidos.No_Pedido"
    Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Pedidos
        While Not .EOF
            If No_Pedido <> .rdoColumns("No_Pedido") Then
                'Imprime los pedidos
'                Grid_Reporte.AddItem ""
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Pedidos.rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Clientes.EOF Then
                    Print #1, "Cliente: "; Spc(2); Rs_Consulta_Clientes.rdoColumns("Nombre")
'                    Grid_Reporte.AddItem "Cliente:" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre")
                    Cliente = Rs_Consulta_Clientes.rdoColumns("Nombre")
                Else
                    Print #1, "Cliente: ";
'                    Grid_Reporte.AddItem "Cliente:"
                End If
                Rs_Consulta_Clientes.Close
                Print #1, "Pedido : "; Format(.rdoColumns("No_Pedido"), "0000000000"); Spc(5); "Fecha : "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
'                Grid_Reporte.AddItem "Pedido :" & Chr(9) & .rdoColumns("No_Pedido") & Chr(9) & "Fecha :" & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                No_Pedido = .rdoColumns("No_Pedido")
            End If
            Print #1, Spc(2); .rdoColumns("Codigo"); Spc(20 - Len(.rdoColumns("Codigo"))); Mid(.rdoColumns("Nombre"), 1, 52); Spc(57 - Len(Mid(.rdoColumns("Nombre"), 1, 52))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "###,##0"), 8); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad_Entregada"), "###,##0"), 8); Spc(5); .rdoColumns("Estatus")
            Grid_Reporte.AddItem Cliente & Chr(9) & .rdoColumns("No_Pedido") & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & .rdoColumns("Codigo") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Cantidad_Entregada") & Chr(9) & .rdoColumns("Estatus")
            .MoveNext
        Wend
    End With
    Rs_Consulta_Pedidos.Close
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1,
    Grid_Reporte.AddItem ""
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Ventas_Clientes
    'DESCRIPCIÓN: Consulta las ventas que se realizaron durante el intervalo de
    '             tiempo dado por el usuario y puede mostrarlo o no detallado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 01-Julio-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Ventas_Clientes()
Dim Mi_SQL As String
Dim Total_Ventas As Double                              'Indica el total de las ventas obtenidas
Dim Gran_Total As Double
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Ciudad As String                                    'Almacena el nombre de la ciudad
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Rs_Consulta_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Rs_Consulta_Detalles_Facturas As rdoResultset
Dim Encadezado_Reporte As Boolean

    If Rol <> "ALMACEN" And Rol <> "ALMACEN PT" Then
        If Cmb_Producto_Venta.ListIndex > -1 Then
            Call Encabezado_Reportes("REPORTE DE VENTAS POR CLIENTE", Format(DTP_Fecha_Inicio_Historico_Ventas, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Ventas, "dd/MMM/yyyy"), True)
            'Consulta el nombre del cliente
            Mi_SQL = "SELECT * FROM Cat_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.ListIndex), "00000") & "'"
        Else
            MsgBox "Debe seleccionar un cliente", vbExclamation
            Cmb_Producto_Venta.SetFocus
            Exit Sub
        End If
    Else
        Call Encabezado_Reportes("REPORTE DE VENTAS POR CLIENTE", Format(DTP_Fecha_Inicio_Historico_Ventas, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Ventas, "dd/MMM/yyyy"), True)
        'Consulta los clientes con facuras dentro del rango
        Mi_SQL = "SELECT * FROM Cat_Clientes"
        If Chk_Numero_Lote.Value = 1 Then
            If Encadezado_Reporte = False Then
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Codigo" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                Encadezado_Reporte = True
            End If
        End If
    End If
    Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Gran_Total = 0
    While Not Rs_Consulta_Clientes.EOF
        With Rs_Consulta_Clientes
            If Not IsNull(.rdoColumns("Ciudad_ID")) Then
                Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & .rdoColumns("Ciudad_ID") & "'"
                Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Ciudades.EOF Then
                    Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                Else
                    Ciudad = ""
                End If
                Rs_Consulta_Ciudades.Close
            Else
                Ciudad = ""
            End If
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Print #1, Spc(2); "Cliente:   "; .rdoColumns("Nombre")
            Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
            Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
            Print #1, Spc(2); "Ciudad:    "; Ciudad
            If Chk_Numero_Lote.Value = 0 Then
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & .rdoColumns("Nombre")
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End If
        End With
        Total_Ventas = 0
        'Consulta las facturas del cliente y las fechas seleccionadas
        Mi_SQL = "SELECT No_Factura,Empresa,Pagada,Cancelada,Cliente_ID,Total,Fecha"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
        If Chk_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Fecha"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Chk_Numero_Lote.Value = 0 Then
                Print #1,
                Print #1, Conectar_Ayudante.Centra("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Factura        Fecha              Importe        Estatus"
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem "FACTURAS"
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            End If
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Sin Pagar"
                    End If
                End If
                If Chk_Numero_Lote.Value = 0 Then
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                        Gran_Total = Gran_Total + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                Else
                    Print #1,
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "   No. Factura        Fecha              Importe        Estatus"
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); Spc(6); Estatus
                    End If
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    'Consulta el detalle de la factura
                    Mi_SQL = "SELECT * FROM Adm_Detalles_Facturas"
                    Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & "'"
                    Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")
                    Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    While Not Rs_Consulta_Detalles_Facturas.EOF
                        Print #1, Spc(1); Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"); Spc(15 - Len(Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"))); Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50); Spc(55 - Len(Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0"), 9); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe") + Rs_Consulta_Detalles_Facturas.rdoColumns("IVA"), "#,###,##0.00"), 13)
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID") _
                            & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") _
                            & Chr(9) & "F -" & Empresa _
                            & Chr(9) & Estatus _
                            & Chr(9) & Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo") _
                            & Chr(9) & Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(10)), Chr(13)) _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Precio"), "#,###,##0.00") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe"), "#,###,##0.00")
                        Total_Ventas = Total_Ventas + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Gran_Total = Gran_Total + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Rs_Consulta_Detalles_Facturas.MoveNext
                    Wend
                    Rs_Consulta_Detalles_Facturas.Close
                End If
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
        Else
            If Chk_Numero_Lote.Value = 0 Then
                Print #1,
                Print #1, Conectar_Ayudante.Centra("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Factura        Fecha              Importe        Estatus                                                         "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   SIN FACTURAS                                                                                                         "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem "SIN FACTURAS"
            End If
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        If Chk_Numero_Lote.Value = 1 Then
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1,
        End If
        Print #1,
        Print #1, Spc(35); "-------------------"
        Print #1, Spc(18); "Total Venta"; Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,##0.00"), 14)
        Print #1,
        If Encadezado_Reporte = False Then
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Venta" & Chr(9) & Format(Total_Ventas, "#,###,##0.00")
            Encadezado_Reporte = True
        End If
        'Consulta las remisiones del cliente y las fechas seleccionadas
        Mi_SQL = "SELECT No_Remision,Empresa,Pagada,Cancelada,Cliente_ID,Total,Fecha"
        Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
        If Chk_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Fecha"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Chk_Numero_Lote.Value = 0 Then
                Print #1,
                Print #1, Conectar_Ayudante.Centra("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Remisiones     Fecha              Importe        Estatus                                                         "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem "REMISIONES"
                Grid_Reporte.AddItem "No. Remisiones" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            Else
                If Encadezado_Reporte = False Then
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Codigo" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                    Encadezado_Reporte = True
                End If
            End If
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Sin Pagar"
                    End If
                End If
                If Chk_Numero_Lote.Value = 0 Then
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                        Gran_Total = Gran_Total + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); Spc(6); Estatus
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                Else
                    Print #1,
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "   No. Remision       Fecha              Importe        Estatus"
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(6); Estatus
                    Else
                        Print #1, Spc(1); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision"); Spc(2); Empresa; _
                            Spc(3); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                            Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); Spc(6); Estatus
                    End If
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    'Consulta el detalle de la factura
                    Mi_SQL = "SELECT * FROM Adm_Detalles_Remisiones"
                    Mi_SQL = Mi_SQL & " WHERE No_Remision='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") & "'"
                    Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")
                    Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    While Not Rs_Consulta_Detalles_Facturas.EOF
                        Print #1, Spc(1); Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"); Spc(15 - Len(Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"))); Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50); Spc(55 - Len(Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0"), 9); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe") + Rs_Consulta_Detalles_Facturas.rdoColumns("IVA"), "#,###,##0.00"), 13)
                        Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID") _
                            & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Remision") _
                            & Chr(9) & "R -" & Empresa _
                            & Chr(9) & Estatus _
                            & Chr(9) & Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo") _
                            & Chr(9) & Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(10)), Chr(13)) _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Precio"), "#,###,##0.00") _
                            & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe"), "#,###,##0.00")
                        Total_Ventas = Total_Ventas + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Gran_Total = Gran_Total + Rs_Consulta_Detalles_Facturas.rdoColumns("Importe")
                        Rs_Consulta_Detalles_Facturas.MoveNext
                    Wend
                    Rs_Consulta_Detalles_Facturas.Close
                End If
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
        Else
            If Chk_Numero_Lote.Value = 0 Then
                Print #1,
                Print #1, Conectar_Ayudante.Centra("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
                Print #1,
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   No. Remision       Fecha              Importe        Estatus"
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   SIN REMISIONES                                                                                                         "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha Venta" & Chr(9) & "Importe" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem "SIN REMISIONES"
            End If
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        Rs_Consulta_Clientes.MoveNext
    Wend
    Rs_Consulta_Clientes.Close
    If Chk_Numero_Lote.Value = 1 Then
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1,
    End If
    Print #1, Spc(35); "-------------------"
    Print #1, Spc(18); "Gran Total "; Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Venta" & Chr(9) & Format(Gran_Total, "#,###,##0.00")
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Reporte_Pedidos
    'DESCRIPCION: Muestra el reporte de los pedidos de acuerdo al filtro seleccionado
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 14-Mayo-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACION:
    'PARAMETROS:
'*******************************************************************************
Private Sub Reporte_Pedidos()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset        'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Ciudad As String                            'Almacena el nombre de la ciudad
Dim Total As Double                             'Almacena el total del pedido
Dim Total_Producto As Double                    'Almacena el total de productos
Dim Productos_Pedido As Double                  'Almacena el total del producto por pedido

    'Coloca el encabezado de los reportes
    Call Encabezado_Reportes("REPORTE DE PEDIDOS", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
    Mi_SQL = "SELECT Ope_Pedidos.No_Pedido,Ope_Pedidos.Total,Fecha_Pedido,Ope_Pedidos.Cliente_ID,Detalles_Pedidos.Producto_ID,Detalles_Pedidos.Descripcion, Detalles_Pedidos.Cantidad,Detalles_Pedidos.Importe,Cat_Productos.Codigo,Cat_Productos.Producto_ID"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Detalles_Pedidos,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.No_Pedido=Detalles_Pedidos.No_Pedido"
    Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID=Detalles_Pedidos.Producto_ID"
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Reporte_Pedido_Producto.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Cliente_ID='" & Format(Cmb_Reporte_Pedido_Producto.ItemData(Cmb_Reporte_Pedido_Producto.ListIndex), "00000") & "'"
    If Trim(Txt_Reporte_No_Pedido.Text) <> "" Then Mi_SQL = Mi_SQL & " AND Ope_Pedidos.No_Pedido='" & Format(Trim(Txt_Reporte_No_Pedido.Text), "0000000000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.Fecha_Pedido,Ope_Pedidos.No_Pedido"
    Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Pedidos
        While Not .EOF
            If No_Pedido <> .rdoColumns("No_Pedido") Then
            Mi_SQL = "SELECT Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad"
            Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Ciudades"
            Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Cliente_ID = '" & Rs_Consulta_Pedidos.rdoColumns("Cliente_ID") & "'"
            Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Clientes
                If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                    Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                    Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                    Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                    Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Ciudades.EOF Then
                        Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                    Else
                        Ciudad = ""
                    End If
                    Rs_Consulta_Ciudades.Close
                Else
                    Ciudad = ""
                End If
                Print #1,
                If Productos_Pedido > 0 Then
                    Print #1, Spc(90); "Total Productos:  "; Format(Productos_Pedido, "###,###")
                    Grid_Reporte.AddItem "Total Producto:" & Chr(9) & Format(Productos_Pedido, "###,###")
                    Productos_Pedido = 0
                    Print #1,
                    Grid_Reporte.AddItem ""
                End If
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Print #1, Spc(2); "Cliente:   "; Rs_Consulta_Clientes.rdoColumns("Nombre")
                Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre")
                Grid_Reporte.AddItem "Direccion:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia: " & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End With
                'Imprime los pedidos
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(1); "No. Pedido : "; .rdoColumns("No_Pedido"); Spc(5); "Fecha : "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(3); "Total Pedido:"; Format(.rdoColumns("Total"), "#,###,##0.00")
                Grid_Reporte.AddItem "Pedido :" & Chr(9) & .rdoColumns("No_Pedido") & Chr(9) & "Fecha :" & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") & Chr(9) & "Total Pedido: " & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, " Codigo                 Descripcion                                                      Cantidad        Importe"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Importe"
                Print #1, "________________________________________________________________________________________________________________________"
                Print #1,
                No_Pedido = .rdoColumns("No_Pedido")
            End If
            Print #1, Spc(2); .rdoColumns("Codigo"); Spc(8); Mid(.rdoColumns("Descripcion"), 1, 59); Spc(62 - Len(Mid(.rdoColumns("Descripcion"), 1, 59))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "###,###"), 10); Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Importe")), "###,##0.00"), 10)
            Grid_Reporte.AddItem .rdoColumns("Codigo") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Importe")
            Productos_Pedido = Productos_Pedido + .rdoColumns("Cantidad")
            Total_Producto = Total_Producto + .rdoColumns("Cantidad")
            Total = Total + .rdoColumns("Importe")
            .MoveNext
        Wend
    End With
    Rs_Consulta_Pedidos.Close
    Print #1,
    Print #1, Spc(90); "Total Productos:  "; Format(Productos_Pedido, "###,###")
    Grid_Reporte.AddItem "Total Productos:" & Chr(9) & Format(Productos_Pedido, "###,###")
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(83); "Total Productos: "; Format(Total_Producto, "###,###")
    Print #1, Spc(83); "GRAN TOTAL:      "; Format(Total, "#,###,##0.00")
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Productos:" & Chr(9) & Format(Total_Producto, "###,###")
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "GRAN TOTAL:" & Chr(9) & Format(Total, "#,###,##0.00")
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Pedidos
    'DESCRIPCIÓN: Muestra el reporte de los pedidos hechos, con su detalle
    'CREO: Rafael Muñoz
    'FECHA_CREO: 31/Enero/2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Pedidos_Productos()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Total As Double                             'Almacena el total del producto vendido
Dim Total_Producto As Double                    'Almacena el total de los productos pedidos

    If Cmb_Reporte_Pedido_Producto.ListIndex > -1 Then
        'Coloca el encabezado de los reportes
        Call Encabezado_Reportes("REPORTE DE PEDIDOS POR PRODUCTO", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
        Print #1, "Producto: "; Cmb_Reporte_Pedido_Producto.Text
        Grid_Reporte.AddItem ""
        'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
        Mi_SQL = "SELECT Ope_Pedidos.No_Pedido,Ope_Pedidos.Total,Fecha_Pedido,Detalles_Pedidos.Producto_ID,Cat_Productos.Producto_ID,Detalles_Pedidos.Descripcion, Detalles_Pedidos.Cantidad,Detalles_Pedidos.Importe, Detalles_Pedidos.Precio"
        Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Cat_Productos,Detalles_Pedidos"
        Mi_SQL = Mi_SQL & " WHERE Detalles_Pedidos.Producto_ID= '" & Format(Cmb_Reporte_Pedido_Producto.ItemData(Cmb_Reporte_Pedido_Producto.ListIndex), "00000") & "' AND Cat_Productos.Producto_Id=Detalles_Pedidos.Producto_ID "
        Mi_SQL = Mi_SQL & " AND Ope_Pedidos.No_Pedido=Detalles_Pedidos.No_Pedido"
        Mi_SQL = Mi_SQL & " AND Fecha_Pedido>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Fecha_Pedido<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.Fecha_Pedido,Ope_Pedidos.No_Pedido"
        Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Pedidos
            While Not .EOF
                If No_Pedido <> .rdoColumns("No_Pedido") Then
                    'Imprime los pedidos
                    Print #1,
                    Grid_Reporte.AddItem ""
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "No. Pedido : "; .rdoColumns("No_Pedido"); Spc(5); "Fecha : "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                    Grid_Reporte.AddItem "Pedido :" & Chr(9) & .rdoColumns("No_Pedido") & Chr(9) & "Fecha :" & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                    Print #1, "------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "    Cantidad                    Precio                     Importe"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Producto_ID" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
                    Print #1, "________________________________________________________________________________________________________________________"
                    No_Pedido = .rdoColumns("No_Pedido")
                End If
                Print #1, Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "###,###"), 7); Spc(20); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Precio")), "###,##0.00"), 7); Spc(20); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Importe")), "###,##0.00"), 7)
                Grid_Reporte.AddItem .rdoColumns("Producto_ID") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Precio") & Chr(9) & .rdoColumns("Importe")
                Total = Total + .rdoColumns("Importe")
                Total_Producto = Total_Producto + .rdoColumns("Cantidad")
                .MoveNext
            Wend
        End With
        Rs_Consulta_Pedidos.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(40); "Total de producto:"; Total_Producto
        Print #1, Spc(40); "GRAN TOTAL:       "; Format(Total, "#,###,##0.00")
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "Total Productos:" & Chr(9) & Format(Total_Producto, "###,###")
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "Seleccione un producto de la lista", vbInformation, "NATURAL HEALTH"
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Reporte_Facturas_Productos
    'DESCRIPCION: Muestra el reporte de los productos facturados
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 13-Mayo-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACION:
    'PARAMETROS:
'*******************************************************************************
Private Sub Reporte_Facturas_Productos()
Dim Mi_SQL As String
Dim Rs_Consulta_Datos_Factura As rdoResultset
Dim Rs_Consulta_Datos_Pedido As rdoResultset
Dim Empresa As String
Dim No_Pedido As String
Dim Total As Double
Dim Total_Producto As Double

    If Cmb_Reporte_Pedido_Producto.ListIndex > -1 Then
        'Coloca el encabezado de los reportes
        Call Encabezado_Reportes("REPORTE DE PRODUCTOS POR FACTURA", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
        Print #1, "PRODUCTO : "; Cmb_Reporte_Pedido_Producto.Text
        Grid_Reporte.AddItem Cmb_Reporte_Pedido_Producto.Text
        'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
        Mi_SQL = "SELECT Adm_Facturas_Clientes.No_Factura,Adm_Facturas_Clientes.Empresa,Adm_Facturas_Clientes.Fecha,Adm_Facturas_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " ,Adm_Detalles_Facturas.Cantidad,Adm_Detalles_Facturas.Precio,Adm_Detalles_Facturas.Importe,Cat_Clientes.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Producto_ID='" & Format(Cmb_Reporte_Pedido_Producto.ItemData(Cmb_Reporte_Pedido_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha>='" & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha<='" & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Facturas_Clientes.Fecha"
        Set Rs_Consulta_Datos_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Datos_Factura
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, " Cliente                                     Factura         Fecha         Pedido     Cantidad    Precio       Importe  "
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Cliente" & Chr(9) & "Factura" & Chr(9) & "Fecha" & Chr(9) & "Pedido" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If .rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                'Consulta el número del pedido
                'Consulta los datos de No_Pedido y Orden_Compra
                Mi_SQL = "SELECT DISTINCT Detalles_Salidas.No_Salida,Ope_Salidas.No_Pedido"
                Mi_SQL = Mi_SQL & " FROM Detalles_Salidas,Ope_Salidas"
                Mi_SQL = Mi_SQL & " WHERE Detalles_Salidas.No_Salida=Ope_Salidas.No_Salida"
                Mi_SQL = Mi_SQL & " AND Detalles_Salidas.No_Factura='" & .rdoColumns("No_Factura") & "'"
                Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Empresa_Factura=" & .rdoColumns("Empresa")
                Set Rs_Consulta_Datos_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Datos_Pedido.EOF Then
                    No_Pedido = Rs_Consulta_Datos_Pedido.rdoColumns("No_Pedido")
                Else
                    No_Pedido = "          "
                End If
                Rs_Consulta_Datos_Pedido.Close
                Print #1, Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); .rdoColumns("No_Factura"); " "; Empresa; _
                    Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(3); No_Pedido; _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0"), 7); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Precio")), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Importe")), "#,###,##0.00"), 14)
                Grid_Reporte.AddItem .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("No_Factura") & " " & Empresa _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & No_Pedido _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0") _
                    & Chr(9) & Format(.rdoColumns("Precio"), "#,###,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Importe"), "#,###,##0.00")
                Total_Producto = Total_Producto + .rdoColumns("Cantidad")
                Total = Total + .rdoColumns("Importe")
                .MoveNext
            Wend
        End With
        Rs_Consulta_Datos_Factura.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(40); "TOTAL DE PIEZAS :  "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Producto, "#,###,##0"), 14)
        Print #1, Spc(40); "GRAN TOTAL :       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTAL DE PIEZAS : " & Chr(9) & Format(Total_Producto, "###,##0")
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL :      " & Chr(9) & Format(Total, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "Seleccione un producto de la lista", vbExclamation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Reporte_Remisiones_Productos
    'DESCRIPCION: Muestra el reporte de los productos remisionados
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 19-Mayo-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACION:
    'PARAMETROS:
'*******************************************************************************
Private Sub Reporte_Remisiones_Productos()
Dim Mi_SQL As String
Dim Rs_Consulta_Datos_Factura As rdoResultset
Dim Rs_Consulta_Datos_Pedido As rdoResultset
Dim Empresa As String
Dim No_Pedido As String
Dim Total As Double
Dim Total_Producto As Double

    If Cmb_Reporte_Pedido_Producto.ListIndex > -1 Then
        'Coloca el encabezado de los reportes
        Call Encabezado_Reportes("REPORTE DE PRODUCTOS POR REMISION", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
        Print #1, "PRODUCTO : "; Cmb_Reporte_Pedido_Producto.Text
        Grid_Reporte.AddItem Cmb_Reporte_Pedido_Producto.Text
        'Consulta los pedidos solicitados a la matriz y su detalle, incluyendo las tablas de productos y sucursales para tomar datos a mostrar
        Mi_SQL = "SELECT Adm_Remisiones_Clientes.No_Remision,Adm_Remisiones_Clientes.Empresa,Adm_Remisiones_Clientes.Fecha,Adm_Remisiones_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " ,Adm_Detalles_Remisiones.Cantidad,Adm_Detalles_Remisiones.Precio,Adm_Detalles_Remisiones.Importe,Cat_Clientes.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Adm_Detalles_Remisiones,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.No_Remision=Adm_Detalles_Remisiones.No_Remision"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Empresa=Adm_Detalles_Remisiones.Empresa"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Detalles_Remisiones.Producto_ID='" & Format(Cmb_Reporte_Pedido_Producto.ItemData(Cmb_Reporte_Pedido_Producto.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Fecha>='" & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Fecha<='" & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Remisiones_Clientes.Fecha"
        Set Rs_Consulta_Datos_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Datos_Factura
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, " Cliente                                    Remision         Fecha         Pedido     Cantidad    Precio       Importe  "
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Cliente" & Chr(9) & "Remision" & Chr(9) & "Fecha" & Chr(9) & "Pedido" & Chr(9) & "Cantidad" & Chr(9) & "Precio" & Chr(9) & "Importe"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If .rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                'Consulta el número del pedido
                'Consulta los datos de No_Pedido y Orden_Compra
                Mi_SQL = "SELECT DISTINCT Detalles_Salidas.No_Salida,Ope_Salidas.No_Pedido"
                Mi_SQL = Mi_SQL & " FROM Detalles_Salidas,Ope_Salidas"
                Mi_SQL = Mi_SQL & " WHERE Detalles_Salidas.No_Salida=Ope_Salidas.No_Salida"
                Mi_SQL = Mi_SQL & " AND Detalles_Salidas.No_Remision='" & .rdoColumns("No_Remision") & "'"
                Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Empresa_Factura=" & .rdoColumns("Empresa")
                Set Rs_Consulta_Datos_Pedido = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Datos_Pedido.EOF Then
                    No_Pedido = Rs_Consulta_Datos_Pedido.rdoColumns("No_Pedido")
                Else
                    No_Pedido = "          "
                End If
                Rs_Consulta_Datos_Pedido.Close
                Print #1, Mid(.rdoColumns("Nombre"), 1, 40); _
                    Spc(42 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); .rdoColumns("No_Remision"); " "; Empresa; _
                    Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(3); No_Pedido; _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0"), 7); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Precio")), "#,##0.00"), 8); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Importe")), "#,###,##0.00"), 14)
                Grid_Reporte.AddItem .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("No_Remision") & " " & Empresa _
                    & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & No_Pedido _
                    & Chr(9) & Format(.rdoColumns("Cantidad"), "#,###,##0") _
                    & Chr(9) & Format(.rdoColumns("Precio"), "#,###,##0.00") _
                    & Chr(9) & Format(.rdoColumns("Importe"), "#,###,##0.00")
                Total_Producto = Total_Producto + .rdoColumns("Cantidad")
                Total = Total + .rdoColumns("Importe")
                .MoveNext
            Wend
        End With
        Rs_Consulta_Datos_Factura.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(40); "TOTAL DE PIEZAS :  "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Producto, "#,###,##0"), 14)
        Print #1, Spc(40); "GRAN TOTAL :       "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "TOTAL DE PIEZAS : " & Chr(9) & Format(Total_Producto, "###,##0")
        Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & "GRAN TOTAL :      " & Chr(9) & Format(Total, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "Seleccione un producto de la lista", vbExclamation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Btn_Reporte_Historico_Facturas_Click
    'DESCRIPCIÓN: Consulta las factutas que fueron dadas de alta en la sucursal
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 18-Marzo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_General_Facturas()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Total_IVA As Double                                 'Almacena el total del IVA
Dim Total_Tasa_0 As Double                              'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                             'Almacena el total de la tasa 15
Dim Total_Facturas As Double                            'Almacena el total de la factura

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Factura,Empresa,Cliente_ID,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Saldo,Abono"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
    Mi_SQL = Mi_SQL & " WHERE (Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & "')"
    If Cmb_Empresa_Facturas.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Clasificacion_Factura='" & Cmb_Empresa_Facturas.Text & "'"
    End If
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Pagada='S'"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND (Pagada='S' OR Cancelada='S')"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Cancelada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Pagada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Cancelada='S' "
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND (Pagada='N' AND Cancelada='N')"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            'Valida que exista la ruta si no la crea
            If Len(Dir("C:\REP", vbDirectory)) = 0 Then
                MkDir "C:\REP"
            End If
            Open "C:\REP\Reportes.txt" For Output As #1
            Print #1, "************************************************************************************************************************"
            Print #1,
            If Cmb_Empresa_Facturas.ListIndex = 0 Then
                Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            Else
                Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            End If
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, Conectar_Ayudante.Centra("REPORTE DE GENERAL DE FACTURAS   [" & Cmb_Empresa_Facturas.Text & "]", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "REPORTE DE GENERAL DE FACTURAS   [" & Cmb_Empresa_Facturas.Text & "]"
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, Conectar_Ayudante.Centra("DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy")
            Print #1, "************************************************************************************************************************"
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            If Rol = "CUENTAS POR PAGAR" Then
                Print #1, "                                                             Ventas       Ventas                                         "
                Print #1, "  No. Factura     Cliente                       Fecha        Sin IVA      Con IVA        IVA         Total     Estatus   "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Venta Sin IVA" & Chr(9) & "Venta Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            Else
                Print #1, "  No. Factura     Cliente                       Fecha        Abono        Saldo          IVA         Total     Estatus   "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            End If
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            While Not .EOF
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
                If .rdoColumns("Pagada") = "S" Then
                    Estatus = "Pagada"
                End If
                If .rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                End If
                If .rdoColumns("Pagada") = "N" And .rdoColumns("Cancelada") = "N" Then
                    Estatus = "Pendiente"
                End If
                If .rdoColumns("Cancelada") = "N" Then
                    If Rol = "CUENTAS POR PAGAR" Then
                        Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 30); _
                            Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "#,###,##0.00"), 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "#,###,##0.00"), 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                            Spc(2); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                            Format(.rdoColumns("Subtotal_Tasa_0"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("Subtotal_Tasa_15"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("IVA"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_IVA = Total_IVA + .rdoColumns("IVA")
                        Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Subtotal_Tasa_0")
                        Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Subtotal_Tasa_15")
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Else
                        Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 30); _
                            Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                            Spc(2); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                            Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("IVA"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_IVA = Total_IVA + .rdoColumns("IVA")
                        Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Abono")
                        Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Saldo")
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    End If
                Else
                    If Rol = "CUENTAS POR PAGAR" Then
                        Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 30); _
                            Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 10); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                            Spc(2); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                            "0.00" & Chr(9) & _
                            "0.00" & Chr(9) & _
                            "0.00" & Chr(9) & _
                            "0.00" & Chr(9) & Estatus
                    Else
                        Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 30); _
                            Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 10); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                            Spc(2); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                            "0.00" & Chr(9) & _
                            "0.00" & Chr(9) & _
                            "0.00" & Chr(9) & _
                            Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                    End If
                End If
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(26); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 13); Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 13); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_IVA, "###,##0.00"), 10); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13305
            Me.Height = 8340
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_General_Remisiones
    'DESCRIPCIÓN: Muestra el reporte general de las remisiones
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_General_Remisiones()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Total_IVA As Double                                 'Almacena el total del IVA
Dim Total_Tasa_0 As Double                              'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                             'Almacena el total de la tasa 15
Dim Total_Facturas As Double                            'Almacena el total de la factura

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Remision,Empresa,Cliente_ID,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Saldo,Abono"
    Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes"
    Mi_SQL = Mi_SQL & " WHERE (Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & "')"
    If Cmb_Empresa_Facturas.ListIndex > 0 Then
        If Cmb_Empresa_Facturas.Text <> "PRONACEN" Then
            Mi_SQL = Mi_SQL & " AND Empresa=1"
        Else
            Mi_SQL = Mi_SQL & " AND Empresa=2"
        End If
    End If
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Pagada='S'"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND (Pagada='S' OR Cancelada='S')"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Cancelada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Pagada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Cancelada='S' "
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND (Pagada='N' AND Cancelada='N')"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            'Valida que exista la ruta si no la crea
            If Len(Dir("C:\REP", vbDirectory)) = 0 Then
                MkDir "C:\REP"
            End If
            Open "C:\REP\Reportes.txt" For Output As #1
            Print #1, "************************************************************************************************************************"
            Print #1,
            Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, Conectar_Ayudante.Centra("REPORTE DE GENERAL DE FACTURAS", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "REPORTE DE GENERAL DE FACTURAS"
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, Conectar_Ayudante.Centra("DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy")
            Print #1, "************************************************************************************************************************"
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            If Rol = "CUENTAS POR PAGAR" Then
                Print #1, "                                                             Ventas       Ventas                                         "
                Print #1, " No. Remision     Cliente                       Fecha        Sin IVA       Con IVA        IVA         Total     Estatus   "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Venta Sin IVA" & Chr(9) & "Venta Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            Else
                Print #1, " No. Remision     Cliente                       Fecha        Abono        Saldo          IVA         Total     Estatus   "
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Abono" & Chr(9) & "Saldo" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
            End If
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            While Not .EOF
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
                If .rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If .rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Pendiente"
                    End If
                End If
                If .rdoColumns("Cancelada") = "N" Then
                    Print #1, Spc(1); .rdoColumns("No_Remision"); Spc(2); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "###,##0.00"), 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("IVA"), "#,###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                    Total_IVA = Total_IVA + .rdoColumns("IVA")
                    Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Abono")
                    Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Saldo")
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Else
                    Print #1, Spc(1); .rdoColumns("No_Remision"); Spc(2); Mid(Cliente, 1, 30); _
                        Spc(32 - Len(Mid(Cliente, 1, 30))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 10); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 12); _
                        Spc(2); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Cliente & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        "0.00" & Chr(9) & _
                        Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                End If
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(26); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 13); Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 13); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_IVA, "###,##0.00"), 10); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13305
            Me.Height = 8340
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay remisiones en las fechas que proporcionó", vbExclamation
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Reporte_Facturas_Proveedores
'DESCRIPCION: Consulta las facturas de proveedor según los filtros de búsqueda
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 10-Junio-2008
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Reporte_Facturas_Proveedores()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Proveedores As rdoResultset
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Proveedor As String
Dim Total_Facturas As Double                            'Almacena el total de la factura

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Factura,Empresa,Proveedor_ID,Fecha,Tipo,Total,Pagada,Cancelada"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE (Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & "')"
    Select Case Cmb_Empresa_Facturas.Text
        Case "NATURAL HEALTH"
            Mi_SQL = Mi_SQL & " AND Empresa=1"
        Case "PRONACEN"
            Mi_SQL = Mi_SQL & " AND Empresa=2"
        Case "GRUPO NH"
            Mi_SQL = Mi_SQL & " AND Empresa=4"
    End Select
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Pagada='S'"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND (Pagada='S' OR Cancelada='S')"
    If Chk_Facturas_Pagadas.Value = 1 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Cancelada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND Pagada<>'S'"
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 1 And Chk_Facturas_Pendientes.Value = 0 Then Mi_SQL = Mi_SQL & " AND Cancelada='S' "
    If Chk_Facturas_Pagadas.Value = 0 And Chk_Facturas_Canceladas.Value = 0 And Chk_Facturas_Pendientes.Value = 1 Then Mi_SQL = Mi_SQL & " AND (Pagada='N' AND Cancelada='N')"
    'Check para las facturas sin presupuesto
    If Chk_Sin_Presupuesto.Value = 1 Then
        'Consulta que no esté en algún prorrateo
        Mi_SQL = Mi_SQL & " AND No_Factura NOT IN(SELECT No_Factura FROM Ope_Presupuestos_Prorrateo)"
        Mi_SQL = Mi_SQL & " AND Proveedor_ID NOT IN(SELECT Proveedor_ID FROM Ope_Presupuestos_Prorrateo)"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Facturas_Proveedores.EOF Then
        With Rs_Consulta_Facturas_Proveedores
            'Valida que exista la ruta si no la crea
            If Len(Dir("C:\REP", vbDirectory)) = 0 Then
                MkDir "C:\REP"
            End If
            Open "C:\REP\Reportes.txt" For Output As #1
            Print #1, "************************************************************************************************************************"
            Print #1,
            Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, Conectar_Ayudante.Centra("REPORTE DE GENERAL DE FACTURAS PROVEEDORES", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "REPORTE DE GENERAL DE FACTURAS PROVEEDORES"
            If Chk_Sin_Presupuesto.Value = 1 Then
                Print #1, Conectar_Ayudante.Centra("SIN ASIGNACION A PRESUPUESTO", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "SIN ASIGNACION A PRESUPUESTO"
            End If
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, Conectar_Ayudante.Centra("DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy") & " AL " & Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy")
            Print #1, "************************************************************************************************************************"
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Factura     Proveedor                               Fecha         Tipo                     Total      Estatus      "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Fecha" & Chr(9) & "Tipo" & Chr(9) & "Total" & Chr(9) & "Estatus"
            Grid_Reporte.AddItem ""
            While Not .EOF
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Proveedor_ID,Nombre FROM Cat_Proveedores"
                Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & .rdoColumns("Proveedor_ID") & "'"
                Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Proveedores.EOF Then
                    Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                Else
                    Proveedor = ""
                End If
                Rs_Consulta_Cat_Proveedores.Close
                If .rdoColumns("Cancelada") = "S" Then
                    Estatus = "CANCELADA"
                Else
                    If .rdoColumns("Pagada") = "S" Then
                        Estatus = "PAGADA"
                    Else
                        Estatus = "PENDIENTE"
                    End If
                End If
                If .rdoColumns("Cancelada") = "N" Then
                    Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                        Spc(15 - Len(Trim(.rdoColumns("No_Factura")))); Mid(Proveedor, 1, 37); _
                        Spc(40 - Len(Mid(Proveedor, 1, 37))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(5); UCase(.rdoColumns("Tipo")); _
                        Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                        Spc(3); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Proveedor & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                        & Chr(9) & .rdoColumns("Tipo") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Else
                    Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                        Spc(15 - Len(Trim(.rdoColumns("No_Factura")))); Mid(Proveedor, 1, 37); _
                        Spc(40 - Len(Mid(Proveedor, 1, 37))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(5); UCase(.rdoColumns("Tipo")); _
                        Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                        Spc(3); Estatus
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Proveedor & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                        & Chr(9) & .rdoColumns("Tipo") & Chr(9) & "0.00" & Chr(9) & Estatus
                End If
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(61); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13305
            Me.Height = 8340
            Me.Left = 0
            Me.Top = 0
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Facturas_Proveedores.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Estado_Cuenta_General
    'DESCRIPCIÓN: Consulta el estado de cuenta de las facturas
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 04-Abril-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Estado_Cuenta_General()
Dim Mi_SQL As String
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Vendedor As String
Dim Telefono_Cliente As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Total_Pagos As Double                               'Almacena el total de los pagos hechos en las facturas
Dim Total_Saldo As Double                               'Almacena el total del saldo de las facturas
Dim Total_Facturas As Double                            'Almacena el total de la factura

    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Factura,Adm_Facturas_Clientes.Cliente_ID,Empresa,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Abono,Saldo"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Nombre,Cat_Clientes.Vendedor_ID,Cat_Clientes.Telefono,Cat_Vendedores.Clave"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes,Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha>=" & Par_Fecha & Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha<=" & Par_Fecha & Format(Dtp_Comision_Fecha_Fin.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Vendedor_Comision.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID='" & Format(Cmb_Vendedor_Comision.ItemData(Cmb_Vendedor_Comision.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Empresa,No_Factura"
    Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
        With Rs_Consulta_Adm_Facturas_Clientes
            Call Encabezado_Reportes("ESTADO DE CUENTA GENERAL", Format(Dtp_Comision_Fecha_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Comision_Fecha_Fin.Value, "dd/MMM/yyyy"), True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "Empresa  Factura     Fecha          Importe            Pagos             Saldo       Cliente                    Vendedor "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Empresa" & Chr(9) & "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Cliente" & Chr(9) & "Vendedor" & Chr(9) & "Telefono"
            Grid_Reporte.AddItem ""
            While Not .EOF
                Cliente = .rdoColumns("Nombre")
                Vendedor = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Telefono")) Then
                    Telefono_Cliente = .rdoColumns("Telefono")
                Else
                    Telefono_Cliente = ""
                End If
                If .rdoColumns("Empresa") = 1 Then
                    Print #1, "  NHE  "; .rdoColumns("No_Factura"); _
                        Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                        Spc(2); Mid(Cliente, 1, 33); _
                        Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor
                    Grid_Reporte.AddItem "  NHE  " & Chr(9) & .rdoColumns("No_Factura") _
                        & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                        & Chr(9) & Cliente & Chr(9) & Vendedor
                Else
                     Print #1, "  PNC  "; .rdoColumns("No_Factura"); _
                        Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                        Spc(2); Mid(Cliente, 1, 33); _
                        Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor
                    Grid_Reporte.AddItem "  PNC  " & Chr(9) & .rdoColumns("No_Factura") _
                        & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                        & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente
                End If
                Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Adm_Facturas_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Estado_Cuenta_General_Remisiones
    'DESCRIPCIÓN: Consulta el estado de cuenta de las remisiones
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Estado_Cuenta_General_Remisiones()
Dim Mi_SQL As String
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Vendedor As String
Dim Telefono_Cliente As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Remisiones_Clientes As rdoResultset   'Manejo de registro
Dim Total_Pagos As Double                               'Almacena el total de los pagos hechos en las facturas
Dim Total_Saldo As Double                               'Almacena el total del saldo de las facturas
Dim Total_Facturas As Double                            'Almacena el total de la factura

    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Remision,Adm_Remisiones_Clientes.Cliente_ID,Empresa,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Abono,Saldo"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Nombre,Cat_Clientes.Vendedor_ID,Cat_Clientes.Telefono,Cat_Vendedores.Clave"
    Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes,Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    Mi_SQL = Mi_SQL & " AND Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Comision_Fecha_Fin.Value, "MM/dd/yyyy") & Par_Fecha
    If Cmb_Vendedor_Comision.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID='" & Format(Cmb_Vendedor_Comision.ItemData(Cmb_Vendedor_Comision.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Empresa,No_Remision"
    Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Remisiones_Clientes.EOF Then
        With Rs_Consulta_Adm_Remisiones_Clientes
            Call Encabezado_Reportes("ESTADO DE CUENTA GENERAL", Format(Dtp_Comision_Fecha_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Comision_Fecha_Fin.Value, "dd/MMM/yyyy"), True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Print #1, "Empresa  Remision    Fecha          Importe            Pagos             Saldo       Cliente                    Vendedor "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Empresa" & Chr(9) & "No. Remision" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Cliente" & Chr(9) & "Vendedor" & Chr(9) & "Telefono"
            Grid_Reporte.AddItem ""
            While Not .EOF
                Cliente = .rdoColumns("Nombre")
                Vendedor = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Telefono")) Then
                    Telefono_Cliente = .rdoColumns("Telefono")
                Else
                    Telefono_Cliente = ""
                End If
                If .rdoColumns("Empresa") = 1 Then
                    Print #1, "  NHE  "; .rdoColumns("No_Remision"); _
                        Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                        Spc(2); Mid(Cliente, 1, 33); _
                        Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor
                    Grid_Reporte.AddItem "  NHE  " & Chr(9) & .rdoColumns("No_Remision") _
                        & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                        & Chr(9) & Cliente & Chr(9) & Vendedor
                Else
                     Print #1, "  PNC  "; .rdoColumns("No_Remision"); _
                        Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                        Spc(2); Mid(Cliente, 1, 33); _
                        Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor
                    Grid_Reporte.AddItem "  PNC  " & Chr(9) & .rdoColumns("No_Remision") _
                        & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                        & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                        & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente
                End If
                Total_Facturas = Total_Facturas + .rdoColumns("Total")
                Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                .MoveNext
            Wend
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Adm_Remisiones_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Anticipos_Proveedores
    'DESCRIPCIÓN: Muestra el reporte de los anticipos de proveedores
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 26-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Anticipos_Proveedores()
Dim Mi_SQL As String
Dim Rs_Consulta_Anticipos As rdoResultset
Dim Rs_Consulta_Movimientos As rdoResultset
Dim Comentarios_Presupuestos As String
Dim Cantidad_Cheque As Double
Dim Proveedor_ID As String
Dim Respuesta As String
Dim Estatus_Anticipo As String

    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    If MsgBox("¿Desea ver el reporte detallado en aplicaciones?, de clic en NO para verlo general", vbQuestion + vbYesNo) = vbYes Then Respuesta = "SI"
    'Busca los anticipos del proveedor
    Mi_SQL = "SELECT Adm_Anticipos_Proveedores.*,Cat_Proveedores.Nombre AS Proveedor,Cat_Proveedores.Direccion,Cat_Proveedores.Colonia,Cat_Proveedores.CP,Cat_Proveedores.Ciudad_Estado,Cat_Bancos.Nombre AS Banco"
    Mi_SQL = Mi_SQL & " FROM Adm_Anticipos_Proveedores,Cat_Proveedores,Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Adm_Anticipos_Proveedores.Proveedor_ID=Cat_Proveedores.Proveedor_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Banco_ID=Cat_Bancos.Banco_ID"
    If Cmb_Cliente_Anticipo.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Proveedor_ID='" & Format(Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.ListIndex), "00000") & "'"
    End If
    If Cmb_Empresa_Anticipo.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Empresa=" & Cmb_Empresa_Anticipo.ListIndex
    End If
    If Cmb_Estatus_Anticipo.ListIndex > 0 Then
        If Cmb_Estatus_Anticipo.Text = "APLICADO" Then
            Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Aplicado='S'"
        Else
            Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Aplicado='N'"
        End If
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Anticipos_Proveedores.Fecha BETWEEN '" & Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Anticipos_Proveedores.Proveedor_ID"
    Set Rs_Consulta_Anticipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Anticipos.EOF Then
        Call Encabezado_Reportes("REPORTE DE ANTICIPOS DE PROVEEDORES", Format(Dtp_Fecha_Inicio_Anticipo_Cliente.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Anticipo_Cliente.Value, "dd/MMM/yyyy"), True)
        While Not Rs_Consulta_Anticipos.EOF
            If Proveedor_ID <> Rs_Consulta_Anticipos.rdoColumns("Proveedor_ID") Then
                Print #1, ""
                Print #1, "________________________________________________________________________________________________________________________"
                Print #1, Spc(2); "Proveedor: "; Rs_Consulta_Anticipos.rdoColumns("Proveedor")
                Print #1, Spc(2); "Direccion: "; Rs_Consulta_Anticipos.rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; Rs_Consulta_Anticipos.rdoColumns("Colonia"); Spc(2); "CP: "; Rs_Consulta_Anticipos.rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Rs_Consulta_Anticipos.rdoColumns("Ciudad_Estado")
                Print #1,
                Print #1, "  Fecha   Anticipo     Banco          F.Pago Referencia     Concepto                      Cantidad     Saldo   Estatus  "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Proveedor:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Proveedor")
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Ciudad_Estado")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Fecha" & Chr(9) & "Anticipo" & Chr(9) & "Banco" & Chr(9) & "Forma Pago" & Chr(9) & "Concepto" & Chr(9) & "Cantidad" & Chr(9) & "Saldo" & Chr(9) & "Estatus"
                Grid_Reporte.AddItem ""
                Proveedor_ID = Rs_Consulta_Anticipos.rdoColumns("Proveedor_ID")
            End If
            If Respuesta = "SI" Then
                'Busca si esta en los detalles de anticipos los movimientos
                Mi_SQL = "SELECT Adm_Detalles_Anticipos_Proveedores.*,Adm_Facturas_Proveedores.*"
                Mi_SQL = Mi_SQL & " FROM Adm_Detalles_Anticipos_Proveedores,Adm_Facturas_Proveedores"
                Mi_SQL = Mi_SQL & " WHERE Adm_Detalles_Anticipos_Proveedores.Proveedor_ID=Adm_Facturas_Proveedores.Proveedor_ID"
                Mi_SQL = Mi_SQL & " AND Adm_Detalles_Anticipos_Proveedores.No_Factura=Adm_Facturas_Proveedores.No_Factura"
                Mi_SQL = Mi_SQL & " AND Adm_Detalles_Anticipos_Proveedores.No_Anticipo=" & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo")
                Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Movimientos.EOF Then
                    While Not Rs_Consulta_Movimientos.EOF
                        If Rs_Consulta_Anticipos.rdoColumns("Aplicado") = "S" Then
                            Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                                Spc(2); Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18); _
                                Spc(20 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                                Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                                Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26); _
                                Spc(28 - Len(Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 12); _
                                Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                                Spc(1); "APLICADO"
                            Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                                & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                                & Chr(9) & "FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")) _
                                & Chr(9) & Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00") _
                                & Chr(9) & "0.00" _
                                & Chr(9) & "APLICADO"
                        Else
                            Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                                Spc(2); Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Banco")), 1, 18); _
                                Spc(20 - Len(Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Banco")), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                                Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                                Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26); _
                                Spc(28 - Len(Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 12); _
                                Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00"), 12); _
                                Spc(1); "PENDIENTE"
                            Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                                & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                                & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                                & Chr(9) & "FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")) _
                                & Chr(9) & Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00") _
                                & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00") _
                                & Chr(9) & "PENDIENTE"
                        End If
                        Rs_Consulta_Movimientos.MoveNext
                    Wend
                Else
                    Rs_Consulta_Movimientos.Close
                    'Consulta los movimientos generados por el cheque consultado
                    Mi_SQL = "SELECT No_Movimiento,Proveedor_Cliente,Estatus,Fecha,Forma_Pago,No_Cheque,Referencia,No_Factura,Empresa,Cantidad,Banco,Concepto,Cuenta,Beneficiario"
                    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
                    Mi_SQL = Mi_SQL & " WHERE Tipo='E'"
                    Mi_SQL = Mi_SQL & " AND Proveedor_Cliente='" & Rs_Consulta_Anticipos.rdoColumns("Proveedor_ID") & "'"
                    Mi_SQL = Mi_SQL & " AND Banco_ID='" & Rs_Consulta_Anticipos.rdoColumns("Banco_ID") & "'"
                    Mi_SQL = Mi_SQL & " AND Forma_Pago='" & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") & "'"
                    If Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") = "CHEQUE" Then
                        Mi_SQL = Mi_SQL & " AND No_Cheque='" & Rs_Consulta_Anticipos.rdoColumns("Referencia") & "'"
                    Else
                        Mi_SQL = Mi_SQL & " AND Referencia='" & Rs_Consulta_Anticipos.rdoColumns("Referencia") & "'"
                    End If
                    Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Movimientos.EOF Then
                        While Not Rs_Consulta_Movimientos.EOF
                            If Not IsNull(Rs_Consulta_Movimientos.rdoColumns("No_Factura")) Then
                                Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                                    Spc(2); Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18); _
                                    Spc(20 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                                    Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                                    Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26); _
                                    Spc(28 - Len(Mid("FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")), 1, 26))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 12); _
                                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12); _
                                    Spc(1); "APLICADO"
                                Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                                    & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                                    & Chr(9) & "FACTURA " & Trim(Rs_Consulta_Movimientos.rdoColumns("No_Factura")) _
                                    & Chr(9) & Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00") _
                                    & Chr(9) & "0.00" _
                                    & Chr(9) & "APLICADO"
                            Else
                                Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                                    Spc(2); Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18); _
                                    Spc(20 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                                    Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                                    Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid(Trim(Rs_Consulta_Movimientos.rdoColumns("Concepto")), 1, 26); _
                                    Spc(28 - Len(Mid(Trim(Rs_Consulta_Movimientos.rdoColumns("Concepto")), 1, 26))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 12); _
                                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00"), 12); _
                                    Spc(1); "PENDIENTE"
                                Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                                    & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                                    & Chr(9) & Trim(Rs_Consulta_Movimientos.rdoColumns("Concepto")) _
                                    & Chr(9) & Format(Rs_Consulta_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00") _
                                    & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00") _
                                    & Chr(9) & "PENDIENTE"
                            End If
                            Rs_Consulta_Movimientos.MoveNext
                        Wend
                    End If
                End If
                Rs_Consulta_Movimientos.Close
            Else
                If Val(Rs_Consulta_Anticipos.rdoColumns("Total")) > Val(Rs_Consulta_Anticipos.rdoColumns("Pago")) Then
                    Estatus_Anticipo = "PENDIENTE"
                Else
                    Estatus_Anticipo = "APLICADO"
                End If
                Print #1, Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Rs_Consulta_Anticipos.rdoColumns("No_Anticipo"), 6); _
                    Spc(2); Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18); _
                    Spc(20 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Banco"), 1, 18))); Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6); _
                    Spc(7 - Len(Mid(Rs_Consulta_Anticipos.rdoColumns("Forma_Pago"), 1, 6))); Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")); _
                    Spc(12 - Len(Rs_Consulta_Anticipos.rdoColumns("Referencia"))); Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")), 1, 26); _
                    Spc(28 - Len(Mid(Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")), 1, 26))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total"), "#,###,##0.00"), 12); _
                    Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00"), 12); _
                    Spc(1); Estatus_Anticipo
                Grid_Reporte.AddItem Format(Rs_Consulta_Anticipos.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("No_Anticipo") _
                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Banco") _
                    & Chr(9) & Rs_Consulta_Anticipos.rdoColumns("Forma_Pago") _
                    & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Referencia")) _
                    & Chr(9) & Trim(Rs_Consulta_Anticipos.rdoColumns("Concepto")) _
                    & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total"), "#,###,##0.00") _
                    & Chr(9) & Format(Rs_Consulta_Anticipos.rdoColumns("Total") - Rs_Consulta_Anticipos.rdoColumns("Pago"), "#,###,##0.00") _
                    & Chr(9) & Estatus_Anticipo
            End If
            Rs_Consulta_Anticipos.MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No se encontraron registros con la búsqueda especificada", vbExclamation
    End If
    Rs_Consulta_Anticipos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Gastos
    'DESCRIPCIÓN: Muestra el reporte de gastos del cheque seleccionado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 23-Agosto-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Gastos()
Dim Mi_SQL As String
Dim Rs_Consulta_Movimientos As rdoResultset
Dim Rs_Consulta_Detalles_Movimientos As rdoResultset
Dim Comentarios_Presupuestos As String
Dim Cantidad_Cheque As Double

    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    'Consulta los movimientos generados por el cheque consultado
    Mi_SQL = "SELECT No_Movimiento,Proveedor_Cliente,Estatus,Fecha,No_Cheque,Referencia,No_Factura,Empresa,Cantidad,Banco,Concepto,Cuenta,Beneficiario"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Tipo='E'"
    Mi_SQL = Mi_SQL & " AND Banco_ID='" & Format(Cmb_Banco.ItemData(Cmb_Banco.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND Forma_Pago='" & Cmb_Tipo_Pago.Text & "'"
    If Cmb_Tipo_Pago.ListIndex = 0 Then
        Mi_SQL = Mi_SQL & " AND No_Cheque='" & Format(Txt_No_Cheque.Text, "0000000000") & "'"
    Else
        Mi_SQL = Mi_SQL & " AND Referencia='" & Txt_No_Cheque.Text & "'"
    End If
    Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Movimientos.EOF Then
        With Rs_Consulta_Movimientos
            Call Encabezado_Reportes("REPORTE DE GASTOS", Format(Now, "dd/MMM/yyyy"), "", True)
            Print #1, "Tipo         : "; Cmb_Tipo_Pago.Text
            If Cmb_Tipo_Pago.ListIndex = 0 Then
                Print #1, "No. Cheque   : "; .rdoColumns("No_Cheque")
                Grid_Reporte.AddItem "No. Cheque: " & Chr(9) & .rdoColumns("No_Cheque")
            Else
                Print #1, "Referencia   : "; .rdoColumns("Referencia")
                Grid_Reporte.AddItem "Referencia: " & Chr(9) & .rdoColumns("Referencia")
            End If
            Print #1, "Banco        : "; .rdoColumns("Cuenta") & " - " & .rdoColumns("Banco")
            Print #1, "Beneficiario : "; .rdoColumns("Beneficiario")
            If .rdoColumns("Estatus") = "A" Then
                Print #1, "Estatus      : ACTIVO"
            Else
                Print #1, "Estatus      : CANCELADO"
            End If
            Print #1,
            Print #1, "Movimientos  :"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "   Cuenta     No. Factura             Proveedor                                 Tipo Gasto            Cantidad           "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem "Tipo: " & Chr(9) & Cmb_Tipo_Pago.Text
            Grid_Reporte.AddItem "Banco: " & Chr(9) & .rdoColumns("Cuenta") & " - " & .rdoColumns("Banco")
            Grid_Reporte.AddItem "Beneficiario: " & Chr(9) & .rdoColumns("Beneficiario")
            If .rdoColumns("Estatus") = "A" Then
                Grid_Reporte.AddItem "Estatus:" & Chr(9) & "ACTIVO"
            Else
                Grid_Reporte.AddItem "Estatus:" & Chr(9) & "CANCELADO"
            End If
            Grid_Reporte.AddItem "Movimientos"
            Grid_Reporte.AddItem "Cuenta" & Chr(9) & "No. Factura" & Chr(9) & "Proveedor" & Chr(9) & "Tipo de Gasto" & Chr(9) & "Cantidad"
            Grid_Reporte.AddItem ""
            While Not .EOF
                Cantidad_Cheque = Cantidad_Cheque + Val(.rdoColumns("Cantidad"))
                'Consulta si tiene pagos en caja chica para mostrar los detalles
                Mi_SQL = "SELECT Detalles_Corte_Caja_Chica.*"
                Mi_SQL = Mi_SQL & " FROM Adm_Corte_Caja_Chica,Detalles_Corte_Caja_Chica"
                Mi_SQL = Mi_SQL & " WHERE Adm_Corte_Caja_Chica.No_Corte=Detalles_Corte_Caja_Chica.No_Corte"
                Mi_SQL = Mi_SQL & " AND Adm_Corte_Caja_Chica.No_Movimiento='" & .rdoColumns("No_Movimiento") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Corte_Caja_Chica.No_Cheque='" & .rdoColumns("No_Cheque") & "'"
                Set Rs_Consulta_Detalles_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Detalles_Movimientos.EOF Then
                    While Not Rs_Consulta_Detalles_Movimientos.EOF
                        If Not IsNull(Rs_Consulta_Detalles_Movimientos.rdoColumns("Comentarios_Presupuestos")) Then
                            Comentarios_Presupuestos = Rs_Consulta_Detalles_Movimientos.rdoColumns("Comentarios_Presupuestos")
                        Else
                            Comentarios_Presupuestos = ""
                        End If
                        If Not IsNull(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal")) Then
                            Print #1, Spc(1); Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal"); _
                                Spc(13 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal"))); Rs_Consulta_Detalles_Movimientos.rdoColumns("Referencia"); _
                                Spc(23 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Referencia"))); Mid(Comentarios_Presupuestos, 1, 40); _
                                Spc(43 - Len(Mid(Comentarios_Presupuestos, 1, 40))); Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"); _
                                Spc(15 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                        Else
                            Print #1, Spc(14); Rs_Consulta_Detalles_Movimientos.rdoColumns("Referencia"); _
                                Spc(23 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Referencia"))); Mid(Comentarios_Presupuestos, 1, 40); _
                                Spc(43 - Len(Mid(Comentarios_Presupuestos, 1, 40))); Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"); _
                                Spc(15 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                        End If
                        Rs_Consulta_Detalles_Movimientos.MoveNext
                    Wend
                Else
                    If Not IsNull(.rdoColumns("Proveedor_Cliente")) Then
                        Rs_Consulta_Detalles_Movimientos.Close
                        'Busca datos de fatura del proveedor
                        Mi_SQL = "SELECT * FROM Adm_Facturas_Proveedores"
                        Mi_SQL = Mi_SQL & " WHERE No_Factura='" & .rdoColumns("No_Factura") & "'"
                        Mi_SQL = Mi_SQL & " AND Proveedor_ID='" & .rdoColumns("Proveedor_Cliente") & "'"
                        Set Rs_Consulta_Detalles_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Detalles_Movimientos.EOF Then
                            If Not IsNull(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal")) Then
                                Print #1, Spc(1); Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal"); _
                                    Spc(13 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal"))); Rs_Consulta_Detalles_Movimientos.rdoColumns("No_Factura"); _
                                    Spc(23 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("No_Factura"))); Mid(.rdoColumns("Beneficiario"), 1, 40); _
                                    Spc(43 - Len(Mid(.rdoColumns("Beneficiario"), 1, 40))); Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"); _
                                    Spc(15 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                            Else
                                Print #1, Spc(14); Rs_Consulta_Detalles_Movimientos.rdoColumns("No_Factura"); _
                                    Spc(23 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("No_Factura"))); Mid(.rdoColumns("Beneficiario"), 1, 40); _
                                    Spc(43 - Len(Mid(.rdoColumns("Beneficiario"), 1, 40))); Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"); _
                                    Spc(15 - Len(Rs_Consulta_Detalles_Movimientos.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                            End If
                        End If
                    Else
                        Rs_Consulta_Detalles_Movimientos.Close
                        'Busca datos de movimientos
                        Mi_SQL = "SELECT * FROM Adm_Movimientos_Presupuestos"
                        Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & .rdoColumns("No_Movimiento") & "'"
                        Set Rs_Consulta_Detalles_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Detalles_Movimientos.EOF Then
                            If Not IsNull(Rs_Consulta_Detalles_Movimientos.rdoColumns("Comentarios_Presupuesto")) Then
                                Comentarios_Presupuestos = Rs_Consulta_Detalles_Movimientos.rdoColumns("Comentarios_Presupuesto")
                            Else
                                Comentarios_Presupuestos = ""
                            End If
                            If Not IsNull(Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal")) Then
                                Print #1, Spc(1); Rs_Consulta_Detalles_Movimientos.rdoColumns("Cuenta_Presupuestal"); _
                                    Spc(36); Mid(Comentarios_Presupuestos, 1, 40); _
                                    Spc(58); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                            Else
                                Print #1, Spc(50); Mid(Comentarios_Presupuestos, 1, 40); _
                                    Spc(58); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###,##0.00"), 15)
                            End If
                        End If
                    End If
                End If
                Rs_Consulta_Detalles_Movimientos.Close
                .MoveNext
            Wend
            Print #1,
            Print #1, "Total Cheque : "; Format(Cantidad_Cheque, "#,###,##0.00")
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No existe el número de cheque seleccionado", vbInformation
    End If
    Rs_Consulta_Movimientos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cartera_Vencida
    'DESCRIPCIÓN: Consulta las facucturas vencidas
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 07-Abril-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Cartera_Vencida()
Dim Mi_SQL As String
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Vendedor As String
Dim Telefono_Cliente As String
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Total_Pagos As Double                               'Almacena el total de los pagos hechos en las facturas
Dim Total_Saldo As Double                               'Almacena el total del saldo de las facturas
Dim Total_Facturas As Double                            'Almacena el total de la factura
Dim Fecha_Vencimiento As Integer

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Factura,Adm_Facturas_Clientes.Cliente_ID,Empresa,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Abono,Saldo"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Nombre,Cat_Clientes.Vendedor_ID,Cat_Clientes.Telefono,Cat_Vendedores.Clave"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes,Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    If Chk_General.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    End If
    Mi_SQL = Mi_SQL & " AND Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Pagada='N'"
    Mi_SQL = Mi_SQL & " ORDER BY Empresa,No_Factura"
    Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
        With Rs_Consulta_Adm_Facturas_Clientes
            If Chk_General.Value = 0 Then
                Call Encabezado_Reportes("CARTERA VENCIDA", Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "dd/MMM/yyyy"), True)
            Else
                Call Encabezado_Reportes("CARTERA VENCIDA", Format(Now, "dd/MMM/yyyy"), "", True)
            End If
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "Empresa  Factura     Fecha         Importe           Pagos            Saldo      Cliente                    Vendedor Dias"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Empresa" & Chr(9) & "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Cliente" & Chr(9) & "Vendedor" & Chr(9) & "Telefono" & Chr(9) & "Dias Vencida"
            Grid_Reporte.AddItem ""
            While Not .EOF
                Cliente = .rdoColumns("Nombre")
                Vendedor = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Telefono")) Then
                    Telefono_Cliente = .rdoColumns("Telefono")
                Else
                    Telefono_Cliente = ""
                End If
                Fecha_Vencimiento = DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                If Fecha_Vencimiento > 30 Then
                    'Valida las facturas con 30 días de que se generó para ver si ya está vencida
                    If .rdoColumns("Empresa") = 1 Then
                        Print #1, "  NHE  "; .rdoColumns("No_Factura"); _
                            Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                            Spc(1); Mid(Cliente, 1, 33); _
                            Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor; Spc(5 - Len(Vendedor)); Val(Fecha_Vencimiento - 30)
                        Grid_Reporte.AddItem "  NHE  " & Chr(9) & .rdoColumns("No_Factura") _
                            & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                            & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente & Chr(9) & Fecha_Vencimiento - 30
                    Else
                         Print #1, "  PNC  "; .rdoColumns("No_Factura"); _
                            Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                            Spc(1); Mid(Cliente, 1, 33); _
                            Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor; Spc(5 - Len(Vendedor)); Val(Fecha_Vencimiento - 30)
                        Grid_Reporte.AddItem "  PNC  " & Chr(9) & .rdoColumns("No_Factura") _
                            & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                            & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente & Chr(9) & Fecha_Vencimiento - 30
                    End If
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                    Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                End If
                .MoveNext
            Wend
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    End If
    Rs_Consulta_Adm_Facturas_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cartera_Vencida_Remisiones
    'DESCRIPCIÓN: Consulta las remisiones vencidas
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Cartera_Vencida_Remisiones()
Dim Mi_SQL As String
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Vendedor As String
Dim Telefono_Cliente As String
Dim Rs_Consulta_Adm_Remisiones_Clientes As rdoResultset   'Manejo de registro
Dim Total_Pagos As Double                               'Almacena el total de los pagos hechos en las facturas
Dim Total_Saldo As Double                               'Almacena el total del saldo de las facturas
Dim Total_Facturas As Double                            'Almacena el total de la factura
Dim Fecha_Vencimiento As Integer

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Remision,Adm_Remisiones_Clientes.Cliente_ID,Empresa,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada,Abono,Saldo"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Nombre,Cat_Clientes.Vendedor_ID,Cat_Clientes.Telefono,Cat_Vendedores.Clave"
    Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes,Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    If Chk_General.Value = 0 Then
        Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    End If
    Mi_SQL = Mi_SQL & " AND Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Pagada='N'"
    Mi_SQL = Mi_SQL & " ORDER BY Empresa,No_Remision"
    Set Rs_Consulta_Adm_Remisiones_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Adm_Remisiones_Clientes.EOF Then
        With Rs_Consulta_Adm_Remisiones_Clientes
            If Chk_General.Value = 0 Then
                Call Encabezado_Reportes("CARTERA VENCIDA REMISIONES", Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "dd/MMM/yyyy"), True)
            Else
                Call Encabezado_Reportes("CARTERA VENCIDA REMISIONES", Format(Now, "dd/MMM/yyyy"), "", True)
            End If
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "Empresa  Remision    Fecha         Importe           Pagos            Saldo      Cliente                    Vendedor Dias"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Empresa" & Chr(9) & "No. Remision" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Cliente" & Chr(9) & "Vendedor" & Chr(9) & "Telefono" & Chr(9) & "Dias Vencida"
            Grid_Reporte.AddItem ""
            While Not .EOF
                Cliente = .rdoColumns("Nombre")
                Vendedor = .rdoColumns("Clave")
                If Not IsNull(.rdoColumns("Telefono")) Then
                    Telefono_Cliente = .rdoColumns("Telefono")
                Else
                    Telefono_Cliente = ""
                End If
                Fecha_Vencimiento = DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                If Fecha_Vencimiento > 30 Then
                    'Valida las facturas con 30 días de que se generó para ver si ya está vencida
                    If .rdoColumns("Empresa") = 1 Then
                        Print #1, "  NHE  "; .rdoColumns("No_Remision"); _
                            Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                            Spc(1); Mid(Cliente, 1, 33); _
                            Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor; Spc(5 - Len(Vendedor)); Val(Fecha_Vencimiento - 30)
                        Grid_Reporte.AddItem "  NHE  " & Chr(9) & .rdoColumns("No_Remision") _
                            & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                            & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente & Chr(9) & Fecha_Vencimiento - 30
                    Else
                         Print #1, "  PNC  "; .rdoColumns("No_Remision"); _
                            Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                            Spc(1); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                            Spc(1); Mid(Cliente, 1, 33); _
                            Spc(35 - Len(Mid(Cliente, 1, 33))); Vendedor; Spc(5 - Len(Vendedor)); Val(Fecha_Vencimiento - 30)
                        Grid_Reporte.AddItem "  PNC  " & Chr(9) & .rdoColumns("No_Remision") _
                            & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                            & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                            & Chr(9) & Cliente & Chr(9) & Vendedor & Chr(9) & Telefono_Cliente & Chr(9) & Fecha_Vencimiento - 30
                    End If
                    Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                    Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                End If
                .MoveNext
            Wend
            Print #1, "_________________________________________________________________________________________________________________________"
            Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No hay remisiones en las fechas que proporcionó", vbExclamation
    End If
    Rs_Consulta_Adm_Remisiones_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Efectividad
    'DESCRIPCIÓN: De acuerdo a parámetros capturados, evalua la efectividad de
    '             cobranza en el periodo seleccionado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 01-Julio-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Reporte_Efectividad()
Dim Mi_SQL As String
Dim Rs_Facturas_Clientes As rdoResultset
    
    Grid_Reporte.Cols = 12
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT SUM(Total) AS Total_Facturas,SUM(Abono) AS Abono_Facturas,SUM(Saldo) AS Saldo_Facturas"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND MONTH(Fecha)='" & Val(Format(Dtp_Mes_Evaluacion.Value, "MM")) & "'"
    Mi_SQL = Mi_SQL & " AND YEAR(Fecha)='" & Val(Format(Dtp_Mes_Evaluacion.Value, "yyyy")) & "'"
    Set Rs_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Facturas_Clientes
        If Not IsNull(.rdoColumns("Total_Facturas")) Then
            Call Encabezado_Reportes("EFECTIVIDAD DE COBRANZA MENSUAL", Format(Now, "dd/MMM/yyyy"), "", True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " Mes Evaluación     Total por Cobrar     Recuperado       % Al día     Ob. 1er mes      Falta 1er objetivo"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Mes Evaluacion" & Chr(9) & "Total por Cobrar" & Chr(9) & "Recuperado" & Chr(9) & "% Al día" & Chr(9) & "Ob. 1er mes" & Chr(9) & "Falta 1er ob." & Chr(9) & "Ob. 2do mes" & Chr(9) & "Falta 2do ob." & Chr(9) & "Ob. 3er mes" & Chr(9) & "Falta 3er ob." & Chr(9) & "Ob. 4to mes" & Chr(9) & "Falta 4to ob."
            Grid_Reporte.AddItem ""
            Print #1, Spc(1); Format(Dtp_Mes_Evaluacion.Value, "MMMM yyyy"); _
                Spc(18 - Len(Format(Dtp_Mes_Evaluacion.Value, "MMMM yyyy"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total_Facturas"), "#,###,##0.00"), 15); _
                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono_Facturas"), "#,###,##0.00"), 15); _
                Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono_Facturas") / .rdoColumns("Total_Facturas") * 100, "##0.00"), 6); _
                Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Txt_Objetivo_Efectividad_1.Text, "##0.00"), 6); _
                Spc(9); Conectar_Ayudante.Alinea_Derecha_Derecha(Format((Val(Txt_Objetivo_Efectividad_1.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "#,###,##0.00"), 15)
            Print #1,
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " Ob. 2do mes      Falta 2do objetivo     Ob. 3er mes      Falta 3er objetivo     Ob. 4to mes      Falta 4to objetivo"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Txt_Objetivo_Efectividad_2.Text, "##0.00"), 6); _
                Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format((Val(Txt_Objetivo_Efectividad_2.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "#,###,##0.00"), 15); _
                Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Txt_Objetivo_Efectividad_3.Text, "##0.00"), 6); _
                Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format((Val(Txt_Objetivo_Efectividad_3.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "#,###,##0.00"), 15); _
                Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Txt_Objetivo_Efectividad_4.Text, "##0.00"), 6); _
                Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format((Val(Txt_Objetivo_Efectividad_4.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "#,###,##0.00"), 15)
            Grid_Reporte.AddItem Format(Dtp_Mes_Evaluacion.Value, "MMMM yyyy") _
                & Chr(9) & Format(.rdoColumns("Total_Facturas"), "#,###,##0.00") _
                & Chr(9) & Format(.rdoColumns("Abono_Facturas"), "#,###,##0.00") _
                & Chr(9) & Format(.rdoColumns("Abono_Facturas") / .rdoColumns("Total_Facturas") * 100, "##0.00") _
                & Chr(9) & Format(Txt_Objetivo_Efectividad_1.Text, "##0.00") _
                & Chr(9) & Format((Val(Txt_Objetivo_Efectividad_1.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "##0.00") _
                & Chr(9) & Format(Txt_Objetivo_Efectividad_2.Text, "##0.00") _
                & Chr(9) & Format((Val(Txt_Objetivo_Efectividad_2.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "##0.00") _
                & Chr(9) & Format(Txt_Objetivo_Efectividad_3.Text, "##0.00") _
                & Chr(9) & Format((Val(Txt_Objetivo_Efectividad_3.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "##0.00") _
                & Chr(9) & Format(Txt_Objetivo_Efectividad_4.Text, "##0.00") _
                & Chr(9) & Format((Val(Txt_Objetivo_Efectividad_4.Text) * .rdoColumns("Total_Facturas") / 100) - (.rdoColumns("Abono_Facturas")), "##0.00")
            Call Encabezado_Reportes("", "", "", False)
        Else
            MsgBox "No hay ventas en el periodo seleccionado", vbExclamation
        End If
    End With
    Rs_Facturas_Clientes.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: General_Presupuestos
'DESCRIPCION: Muestra el estatus de los presupuestos y el total devengado y pagado
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 30-Mayo-2008
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub General_Presupuestos()
Dim Mi_SQL As String
Dim Rs_Consulta_Presupuestos As rdoResultset
Dim Rs_Consulta_Facturas As rdoResultset
Dim Total As Double
Dim Total_Devengado As Double
Dim Total_Pagado As Double
Dim Total_Presupuesto As Double
Dim Devengado_Presupuesto As Double
Dim Pagado_Presupuesto As Double

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    Call Encabezado_Reportes("REPORTE GENERAL DE PRESUPUESTOS", Format(Now, "dd/MMM/yyyy"), "", True)
    'Consulta los datos del presupuesto
    Mi_SQL = "SELECT No_Presupuesto,Año,Titulo_Proyecto,Total,Abono,Saldo,Cat_Areas.Nombre,Tipo_Presupuesto"
    Mi_SQL = Mi_SQL & " FROM Ope_Presupuestos,Cat_Areas"
    Mi_SQL = Mi_SQL & " WHERE Ope_Presupuestos.Area_ID=Cat_Areas.Area_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Presupuestos.Estatus='AUTORIZADO'"
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Presupuestos.No_Presupuesto"
    Set Rs_Consulta_Presupuestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Presupuestos.EOF Then
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, " Presupuesto      Area          Tipo             Titulo                         Total        Devengado         Pagado   "
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No_Presupuesto" & Chr(9) & "Area" & Chr(9) & "Tipo" & Chr(9) & "Titulo" & Chr(9) & "Total" & Chr(9) & "Devengado" & Chr(9) & "Pagado"
        Grid_Reporte.AddItem ""
        With Rs_Consulta_Presupuestos
            While Not Rs_Consulta_Presupuestos.EOF
                Devengado_Presupuesto = 0
                Pagado_Presupuesto = 0
                Total = Total + .rdoColumns("Total")
                'Consulta los saldos y abonos de las facturas del presupuesto
                Mi_SQL = "SELECT Saldo,Abono,Moneda,Tipo_Cambio FROM Adm_Facturas_Proveedores"
                Mi_SQL = Mi_SQL & " WHERE No_Presupuesto='" & .rdoColumns("No_Presupuesto") & "'"
                Mi_SQL = Mi_SQL & " AND Cancelada='N'"
                Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                While Not Rs_Consulta_Facturas.EOF
                    If UCase(Trim(Rs_Consulta_Facturas.rdoColumns("Moneda"))) = "PESOS" Then
                        Devengado_Presupuesto = Devengado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Saldo")
                        Total_Devengado = Total_Devengado + Rs_Consulta_Facturas.rdoColumns("Saldo")
                        Pagado_Presupuesto = Pagado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Abono")
                        Total_Pagado = Total_Pagado + Rs_Consulta_Facturas.rdoColumns("Abono")
                    Else
                        Devengado_Presupuesto = Devengado_Presupuesto + (Rs_Consulta_Facturas.rdoColumns("Saldo") * Rs_Consulta_Facturas.rdoColumns("Tipo_Cambio"))
                        Total_Devengado = Total_Devengado + (Rs_Consulta_Facturas.rdoColumns("Saldo") * Rs_Consulta_Facturas.rdoColumns("Tipo_Cambio"))
                        Pagado_Presupuesto = Pagado_Presupuesto + (Rs_Consulta_Facturas.rdoColumns("Abono") * Rs_Consulta_Facturas.rdoColumns("Tipo_Cambio"))
                        Total_Pagado = Total_Pagado + (Rs_Consulta_Facturas.rdoColumns("Abono") * Rs_Consulta_Facturas.rdoColumns("Tipo_Cambio"))
                    End If
                    Rs_Consulta_Facturas.MoveNext
                Wend
                Rs_Consulta_Facturas.Close
                'Consulta los saldos y abonos de las facturas prorrateadas del presupuesto
                Mi_SQL = "SELECT SUM(Ope_Presupuestos_Prorrateo.Importe) AS Devengado"
                Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores,Ope_Presupuestos_Prorrateo"
                Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Proveedores.No_Factura=Ope_Presupuestos_Prorrateo.No_Factura"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Proveedor_ID=Ope_Presupuestos_Prorrateo.Proveedor_ID"
                Mi_SQL = Mi_SQL & " AND Ope_Presupuestos_Prorrateo.No_Presupuesto='" & .rdoColumns("No_Presupuesto") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Pagada='N'"
                Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Facturas.EOF Then
                    If Not IsNull(Rs_Consulta_Facturas.rdoColumns("Devengado")) Then
                        Devengado_Presupuesto = Devengado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Devengado")
                        Total_Devengado = Total_Devengado + Rs_Consulta_Facturas.rdoColumns("Devengado")
                    End If
                End If
                Rs_Consulta_Facturas.Close
                'Consulta los saldos y abonos de las facturas prorrateadas del presupuesto
                Mi_SQL = "SELECT SUM(Ope_Presupuestos_Prorrateo.Importe) AS Pagado"
                Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores,Ope_Presupuestos_Prorrateo"
                Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Proveedores.No_Factura=Ope_Presupuestos_Prorrateo.No_Factura"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Proveedor_ID=Ope_Presupuestos_Prorrateo.Proveedor_ID"
                Mi_SQL = Mi_SQL & " AND Ope_Presupuestos_Prorrateo.No_Presupuesto='" & .rdoColumns("No_Presupuesto") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND Adm_Facturas_Proveedores.Pagada='S'"
                Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Facturas.EOF Then
                    If Not IsNull(Rs_Consulta_Facturas.rdoColumns("Pagado")) Then
                        Pagado_Presupuesto = Pagado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Pagado")
                        Total_Pagado = Total_Pagado + Rs_Consulta_Facturas.rdoColumns("Pagado")
                    End If
                End If
                Rs_Consulta_Facturas.Close
                'Consulta las facturas del presupuesto seleccionado pagadas con caja chica
                Mi_SQL = "SELECT SUM(Detalles_Corte_Caja_Chica.Cantidad) AS Pagado"
                Mi_SQL = Mi_SQL & " FROM Adm_Corte_Caja_Chica,Detalles_Corte_Caja_Chica"
                Mi_SQL = Mi_SQL & " WHERE Adm_Corte_Caja_Chica.No_Corte=Detalles_Corte_Caja_Chica.No_Corte"
                Mi_SQL = Mi_SQL & " AND Detalles_Corte_Caja_Chica.Proveedor_ID IS NULL"
                Mi_SQL = Mi_SQL & " AND Detalles_Corte_Caja_Chica.No_Presupuesto='" & .rdoColumns("No_Presupuesto") & "'"
                Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Facturas.EOF Then
                    If Not IsNull(Rs_Consulta_Facturas.rdoColumns("Pagado")) Then
                        Pagado_Presupuesto = Pagado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Pagado")
                        Total_Pagado = Total_Pagado + Rs_Consulta_Facturas.rdoColumns("Pagado")
                    End If
                End If
                Rs_Consulta_Facturas.Close
                'Consulta los gastos del presupuesto seleccionado
                Mi_SQL = "SELECT SUM(Adm_Movimientos_Presupuestos.Cantidad) AS Pagado"
                Mi_SQL = Mi_SQL & " FROM Adm_Movimientos,Adm_Movimientos_Presupuestos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos.No_Movimiento=Adm_Movimientos_Presupuestos.No_Movimiento"
                Mi_SQL = Mi_SQL & " AND Adm_Movimientos_Presupuestos.No_Presupuesto='" & .rdoColumns("No_Presupuesto") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Tipo='E'"
                Mi_SQL = Mi_SQL & " AND Adm_Movimientos.Proveedor_Cliente IS NULL"
                Mi_SQL = Mi_SQL & " AND Adm_Movimientos.No_Factura IS NULL"
                Set Rs_Consulta_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Facturas.EOF Then
                    If Not IsNull(Rs_Consulta_Facturas.rdoColumns("Pagado")) Then
                        Pagado_Presupuesto = Pagado_Presupuesto + Rs_Consulta_Facturas.rdoColumns("Pagado")
                        Total_Pagado = Total_Pagado + Rs_Consulta_Facturas.rdoColumns("Pagado")
                    End If
                End If
                Rs_Consulta_Facturas.Close
                Print #1, Spc(1); .rdoColumns("No_Presupuesto"); Spc(2); Mid(.rdoColumns("Nombre"), 1, 15); _
                    Spc(17 - Len(Mid(.rdoColumns("Nombre"), 1, 15))); Trim(.rdoColumns("Tipo_Presupuesto")); _
                    Spc(11 - Len(Trim(.rdoColumns("Tipo_Presupuesto")))); Mid(Trim(.rdoColumns("Titulo_Proyecto")), 1, 30); _
                    Spc(32 - Len(Mid(Trim(.rdoColumns("Titulo_Proyecto")), 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Devengado_Presupuesto, "#,###,##0.00"), 14); _
                    Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Pagado_Presupuesto, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem .rdoColumns("No_Presupuesto") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Tipo_Presupuesto") & Chr(9) & _
                    .rdoColumns("Titulo_Proyecto") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Format(Devengado_Presupuesto, "#,###,##0.00") & Chr(9) & Format(Pagado_Presupuesto, "#,###,##0.00")
                Rs_Consulta_Presupuestos.MoveNext
            Wend
        End With
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(50); " TOTALES:    "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Devengado, "#,###,##0.00"), 14); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagado, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Total, "#,###,##0.00") & Chr(9) & Format(Total_Devengado, "#,###,##0.00") & Chr(9) & Format(Total_Pagado, "#,###,##0.00")
    End If
    Call Encabezado_Reportes("", "", "", False)
    Rs_Consulta_Presupuestos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: General_Saldos
    'DESCRIPCIÓN: Consulta todos los saldos agrupados por vendedor y por cliente
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 20-Mayo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub General_Saldos()
Dim Mi_SQL As String
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Telefono_Cliente As String
Dim Rs_Consulta_Cat_Vendedores As rdoResultset
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Total_Pagos As Double                               'Almacena el total de los pagos hechos en las facturas
Dim Total_Saldo As Double                               'Almacena el total del saldo de las facturas
Dim Total_Facturas As Double                            'Almacena el total de la factura
Dim Fecha_Vencimiento As Integer
Dim Encabezado_Cliente As Boolean

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    'Consulta los vendedores
    Mi_SQL = "SELECT Vendedor_ID,Nombre FROM Cat_Vendedores ORDER BY Nombre"
    Set Rs_Consulta_Cat_Vendedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE GENERAL DE SALDOS", Format(Now, "dd/MMM/yyyy"), "", True)
    While Not Rs_Consulta_Cat_Vendedores.EOF
        Print #1, "_________________________________________________________________________________________________________________________"
        Print #1,
        Print #1, Conectar_Ayudante.Centra("VENDEDOR: " & Rs_Consulta_Cat_Vendedores.rdoColumns("Nombre"), Len("************************************************************************************************************************"))
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "VENDEDOR:" & Chr(9) & Rs_Consulta_Cat_Vendedores.rdoColumns("Nombre")
        Grid_Reporte.AddItem ""
        'Consulta los clientes del vendedor en turno
        Mi_SQL = "SELECT Cliente_ID,Cat_Clientes.Nombre AS Cliente,RFC,Direccion,Colonia,CP,Telefono,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre AS Estado"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes LEFT JOIN Cat_Ciudades ON Cat_Clientes.Ciudad_ID=Cat_Ciudades.Ciudad_ID"
        Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Estados ON Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
        Mi_SQL = Mi_SQL & " WHERE Vendedor_ID='" & Rs_Consulta_Cat_Vendedores.rdoColumns("Vendedor_ID") & "'"
        Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre"
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Cat_Clientes.EOF
            'Consulta las facturas que esten pendientes de pago del vendedor
            Mi_SQL = "SELECT No_Factura,Empresa,Fecha,Total,Abono,Saldo"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Pagada='N'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Adm_Facturas_Clientes
                If Not .EOF Then
                    Encabezado_Cliente = True
                    Print #1,
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, "CLIENTE  : "; Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"); Spc(72 - Len("CLIENTE  : " & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"))); "R.F.C.   : "; Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
                    Print #1, "DOMICILIO: "; Rs_Consulta_Cat_Clientes.rdoColumns("Direccion"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                    Print #1, "TELEFONO : "; Rs_Consulta_Cat_Clientes.rdoColumns("Telefono"); "   C.P.: "; Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                    If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad")) Then
                        Print #1, "CIUDAD   : "; Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Estado")
                    Else
                        Print #1, "CIUDAD   : "
                    End If
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " Factura  Empresa    Fecha            Importe           Pagos            Saldo       Dias"
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "CLIENTE:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente") & Chr(9) & Chr(9) & "RFC" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
                    Grid_Reporte.AddItem "DOMICILIO:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                    Grid_Reporte.AddItem "TELEFONO:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Telefono") & Chr(9) & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                    If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad")) Then
                        Grid_Reporte.AddItem "CIUDAD   : " & Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado")
                    Else
                        Grid_Reporte.AddItem "CIUDAD:"
                    End If
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Dias"
                    Grid_Reporte.AddItem ""
                    Total_Facturas = 0
                    Total_Pagos = 0
                    Total_Saldo = 0
                    While Not .EOF
                        Fecha_Vencimiento = DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                        If .rdoColumns("Empresa") = 1 Then
                            Print #1, .rdoColumns("No_Factura"); "  NHE  "; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(5); Fecha_Vencimiento
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & "  NHE  " _
                                & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                                & Chr(9) & Fecha_Vencimiento
                        Else
                             Print #1, .rdoColumns("No_Factura"); "  PNC  "; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(5); Fecha_Vencimiento
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & "  PNC  " _
                                & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                                & Chr(9) & Fecha_Vencimiento
                        End If
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                        Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                        .MoveNext
                    Wend
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
                End If
            End With
            Rs_Consulta_Adm_Facturas_Clientes.Close
            'Consulta las remisiones que esten pendientes de pago del vendedor
            Mi_SQL = "SELECT No_Remision,Empresa,Fecha,Total,Abono,Saldo"
            Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND Pagada='N'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Adm_Facturas_Clientes
                If Not .EOF Then
                    If Encabezado_Cliente = False Then
                        Print #1,
                        Print #1, "_________________________________________________________________________________________________________________________"
                        Print #1, "CLIENTE  : "; Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"); Spc(72 - Len("CLIENTE  : " & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente"))); "R.F.C.   : "; Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
                        Print #1, "DOMICILIO: "; Rs_Consulta_Cat_Clientes.rdoColumns("Direccion"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                        Print #1, "TELEFONO : "; Rs_Consulta_Cat_Clientes.rdoColumns("Telefono"); "   C.P.: "; Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                        If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad")) Then
                            Print #1, "CIUDAD   : "; Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad"); ", "; Rs_Consulta_Cat_Clientes.rdoColumns("Estado")
                        Else
                            Print #1, "CIUDAD   : "
                        End If
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "CLIENTE:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente") & Chr(9) & Chr(9) & "RFC" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
                        Grid_Reporte.AddItem "DOMICILIO:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                        Grid_Reporte.AddItem "TELEFONO:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Telefono") & Chr(9) & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("CP")
                        If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad")) Then
                            Grid_Reporte.AddItem "CIUDAD   : " & Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado")
                        Else
                            Grid_Reporte.AddItem "CIUDAD:"
                        End If
                    End If
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "Remision  Empresa    Fecha            Importe           Pagos            Saldo       Dias"
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Pagos" & Chr(9) & "Saldo" & Chr(9) & "Dias"
                    Grid_Reporte.AddItem ""
                    Total_Facturas = 0
                    Total_Pagos = 0
                    Total_Saldo = 0
                    While Not .EOF
                        Fecha_Vencimiento = DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
                        If .rdoColumns("Empresa") = 1 Then
                            Print #1, .rdoColumns("No_Remision"); "  NHE  "; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(5); Fecha_Vencimiento
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & "  NHE  " _
                                & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                                & Chr(9) & Fecha_Vencimiento
                        Else
                             Print #1, .rdoColumns("No_Remision"); "  PNC  "; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(5); Fecha_Vencimiento
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & "  PNC  " _
                                & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15) _
                                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15) _
                                & Chr(9) & Fecha_Vencimiento
                        End If
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Total_Pagos = Total_Pagos + .rdoColumns("Abono")
                        Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                        .MoveNext
                    Wend
                    Print #1, "_________________________________________________________________________________________________________________________"
                    Print #1, Spc(15); "T O T A L E S "; Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos, "#,###,##0.00"), 15); Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00") & Chr(9) & Format(Total_Pagos, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00")
                End If
            End With
            Rs_Consulta_Adm_Facturas_Clientes.Close
            Encabezado_Cliente = False
            Rs_Consulta_Cat_Clientes.MoveNext
        Wend
        Rs_Consulta_Cat_Clientes.Close
        Rs_Consulta_Cat_Vendedores.MoveNext
    Wend
    Call Encabezado_Reportes("", "", "", False)
    Rs_Consulta_Cat_Vendedores.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Clientes
    'DESCRIPCIÓN: Consulta los datos de los clientes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 21-Mayo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Reporte_Clientes()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Saldo_Clientes As rdoResultset
Dim Saldo_Cliente As Double
Dim Total_Compras As Double
Dim Tuvo_Movimientos As String
Dim Año_Ventas As Integer
Dim Columnas As Integer

    MDIFrm_Apl_Principal.MousePointer = 11
    Grid_Reporte.Cols = 13 + (Year(Dtp_Año_Auditoria.Value) - 2007)
    Grid_Reporte.Rows = 0
    'Consulta los clientes del vendedor en turno
    Mi_SQL = "SELECT Cliente_ID,Cat_Clientes.Nombre AS Cliente,Cat_Clientes.RFC,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Telefono,Cat_Clientes.Estatus,Cat_Clientes.Tipo_Precio,Cat_Clientes.Dias_Credito,Cat_Clientes.Limite_Credito"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre AS Estado,Cat_Clientes.Zona_ID,Cat_Zonas.Nombre AS Zona"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Vendedor_ID,Cat_Vendedores.Nombre AS Vendedor"
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes LEFT JOIN Cat_Zonas ON Cat_Clientes.Zona_ID=Cat_Zonas.Zona_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Ciudades ON Cat_Clientes.Ciudad_ID=Cat_Ciudades.Ciudad_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Estados ON Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Vendedores ON Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Vendedor_ID<>'00011'"     'Los clientes que no son empleados
    'Valida si selecciono filtro de ciudad
    If Cmb_Cliente_Estadistico.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Cliente_Estadistico.ItemData(Cmb_Cliente_Estadistico.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Clientes
        Call Encabezado_Reportes("AUDITORIA DE CLIENTES " & Year(Dtp_Año_Auditoria.Value), Format(Now, "dd/MMM/yyyy"), "", True)
        Grid_Reporte.AddItem "CLIENTE" & Chr(9) & "RFC" & Chr(9) & "DOMICILIO" & Chr(9) & "TELEFONO" & Chr(9) & "CP" & Chr(9) & "CIUDAD" & Chr(9) & "VENDEDOR" & Chr(9) & "LISTA" & Chr(9) & "CREDITO" & Chr(9) & "DIAS" & Chr(9) & "MOVIMIENTOS" & Chr(9) & "TOTAL COMPRAS" & Chr(9) & "TOTAL SALDO"
        While Not .EOF
            Saldo_Cliente = 0
            Total_Compras = 0
            'Consulta el saldo general del cliente
            Mi_SQL = "SELECT SUM(Saldo) AS Saldo_Factura"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Pagada='N' AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND YEAR(Fecha)<=" & Year(Dtp_Año_Auditoria.Value)
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Consulta_Saldo_Clientes.rdoColumns("Saldo_Factura")) Then
                Saldo_Cliente = Rs_Consulta_Saldo_Clientes.rdoColumns("Saldo_Factura")
            Else
                Saldo_Cliente = 0
            End If
            Rs_Consulta_Saldo_Clientes.Close
            'Consulta si tuvo movimientos en el año del reporte
            Mi_SQL = "SELECT No_Factura FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND YEAR(Fecha)=" & Year(Dtp_Año_Auditoria.Value)
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Saldo_Clientes.EOF Then
                Tuvo_Movimientos = "SI"
            Else
                Rs_Consulta_Saldo_Clientes.Close
                Mi_SQL = "SELECT No_Remision FROM Adm_Remisiones_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND YEAR(Fecha)=" & Year(Dtp_Año_Auditoria.Value)
                Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Saldo_Clientes.EOF Then
                    Tuvo_Movimientos = "SI"
                Else
                    Tuvo_Movimientos = "NO"
                End If
            End If
            Rs_Consulta_Saldo_Clientes.Close
            'Consulta el total general de compras del cliente
            Mi_SQL = "SELECT SUM(Total) AS Compras"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND YEAR(Fecha)<=" & Year(Dtp_Año_Auditoria.Value)
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")) Then
                Total_Compras = Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")
            Else
                Total_Compras = 0
            End If
            Rs_Consulta_Saldo_Clientes.Close
            'Comienza la impresión
            Print #1,
            Print #1, "CLIENTE  : "; .rdoColumns("Cliente"); Spc(72 - Len("CLIENTE  : " & .rdoColumns("Cliente"))); "R.F.C.   : "; Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
            Print #1, "DOMICILIO: "; .rdoColumns("Direccion"); ", "; .rdoColumns("Colonia")
            Print #1, "TELEFONO : "; .rdoColumns("Telefono"); "   C.P.: "; .rdoColumns("CP")
            If Not IsNull(.rdoColumns("Ciudad")) Then
                Print #1, "CIUDAD   : "; .rdoColumns("Ciudad"); ", "; .rdoColumns("Estado")
            Else
                Print #1, "CIUDAD   : "
            End If
            Print #1, "VENDEDOR : "; .rdoColumns("Vendedor")
            Select Case .rdoColumns("Tipo_Precio")
                Case 1  'Precio Publico
                    Print #1, "LISTA    : PUBLICO        "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "PUBLICO" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 2  'Precio Tienda
                    Print #1, "LISTA    : TIENDA         "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "TIENDA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 3  'Precio Farmacia
                    Print #1, "LISTA    : FARMACIA       "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "FARMACIA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 4  'Precio Departamental
                    Print #1, "LISTA    : DEPARTAMENTAL  "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DEPARTAMENTAL" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 5  'Precio Distribuidor
                    Print #1, "LISTA    : DISTRIBUIDOR   "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DISTRIBUIDOR" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 6  'Precio Especial
                    Print #1, "LISTA    : ESPECIAL       "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "ESPECIAL" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 7  'Precio Bety Vega
                    Print #1, "LISTA    : BETY VEGA      "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "BETY VEGA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 8  'Precio VICMA
                    Print #1, "LISTA    : VICMA          "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "VICMA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 9  'Precio Departamental Anterior
                    Print #1, "LISTA    : DIST. ANTERIOR "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DEPT. ANTERIOR" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
            End Select
            Print #1, "MOVIMIENTOS : "; Tuvo_Movimientos; "               COMPRAS: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Compras, "#,###,##0.00"), 14); "   SALDO  : $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 14)
            Año_Ventas = Year(Dtp_Año_Auditoria.Value)
            Columnas = Grid_Reporte.Cols - (Year(Dtp_Año_Auditoria.Value) - 2007)
            While Año_Ventas > 2007
                'Consulta el total de compras del cliente
                Mi_SQL = "SELECT SUM(Total) AS Compras"
                Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND YEAR(Fecha)=" & Año_Ventas
                Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not IsNull(Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")) Then
                    Total_Compras = Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")
                Else
                    Total_Compras = 0
                End If
                Rs_Consulta_Saldo_Clientes.Close
                Print #1, "VENTAS AÑO "; Año_Ventas; " : "; "     $   "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Compras, "#,###,##0.00"), 13)
                Grid_Reporte.TextMatrix(6, Columnas) = "VENTAS AÑO " & Año_Ventas
                Grid_Reporte.TextMatrix(Grid_Reporte.Rows - 1, Columnas) = Format(Total_Compras, "#,###,##0.00")
                Columnas = Columnas + 1
                Año_Ventas = Año_Ventas - 1
            Wend
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            .MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    End With
    Rs_Consulta_Cat_Clientes.Close
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Estadistico_Clientes
    'DESCRIPCIÓN: Consulta los clientes y muestra estadísticas de interés
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 03-Julio-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Estadistico_Clientes()
Dim Mi_SQL As String
Dim Rs_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Fecha_Factura As rdoResultset
Dim Rs_Adm_Facturas_Clientes As rdoResultset
Dim Rs_Adm_Notas_Credtio_Clientes As rdoResultset
Dim Total_Nota_Credito As Double
Dim Total_Pagos As Double
Dim Compra_Total As Double
Dim Compra_Maxima As Double
Dim Compra_Minima As Double
Dim Cheque_Devuelto As Double
Dim Meses_Trasncurridos As Integer
Dim Puntos_Gestion As Integer
Dim Nivel_Gestion As String

    Call Encabezado_Reportes("ESTADISTICO DE CLIENTES", Format(Now, "dd/MMM/yyyy"), "", True)
    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
    Print #1, "                                  Estadístico   Limite        Notas         Compra        Compra       Cargos     Nivel  "
    Print #1, "      Cliente                     Pago X Mes    Crédito      Crédito        Máxima        Mínima      Bancarios  Gestión "
    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Cliente" & Chr(9) & "Estadístico Pago X Mes" & Chr(9) & "Límte de Crédito" & Chr(9) & "Notas de Crédito" & Chr(9) & "Compra Máxima" & Chr(9) & "Compra Mínima" & Chr(9) & "Cargos Bancarios" & Chr(9) & "Nivel de Gestión"
    Grid_Reporte.AddItem ""
    'Consulta los clientes
    Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes ORDER BY Nombre"
    Set Rs_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Cat_Clientes.EOF
        Puntos_Gestion = 0
        'Consulta la fecha de la primera compra del cliente y calcula el total de meses transcurridos
        Mi_SQL = "SELECT TOP 1 Fecha FROM Adm_Facturas_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Cancelada='N'"
        Mi_SQL = Mi_SQL & " ORDER BY Fecha"
        Set Rs_Consulta_Fecha_Factura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Fecha_Factura.EOF Then
            Meses_Trasncurridos = DateDiff("M", Format(Rs_Consulta_Fecha_Factura.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Now, "MM/dd/yyyy"))
        Else
            Meses_Trasncurridos = 0
        End If
        Rs_Consulta_Fecha_Factura.Close
        'Consulta el total pagado, el máximo y el mínimo de compra del cliente
        Mi_SQL = "SELECT SUM(Abono) AS Total_Pagos,MAX(Total) AS Maxima_Compra,MIN(Total) AS Minima_Compra,SUM(Total) AS Total_Compras"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Cancelada='N'"
        Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")) Then
            Total_Pagos = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Pagos")
            Compra_Maxima = Rs_Adm_Facturas_Clientes.rdoColumns("Maxima_Compra")
            Compra_Minima = Rs_Adm_Facturas_Clientes.rdoColumns("Minima_Compra")
            Compra_Total = Rs_Adm_Facturas_Clientes.rdoColumns("Total_Compras")
            'Valida que el límite supere la cantidad adeudada para sumar un punto a nivel de gestion
            If Meses_Trasncurridos > 0 Then
                If Val(Compra_Total - Total_Pagos) > (Total_Pagos / Meses_Trasncurridos * 2) Then
                    Puntos_Gestion = Puntos_Gestion + 1
                End If
            End If
        Else
            Total_Pagos = 0
            Compra_Maxima = 0
            Compra_Minima = 0
        End If
        Rs_Adm_Facturas_Clientes.Close
        'Consulta las notas de crédito
        Mi_SQL = "SELECT SUM(Importe) FROM Adm_Notas_Credito_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Autorizado='SI' AND Aplicada='SI'"
        Set Rs_Adm_Notas_Credtio_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not IsNull(Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)) Then
            Total_Nota_Credito = Rs_Adm_Notas_Credtio_Clientes.rdoColumns(0)
            'Valida el 20 % de compras superen el total de notas para agregar un punto a nivel de gestion
            If Val(Total_Nota_Credito / Compra_Total * 100) > 20 Then
                Puntos_Gestion = Puntos_Gestion + 1
            End If
        Else
            Total_Nota_Credito = 0
        End If
        Rs_Adm_Notas_Credtio_Clientes.Close
        'Consulta las comisiones por cheques devueltos
        Mi_SQL = "SELECT SUM(Adm_Detalles_Facturas.Importe) AS Importe,SUM(Adm_Detalles_Facturas.IVA) AS Iva"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID='" & Rs_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND Producto='COMISION CHEQUE DEVUELTO'"
        Set Rs_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not IsNull(Rs_Adm_Facturas_Clientes.rdoColumns("Importe")) Then
            Cheque_Devuelto = Rs_Adm_Facturas_Clientes.rdoColumns("Importe") + Rs_Adm_Facturas_Clientes.rdoColumns("Iva")
            Puntos_Gestion = Puntos_Gestion + 1
        Else
            Cheque_Devuelto = 0
        End If
        Rs_Adm_Facturas_Clientes.Close
        'Valida el nivel de gestión
        Select Case Puntos_Gestion
            Case 0
                Nivel_Gestion = "BAJO"
            Case 1
                Nivel_Gestion = "BAJO"
            Case 2
                Nivel_Gestion = "MEDIO"
            Case 3
                Nivel_Gestion = "ALTO"
        End Select
        If Meses_Trasncurridos > 0 Then
            Print #1, Mid(Rs_Cat_Clientes.rdoColumns("Nombre"), 1, 28); _
                Spc(30 - Len(Mid(Rs_Cat_Clientes.rdoColumns("Nombre"), 1, 28))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "#,###,##0.00"), 12); _
                Spc(2); Nivel_Gestion
            Grid_Reporte.AddItem Rs_Cat_Clientes.rdoColumns("Nombre") _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pagos / Meses_Trasncurridos * 2, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Cheque_Devuelto, "###,##0.00"), 9) _
                & Chr(9) & Nivel_Gestion
        Else
            Print #1, Mid(Rs_Cat_Clientes.rdoColumns("Nombre"), 1, 28); _
                Spc(30 - Len(Mid(Rs_Cat_Clientes.rdoColumns("Nombre"), 1, 28))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12); _
                Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12); _
                Spc(2); Nivel_Gestion
            Grid_Reporte.AddItem Rs_Cat_Clientes.rdoColumns("Nombre") _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Nota_Credito, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Maxima, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Compra_Minima, "#,###,##0.00"), 12) _
                & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(0, "###,##0.00"), 9) _
                & Chr(9) & Nivel_Gestion
        End If
        Rs_Cat_Clientes.MoveNext
    Wend
    Rs_Cat_Clientes.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Clientes_Empleados
    'DESCRIPCIÓN: Consulta los datos de los clientes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 21-Mayo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Reporte_Clientes_Empleados()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Saldo_Clientes As rdoResultset
Dim Saldo_Cliente As Double
Dim Total_Compras As Double
Dim Tuvo_Movimientos As String

    Grid_Reporte.Cols = 13
    Grid_Reporte.Rows = 0
    'Consulta los clientes del vendedor en turno
    Mi_SQL = "SELECT Cliente_ID,Cat_Clientes.Nombre AS Cliente,Cat_Clientes.RFC,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Telefono,Cat_Clientes.Estatus,Cat_Clientes.Tipo_Precio,Cat_Clientes.Dias_Credito,Cat_Clientes.Limite_Credito"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Ciudad_ID,Cat_Ciudades.Nombre AS Ciudad,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre AS Estado,Cat_Clientes.Zona_ID,Cat_Zonas.Nombre AS Zona"
    Mi_SQL = Mi_SQL & " ,Cat_Clientes.Vendedor_ID,Cat_Vendedores.Nombre AS Vendedor"
    Mi_SQL = Mi_SQL & " FROM Cat_Clientes LEFT JOIN Cat_Zonas ON Cat_Clientes.Zona_ID=Cat_Zonas.Zona_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Ciudades ON Cat_Clientes.Ciudad_ID=Cat_Ciudades.Ciudad_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Estados ON Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Vendedores ON Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
    Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Vendedor_ID='00011'"      'Los clientes que son empleados
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Clientes.Nombre"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Clientes
        Call Encabezado_Reportes("AUDITORIA DE EMPLEADOS", Format(Now, "dd/MMM/yyyy"), "", True)
        Grid_Reporte.AddItem "CLIENTE" & Chr(9) & "RFC" & Chr(9) & "DOMICILIO" & Chr(9) & "TELEFONO" & Chr(9) & "CP" & Chr(9) & "CIUDAD" & Chr(9) & "VENDEDOR" & Chr(9) & "LISTA" & Chr(9) & "CREDITO" & Chr(9) & "DIAS" & Chr(9) & "MOVIMIENTOS" & Chr(9) & "COMPRAS" & Chr(9) & "SALDO"
        While Not .EOF
            Saldo_Cliente = 0
            Total_Compras = 0
            'Consulta el saldo del cliente
            Mi_SQL = "SELECT SUM(Saldo) AS Saldo_Factura"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Pagada='N' AND Cancelada='N'"
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Consulta_Saldo_Clientes.rdoColumns("Saldo_Factura")) Then
                Saldo_Cliente = Rs_Consulta_Saldo_Clientes.rdoColumns("Saldo_Factura")
            End If
            Rs_Consulta_Saldo_Clientes.Close
            Mi_SQL = "SELECT No_Factura FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Mi_SQL = Mi_SQL & " AND YEAR(Fecha)=" & Format(Now, "yyyy")
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Saldo_Clientes.EOF Then
                Tuvo_Movimientos = "SI"
            Else
                Rs_Consulta_Saldo_Clientes.Close
                Mi_SQL = "SELECT No_Remision FROM Adm_Remisiones_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND YEAR(Fecha)=" & Format(Now, "yyyy")
                Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Saldo_Clientes.EOF Then
                    Tuvo_Movimientos = "SI"
                Else
                    Tuvo_Movimientos = "NO"
                End If
            End If
            Rs_Consulta_Saldo_Clientes.Close
            'Consulta el saldo del cliente
            Mi_SQL = "SELECT SUM(Total) AS Compras"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Cancelada='N'"
            Set Rs_Consulta_Saldo_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not IsNull(Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")) Then
                Total_Compras = Rs_Consulta_Saldo_Clientes.rdoColumns("Compras")
            End If
            Rs_Consulta_Saldo_Clientes.Close
            'Comienza la impresión
            Print #1,
            Print #1, "CLIENTE  : "; .rdoColumns("Cliente"); Spc(72 - Len("CLIENTE  : " & .rdoColumns("Cliente"))); "R.F.C.   : "; Rs_Consulta_Cat_Clientes.rdoColumns("RFC")
            Print #1, "DOMICILIO: "; .rdoColumns("Direccion"); ", "; .rdoColumns("Colonia")
            Print #1, "TELEFONO : "; .rdoColumns("Telefono"); "   C.P.: "; .rdoColumns("CP")
            If Not IsNull(.rdoColumns("Ciudad")) Then
                Print #1, "CIUDAD   : "; .rdoColumns("Ciudad"); ", "; .rdoColumns("Estado")
            Else
                Print #1, "CIUDAD   : "
            End If
            Print #1, "VENDEDOR : "; .rdoColumns("Vendedor")
            Select Case .rdoColumns("Tipo_Precio")
                Case 1  'Precio Publico
                    Print #1, "LISTA    : PUBLICO        "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "PUBLICO" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 2  'Precio Tienda
                    Print #1, "LISTA    : TIENDA         "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "TIENDA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 3  'Precio Farmacia
                    Print #1, "LISTA    : FARMACIA       "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "FARMACIA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 4  'Precio Departamental
                    Print #1, "LISTA    : DEPARTAMENTAL  "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DEPARTAMENTAL" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 5  'Precio Distribuidor
                    Print #1, "LISTA    : DISTRIBUIDOR   "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DISTRIBUIDOR" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 6  'Precio Especial
                    Print #1, "LISTA    : ESPECIAL       "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "ESPECIAL" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 7  'Precio Bety Vega
                    Print #1, "LISTA    : BETY VEGA      "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "BETY VEGA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 8  'Precio VICMA
                    Print #1, "LISTA    : VICMA          "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "VICMA" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
                Case 9  'Precio Departamental Anterior
                    Print #1, "LISTA    : DIST. ANTERIOR "; Spc(5); "CREDITO: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Limite_Credito"), "#,###,##0.00"), 14); Spc(3); "DIAS   : "; .rdoColumns("Dias_Credito")
                    Grid_Reporte.AddItem .rdoColumns("Cliente") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad") & ", " & Rs_Consulta_Cat_Clientes.rdoColumns("Estado") & Chr(9) & .rdoColumns("Vendedor") & Chr(9) & "DEPT. ANTERIOR" & Chr(9) & Format(.rdoColumns("Limite_Credito"), "#,###,##0.00") & Chr(9) & .rdoColumns("Dias_Credito") & Chr(9) & Tuvo_Movimientos & Chr(9) & Format(Total_Compras, "#,###,##0.00") & Chr(9) & Format(Saldo_Cliente, "#,###,##0.00")
            End Select
            Print #1, "MOVIMIENTOS " & Year(Now) & " : "; Tuvo_Movimientos; "          COMPRAS: $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Compras, "#,###,##0.00"), 14); "   SALDO  : $"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Saldo_Cliente, "#,###,##0.00"), 14)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            .MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    End With
    Rs_Consulta_Cat_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Proveedores
    'DESCRIPCIÓN: Consulta los datos de los clientes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 26-Mayo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Reporte_Proveedores()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Proveedores As rdoResultset            'Manejo de registro

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    'Consulta los clientes del vendedor en turno
    Mi_SQL = "SELECT Proveedor_ID,Nombre,RFC,Direccion,Colonia,CP,Telefono1,Fax,Estatus,Ciudad_Estado"
    Mi_SQL = Mi_SQL & " FROM Cat_Proveedores"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Proveedores
        Call Encabezado_Reportes("CATALOGO DE PROVEEDORES", Format(Now, "dd/MMM/yyyy"), "", True)
        Grid_Reporte.AddItem "PROVEEDOR" & Chr(9) & "RFC" & Chr(9) & "DOMICILIO" & Chr(9) & "TELEFONO" & Chr(9) & "CP" & Chr(9) & "CIUDAD" & Chr(9) & "FAX"
        While Not .EOF
            Print #1,
            Print #1, "PROVEEDOR: "; .rdoColumns("Nombre"); Spc(72 - Len("PROVEEDOR: " & .rdoColumns("Nombre"))); "R.F.C.   : "; .rdoColumns("RFC")
            Print #1, "DOMICILIO: "; .rdoColumns("Direccion"); ", "; .rdoColumns("Colonia")
            Print #1, "TELEFONO : "; .rdoColumns("Telefono1"); "   C.P.: "; .rdoColumns("CP"); "   FAX: "; .rdoColumns("Fax")
            Print #1, "CIUDAD   : "; .rdoColumns("Ciudad_Estado")
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem .rdoColumns("Nombre") & Chr(9) & .rdoColumns("RFC") & Chr(9) & .rdoColumns("Direccion") & ", " & .rdoColumns("Colonia") & Chr(9) & .rdoColumns("Telefono1") & Chr(9) & .rdoColumns("CP") & Chr(9) & .rdoColumns("Ciudad_Estado") & Chr(9) & .rdoColumns("Fax")
            .MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    End With
    Rs_Consulta_Cat_Proveedores.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Facturas_Clientes
    'DESCRIPCIÓN: Consulta las factutas del cliente seleccionado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 30-Abril-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Facturas_Clientes()
Dim Cliente As String                                       'Obtiene el nombre del cliente
Dim Rs_Consulta_Clientes As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_IVA As Double                                     'Almacena el total del IVA
Dim Total_Tasa_0 As Double                                  'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                                 'Almacena el total de la tasa 15
Dim Total_Facturas As Double                                'Almacena el total de la factura

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Cliente_Factura.ListIndex > -1 Then
        Mi_SQL = "SELECT Cliente_ID,Nombre,Direccion,Colonia,CP,Ciudad_ID"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Cliente_Factura.ItemData(Cmb_Cliente_Factura.ListIndex), "00000") & "'"
        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Clientes.EOF Then
            With Rs_Consulta_Clientes
                If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                    Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                    Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                    Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                    Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Ciudades.EOF Then
                        Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                    Else
                        Ciudad = ""
                    End If
                    Rs_Consulta_Ciudades.Close
                Else
                    Ciudad = ""
                End If
                Call Encabezado_Reportes("REPORTE DE FACTURAS POR CLIENTE", Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), True)
                Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Factura.Text
                Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Factura.Text
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End With
            Grid_Reporte.AddItem ""
            'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
            Mi_SQL = "SELECT No_Factura,Empresa,Cliente_ID,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
                With Rs_Consulta_Cat_Facturas_Clientes
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "     No. Factura      Fecha           Sin IVA           Con IVA             IVA              Total         Estatus        "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Sin IVA" & _
                                         Chr(9) & "Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                    Grid_Reporte.AddItem ""
                    While Not .EOF
                        If Rs_Consulta_Cat_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                        'Consulta el estatus
                        If .rdoColumns("Cancelada") = "S" Then
                            Estatus = "Cancelada"
                        Else
                            If .rdoColumns("Pagada") = "S" Then
                                Estatus = "Pagada"
                            Else
                                Estatus = "Pendiente"
                            End If
                        End If
                        If .rdoColumns("Cancelada") = "N" Then
                            Print #1, Spc(2); .rdoColumns("No_Factura"); _
                                Spc(2); Empresa; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                Format(.rdoColumns("Subtotal_Tasa_0"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Subtotal_Tasa_15"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & Estatus & Chr(13)
                            Total_IVA = Total_IVA + .rdoColumns("IVA")
                            Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Subtotal_Tasa_0")
                            Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Subtotal_Tasa_15")
                            Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Else
                            Print #1, Spc(2); .rdoColumns("No_Factura"); _
                                Spc(2); Empresa; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & Estatus
                        End If
                        .MoveNext
                    Wend
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Spc(10); "T O T A L E S "; Spc(9); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_IVA, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
                    Call Encabezado_Reportes("", "", "", False)
                End With
            Else
                MsgBox "No hay facturas del cliente seleccionado", vbInformation
            End If
        End If
        Rs_Consulta_Cat_Facturas_Clientes.Close
    Else
        MsgBox "Seleccione un cliente", vbExclamation
    End If
End Sub

    'NOMBRE DE LA FUNCIÓN: Reporte_Remisiones_Clientes
    'DESCRIPCIÓN: Consulta las remisiones del cliente seleccionado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Remisiones_Clientes()
Dim Cliente As String                                       'Obtiene el nombre del cliente
Dim Rs_Consulta_Clientes As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset                    'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_IVA As Double                                     'Almacena el total del IVA
Dim Total_Tasa_0 As Double                                  'Almacena el total de la Sin IVA
Dim Total_Tasa_15 As Double                                 'Almacena el total de la tasa 15
Dim Total_Facturas As Double                                'Almacena el total de la factura

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Cliente_Factura.ListIndex > -1 Then
        Mi_SQL = "SELECT Cliente_ID,Nombre,Direccion,Colonia,CP,Ciudad_ID"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Format(Cmb_Cliente_Factura.ItemData(Cmb_Cliente_Factura.ListIndex), "00000") & "'"
        Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Clientes.EOF Then
            With Rs_Consulta_Clientes
                If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                    Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                    Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                    Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                    Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Ciudades.EOF Then
                        Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                    Else
                        Ciudad = ""
                    End If
                    Rs_Consulta_Ciudades.Close
                Else
                    Ciudad = ""
                End If
                Call Encabezado_Reportes("REPORTE DE FACTURAS POR CLIENTE", Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), True)
                Print #1, Spc(2); "Cliente:   "; Cmb_Cliente_Factura.Text
                Print #1, Spc(2); "Direccion: "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:   "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:    "; Ciudad
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Cmb_Cliente_Factura.Text
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
            End With
            Grid_Reporte.AddItem ""
            'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
            Mi_SQL = "SELECT No_Remision,Empresa,Cliente_ID,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Pagada,Cancelada"
            Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Clientes.rdoColumns("Cliente_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
                With Rs_Consulta_Cat_Facturas_Clientes
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, "    No. Remision      Fecha           Sin IVA           Con IVA             IVA              Total         Estatus        "
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Sin IVA" & _
                                         Chr(9) & "Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Estatus"
                    Grid_Reporte.AddItem ""
                    While Not .EOF
                        If Rs_Consulta_Cat_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                        'Consulta el estatus
                        If .rdoColumns("Cancelada") = "S" Then
                            Estatus = "Cancelada"
                        Else
                            If .rdoColumns("Pagada") = "S" Then
                                Estatus = "Pagada"
                            Else
                                Estatus = "Pendiente"
                            End If
                        End If
                        If .rdoColumns("Cancelada") = "N" Then
                            Print #1, Spc(2); .rdoColumns("No_Remision"); _
                                Spc(2); Empresa; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "#,###,##0.00"), 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                Format(.rdoColumns("Subtotal_Tasa_0"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Subtotal_Tasa_15"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & Estatus & Chr(13)
                            Total_IVA = Total_IVA + .rdoColumns("IVA")
                            Total_Tasa_0 = Total_Tasa_0 + .rdoColumns("Subtotal_Tasa_0")
                            Total_Tasa_15 = Total_Tasa_15 + .rdoColumns("Subtotal_Tasa_15")
                            Total_Facturas = Total_Facturas + .rdoColumns("Total")
                        Else
                            Print #1, Spc(2); .rdoColumns("No_Remision"); _
                                Spc(2); Empresa; _
                                Spc(3); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 14); _
                                Spc(4); Estatus
                            Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & Estatus
                        End If
                        .MoveNext
                    Wend
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, Spc(10); "T O T A L E S "; Spc(9); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_0, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Tasa_15, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_IVA, "#,###,##0.00"), 14); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Tasa_0, "#,###,##0.00") & Chr(9) & Format(Total_Tasa_15, "#,###,##0.00") & Chr(9) & Format(Total_IVA, "#,###,##0.00") & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
                End With
            End If
            Call Encabezado_Reportes("", "", "", False)
        End If
        Rs_Consulta_Cat_Facturas_Clientes.Close
    Else
        MsgBox "Seleccione un cliente", vbExclamation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Facturas_Pagadas
    'DESCRIPCIÓN: Consulta las factutas que fueron pagadas
    'PARÁMETROS :
    'CREO       : Rafael Muñoz
    'FECHA_CREO : 31-Febrero-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Facturas_Pagadas()
Dim Cliente As String                                   'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset   'Manejo de registro

Grid_Reporte.Cols = 8
Grid_Reporte.Rows = 0
'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
Mi_SQL = "SELECT No_Factura, Cliente_ID, Empresa,"
Mi_SQL = Mi_SQL & " Fecha, Subtotal_Tasa_0, Subtotal_Tasa_15, IVA, Total"
Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
Mi_SQL = Mi_SQL & " WHERE Fecha >= " & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
Mi_SQL = Mi_SQL & " AND Pagada='S'"
Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
    With Rs_Consulta_Cat_Facturas_Clientes
        Call Encabezado_Reportes("REPORTE DE FACTURAS PAGADAS", Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), True)
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(2); "No. Factura"; Spc(12 - Len("No. Factura")); "Cliente"; _
                  Spc(42 - Len("Cliente")); "Fecha"; Spc(15 - Len("Fecha")); "Sin IVA"; _
                  Spc(11 - Len("Sin IVA")); "Con IVA"; Spc(16 - Len("Con IVA")); "IVA"; _
                  Spc(10 - Len("IVA")); "Total"; Spc(6 - Len("Total"));
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1,
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Sin IVA" & _
                             Chr(9) & "Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total"
        Grid_Reporte.AddItem ""
        While Not .EOF
            If Not IsNull(.rdoColumns("Cliente_ID")) Then
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
            End If
            Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 40); _
            Spc(42 - Len(Mid(Cliente, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
            Spc(12 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
            Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "##,##0.00"), 9); _
            Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "##,##0.00"), 9))); _
            Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "##,##0.00"), 9); _
            Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "##,##0.00"), 9))); _
            Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9); _
            Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9))); _
            Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9); _
            Spc(10 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9)))
            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & "Cliente" & Chr(9) & _
            Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
            Format(.rdoColumns("Subtotal_Tasa_0"), "###,##0.00") & Chr(9) & _
            Format(.rdoColumns("Subtotal_Tasa_15"), "###,##0.00") & Chr(9) & _
            Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
            Format(.rdoColumns("Total"), "##,##0.00") & Chr(13)
            .MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    End With
Else
    MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
    DTP_Fecha_Inicio_Historico_Factura.Value = Now
    DTP_Fecha_Final_Historico_Factura.Value = Now
End If
Rs_Consulta_Cat_Facturas_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Facturas_Canceladas
    'DESCRIPCIÓN: Consulta las factutas que fueron canceladas
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 25-MArzo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Facturas_Canceladas()
Dim Cliente As String                                 'Obtiene el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset          'Consulta el cliente al cual pertenece la factura emitida
Dim Rs_Consulta_Cat_Facturas_Clientes As rdoResultset 'Consulta las facturas que fueron emitidas en las fechas dadas por el usuario

    Grid_Reporte.Cols = 12
    Grid_Reporte.Rows = 0
    'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
    Mi_SQL = "SELECT No_Factura,Empresa,Cliente_ID,Fecha,Subtotal_Tasa_0,Subtotal_Tasa_15,IVA,Total,Motivo_Cancelacion,Fecha_Cancelo,Usuario_Cancelo,No_Factura_Reposicion,Tipo_Cancelacion"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(DTP_Fecha_Inicio_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(DTP_Fecha_Final_Historico_Factura.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Cancelada='S'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha"
    Set Rs_Consulta_Cat_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Facturas_Clientes.EOF Then
        With Rs_Consulta_Cat_Facturas_Clientes
            Call Encabezado_Reportes("REPORTE DE FACTURAS CANCELADAS", Format(DTP_Fecha_Inicio_Historico_Factura.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Factura.Value, "dd/MMM/yyyy"), True)
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(2); "No. Factura"; Spc(12 - Len("No. Factura")); "Cliente"; _
                      Spc(42 - Len("Cliente")); "Fecha"; Spc(15 - Len("Fecha")); "Sin IVA"; _
                      Spc(11 - Len("Sin IVA")); "Con IVA"; Spc(16 - Len("Con IVA")); "IVA"; _
                      Spc(12 - Len("IVA")); "Total"; Spc(12 - Len("Total"));
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Sin IVA" & Chr(9) & "Con IVA" & Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Fecha Cancelo" & Chr(9) & "Usuario Cancelo" & Chr(9) & "Tipo Cancelacion" & Chr(9) & "Nueva Factura" & Chr(9) & "Motivo Cancelo"
            Grid_Reporte.AddItem ""
            While Not .EOF
                If Not IsNull(.rdoColumns("Cliente_ID")) Then
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Cliente_ID,Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    Else
                        Cliente = "SIN NOMBRE"
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                Else
                    Cliente = "SIN NOMBRE"
                End If
                Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Cliente, 1, 40); _
                    Spc(42 - Len(Mid(Cliente, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(12 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_0"), "##,##0.00"), 11))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Subtotal_Tasa_15"), "##,##0.00"), 11))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 11))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 11); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 11)))
                Print #1, Spc(5); "FECHA CANCELO: "; Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy HH:mm:ss"); Spc(5); "CANCELO: "; .rdoColumns("Usuario_Cancelo")
                Print #1, Spc(5); "MOTIVO: "; .rdoColumns("Motivo_Cancelacion")
                If Not IsNull(.rdoColumns("Tipo_Cancelacion")) Then
                    If Trim(.rdoColumns("Tipo_Cancelacion")) = "REFACTURACION" Then
                        Print #1, Spc(5); "TIPO CANCELACION : REFACTURACION"; Spc(3); "NUEVA FACTURA: "; .rdoColumns("No_Factura_Reposicion")
                    Else
                        Print #1, Spc(5); "TIPO CANCELACION : DEVOLUCION"
                    End If
                Else
                    If Not IsNull(.rdoColumns("Tipo_Cancelacion")) Then
                        Print #1, Spc(5); "NUEVA FACTURA: "; .rdoColumns("No_Factura_Reposicion")
                    End If
                End If
                Print #1,
                Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Cliente & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    Format(.rdoColumns("Subtotal_Tasa_0"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Subtotal_Tasa_15"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Total"), "##,##0.00") & Chr(9) & _
                    Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy HH:mm:ss") & Chr(9) & _
                    .rdoColumns("Usuario_Cancelo") & Chr(9) & _
                    .rdoColumns("Tipo_Cancelacion") & Chr(9) & _
                    .rdoColumns("No_Factura_Reposicion") & Chr(9) & _
                    .rdoColumns("Motivo_Cancelacion")
                .MoveNext
            Wend
            Call Encabezado_Reportes("", "", "", False)
        End With
    Else
        MsgBox "No hay facturas en las fechas que proporcionó", vbInformation
        DTP_Fecha_Inicio_Historico_Factura.Value = Now
        DTP_Fecha_Final_Historico_Factura.Value = Now
    End If
    Rs_Consulta_Cat_Facturas_Clientes.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Ventas_Productos
    'DESCRIPCIÓN: Consulta las ventas del producto seleccionado en un rango de
    '             fechas proporcionado
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 11-Diciembre-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Ventas_Productos()
Dim Dia As Integer                                      'Indica que día es el que se esta consultando
Dim Dias As Integer                                     'Indica la cantidad de días a consultar
Dim Total_Cantidad As Double
Dim Total_Ventas As Double                              'Indica el total de las ventas obtenidas
Dim Empresa As String                                   'Almacena el nombre de la empresa
Dim Estatus As String                                   'Almacena el estatus de la factura
Dim Numero_Lote As String
Dim Fecha_Caducidad As String
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Rs_Consulta_Detalles_Salidas As rdoResultset

    If Cmb_Producto_Venta.ListIndex > -1 Then
        'Consulta las facturas realizadas durante el día que se esta seleccionado
        Mi_SQL = "SELECT Adm_Facturas_Clientes.No_Factura,Adm_Facturas_Clientes.Empresa,Adm_Facturas_Clientes.Fecha,Adm_Facturas_Clientes.Pagada,Adm_Facturas_Clientes.Cancelada,Adm_Detalles_Facturas.Producto_ID,Adm_Detalles_Facturas.Cantidad,Adm_Detalles_Facturas.Numero_Lote,Adm_Detalles_Facturas.Fecha_Caducidad,Adm_Detalles_Facturas.Importe,Adm_Facturas_Clientes.Cliente_ID,Cat_Clientes.Nombre"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas,Cat_Productos,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
        Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Producto_ID=Cat_Productos.Producto_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Producto_ID='" & Format(Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.ListIndex), "00000") & "'"
        If Chk_Fechas.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
        End If
        If Chk_Numero_Lote.Value = 1 And Txt_Numero_Lote.Text <> "" Then
            Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Numero_Lote LIKE '%" & Txt_Numero_Lote.Text & "%'"
        End If
        If Chk_Fecha_Caducidad.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Fecha_Caducidad='" & Format(Dtp_Fecha_Caducidad.Value, "MM/dd/yyyy") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Facturas_Clientes.Fecha"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            Call Encabezado_Reportes("REPORTE DE DISTRIBUCION DE VENTAS POR PRODUCTO", Format(DTP_Fecha_Inicio_Historico_Ventas, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Ventas, "dd/MMM/yyyy"), True)
            Print #1, Conectar_Ayudante.Centra("PRODUCTO : " & Cmb_Producto_Venta.Text, Len("------------------------------------------------------------------------------------------------------------------------"))
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, "Fecha Venta   No. Factura    Estatus   Cliente                       Cantidad  No. Lote         Caducidad       Importe "
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Producto : " & Chr(9) & Cmb_Producto_Venta.Text
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Fecha Venta" & Chr(9) & "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Cliente" & Chr(9) & "Cantidad" & Chr(9) & "Numero Lote" & Chr(9) & "Fecha de Caducidad" & Chr(9) & "Importe"
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                'Consulta la empresa
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                Total_Cantidad = Total_Cantidad + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cantidad")
                'Valida el estatus
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancel"
                Else
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Pagada") = "N" Then
                        Estatus = "S/Paga"
                    Else
                        Estatus = "Pagada"
                    End If
                    Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Importe")
                End If
                'Valida si tiene Lote y caducidad
                If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Numero_Lote")) Then
                    Numero_Lote = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Numero_Lote")
                    Fecha_Caducidad = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha_Caducidad")
                Else
                    'Consulta en la salida la información
                    Mi_SQL = "SELECT No_Salida,Numero_Lote,Fecha_Caducidad"
                    Mi_SQL = Mi_SQL & " FROM Detalles_Salidas"
                    Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & "'"
                    Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")
                    Mi_SQL = Mi_SQL & " AND Producto_ID='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Producto_ID") & "'"
                    If Chk_Numero_Lote.Value = 1 And Txt_Numero_Lote.Text <> "" Then
                        Mi_SQL = Mi_SQL & " AND Numero_Lote LIKE '%" & Txt_Numero_Lote.Text & "%'"
                    End If
                    If Chk_Fecha_Caducidad.Value = 1 Then
                        Mi_SQL = Mi_SQL & " AND Fecha_Caducidad='" & Format(Dtp_Fecha_Caducidad.Value, "MM/dd/yyyy") & "'"
                    End If
                    Set Rs_Consulta_Detalles_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Detalles_Salidas.EOF Then
                        If Not IsNull(Rs_Consulta_Detalles_Salidas.rdoColumns("Numero_Lote")) Then
                            Numero_Lote = Rs_Consulta_Detalles_Salidas.rdoColumns("Numero_Lote")
                            Fecha_Caducidad = Rs_Consulta_Detalles_Salidas.rdoColumns("Fecha_Caducidad")
                        Else
                            Numero_Lote = "               "
                            Fecha_Caducidad = "           "
                        End If
                    Else
                        Numero_Lote = "               "
                        Fecha_Caducidad = "           "
                    End If
                    Rs_Consulta_Detalles_Salidas.Close
                End If
                If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cancelada") = "N" Then
                    Print #1, Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(2); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); _
                        Spc(1); Empresa; _
                        Spc(2); Estatus; _
                        Spc(8 - Len(Estatus)); Mid(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre"), 1, 30); _
                        Spc(32 - Len(Mid(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre"), 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cantidad"), "###,##0"), 7); _
                        Spc(2); Mid(Numero_Lote, 1, 15); _
                        Spc(17 - Len(Mid(Numero_Lote, 1, 15))); Format(Fecha_Caducidad, "dd/MMM/yyyy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Importe"), "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") _
                        & Chr(9) & Empresa _
                        & Chr(9) & Estatus _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cantidad") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Importe"), "#,###,##0.00")
                Else
                    Print #1, Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                        Spc(2); Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); _
                        Spc(1); Empresa; _
                        Spc(2); Estatus; _
                        Spc(8 - Len(Estatus)); Mid(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre"), 1, 30); _
                        Spc(32 - Len(Mid(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre"), 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cantidad"), "###,##0"), 7); _
                        Spc(2); Mid(Numero_Lote, 1, 15); _
                        Spc(17 - Len(Mid(Numero_Lote, 1, 15))); Format(Fecha_Caducidad, "dd/MMM/yyyy"); _
                        Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha("0.00", 12)
                    Grid_Reporte.AddItem Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") _
                        & Chr(9) & Empresa _
                        & Chr(9) & Estatus _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Nombre") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cantidad") _
                        & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Numero_Lote") _
                        & Chr(9) & Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                        & Chr(9) & "0.00"
                End If
                Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(45); "T O T A L E S"; Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cantidad, "#,###,##0"), 14); Spc(23); " $ "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,##0.00"), 17)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Cantidad Total" & Chr(9) & Total_Cantidad & Chr(9) & Chr(9) & "Total Venta" & Chr(9) & Format(Total_Ventas, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        Else
            MsgBox "No hay datos que consultar", vbInformation
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
    Else
        MsgBox "Seleccione un producto de la lista", vbExclamation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Reporte_Materia_Prima
    'DESCRIPCION: Consulta la materia prima registrada en el sistema
    'PARAMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 18-Mayo-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'*******************************************************************************
Private Sub Reporte_Materia_Prima()
Dim Numero_Acta As String
Dim Numero_Lote As String
Dim Fecha_Caducidad As String
Dim Rs_Consulta_Entradas As rdoResultset
Dim Rs_Consulta_Muestras As rdoResultset
Dim Existencia As Double
Dim Comentarios_Libero As String
Dim Usuario_Muestreo As String

    'Consulta las facturas realizadas durante el día que se esta seleccionado
    Mi_SQL = "SELECT Alm_Entradas_Materia_Prima_Detalles.No_Entrada,Alm_Entradas_Materia_Prima_Detalles.Empresa,Alm_Entradas_Materia_Prima.Fecha,Alm_Entradas_Materia_Prima_Detalles.Estatus,Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID,Alm_Entradas_Materia_Prima_Detalles.Cantidad,Alm_Entradas_Materia_Prima_Detalles.Faltante,Alm_Entradas_Materia_Prima_Detalles.Numero_Lote,Alm_Entradas_Materia_Prima_Detalles.Fecha_Caducidad,Alm_Entradas_Materia_Prima_Detalles.No_Acta,Cat_Materias_Primas.Nombre,Cat_Materias_Primas.Codigo,Alm_Entradas_Materia_Prima_Detalles.Comentarios_Libero"
    Mi_SQL = Mi_SQL & " FROM Alm_Entradas_Materia_Prima,Alm_Entradas_Materia_Prima_Detalles,Cat_Materias_Primas"
    Mi_SQL = Mi_SQL & " WHERE Alm_Entradas_Materia_Prima.No_Entrada=Alm_Entradas_Materia_Prima_Detalles.No_Entrada"
    Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID=Cat_Materias_Primas.Materia_Prima_ID"
    If Chk_Producto.Value = 1 And Cmb_Producto_Venta.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID='" & Format(Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.ListIndex), "00000") & "'"
    End If
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Numero_Lote.Value = 1 And Trim(Txt_Numero_Lote.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Numero_Lote LIKE '%" & Trim(Txt_Numero_Lote.Text) & "%'"
    End If
    If Chk_Fecha_Caducidad.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Fecha_Caducidad='" & Format(Dtp_Fecha_Caducidad.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Estatus_Materia_Prima.Value = 1 And Cmb_Estatus_Materia_Prima.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Estatus='" & Cmb_Estatus_Materia_Prima.Text & "'"
    End If
    If Chk_Numero_Acta.Value = 1 And Trim(Txt_Numero_Acta.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.No_Acta LIKE '%" & Trim(Txt_Numero_Acta.Text) & "%'"
    End If
    If InStr(1, Rol, "ALMACEN") Then
        Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima_Detalles.Faltante>0"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Alm_Entradas_Materia_Prima.Fecha"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Entradas.EOF Then
        Grid_Reporte.Cols = 11
        Call Encabezado_Reportes("REPORTE DE MATERIAS PRIMAS", Format(DTP_Fecha_Inicio_Historico_Ventas, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Ventas, "dd/MMM/yyyy"), True)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "  Fecha      Entrada    Estatus   Clave             Producto             Existencia   No. Lote       Caducidad  No. Acta"
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Entrada" & Chr(9) & "Estatus" & Chr(9) & "Clave" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Numero Lote" & Chr(9) & "Fecha de Caducidad" & Chr(9) & "No. Acta" & Chr(9) & "Comentarios Liberacion" & Chr(9) & "Muestreo"
        While Not Rs_Consulta_Entradas.EOF
            'Valida si tiene Lote y caducidad
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Numero_Lote")) Then
                Numero_Lote = Rs_Consulta_Entradas.rdoColumns("Numero_Lote")
            Else
                Numero_Lote = "             "
            End If
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad")) Then
                Fecha_Caducidad = Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy")
            Else
                Fecha_Caducidad = "           "
            End If
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("No_Acta")) Then
                Numero_Acta = Rs_Consulta_Entradas.rdoColumns("No_Acta")
            Else
                Numero_Acta = ""
            End If
            'If InStr(1, Rol, "ALMACEN") Then
                Existencia = Rs_Consulta_Entradas.rdoColumns("Faltante")
            'Else
            '    Existencia = Rs_Consulta_Entradas.rdoColumns("Cantidad")
            'End If
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Comentarios_Libero")) Then
                Comentarios_Libero = "-> " & Mid(Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Entradas.rdoColumns("Comentarios_Libero"), Chr(10)), Chr(13)), 1, 20)
            Else
                Comentarios_Libero = ""
            End If
            'Consulta el usuario que muestreo
            Mi_SQL = "SELECT No_Muestra,Muestreo FROM Alm_Toma_Muestras"
            Mi_SQL = Mi_SQL & " WHERE No_Entrada='" & Rs_Consulta_Entradas.rdoColumns("No_Entrada") & "'"
            Set Rs_Consulta_Muestras = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Muestras.EOF Then
                If Not IsNull(Rs_Consulta_Muestras.rdoColumns("Muestreo")) Then
                    Usuario_Muestreo = Rs_Consulta_Muestras.rdoColumns("Muestreo")
                Else
                    Usuario_Muestreo = ""
                End If
            Else
                Usuario_Muestreo = ""
            End If
            Rs_Consulta_Muestras.Close
            Print #1, Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(1); Rs_Consulta_Entradas.rdoColumns("No_Entrada"); _
                Spc(1); Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 8); _
                Spc(10 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 8))); Mid(Rs_Consulta_Entradas.rdoColumns("Codigo"), 1, 16); _
                Spc(18 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Codigo"), 1, 16))); Mid(Rs_Consulta_Entradas.rdoColumns("Nombre"), 1, 20); _
                Spc(21 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Nombre"), 1, 20))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Existencia, "#,###,##0.00"), 13); _
                Spc(2); Mid(Numero_Lote, 1, 13); _
                Spc(15 - Len(Mid(Numero_Lote, 1, 13))); Fecha_Caducidad; _
                Spc(2); Numero_Acta; _
                Spc(7 - Len(Numero_Acta)); Comentarios_Libero; _
                Spc(22 - Len(Comentarios_Libero)); Usuario_Muestreo
            Grid_Reporte.AddItem Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("No_Entrada") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Estatus") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Codigo") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Nombre") _
                & Chr(9) & Existencia _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Numero_Lote") _
                & Chr(9) & Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                & Chr(9) & Numero_Acta _
                & Chr(9) & Comentarios_Libero _
                & Chr(9) & Usuario_Muestreo
            Rs_Consulta_Entradas.MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay datos que consultar", vbInformation
    End If
    Rs_Consulta_Entradas.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Reporte_Producto_Terminado
'DESCRIPCION: Consulta el producto terminado registrado en el sistema
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 19-Mayo-2010
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICACION:
'*******************************************************************************
Private Sub Reporte_Producto_Terminado()
Dim Numero_Acta As String
Dim Numero_Lote As String
Dim Fecha_Caducidad As String
Dim Rs_Consulta_Entradas As rdoResultset
Dim Existencia As Double

    'Consulta las facturas realizadas durante el día que se esta seleccionado
    Mi_SQL = "SELECT Detalles_Entradas.No_Entrada,Detalles_Entradas.Empresa,Ope_Entradas.Fecha,Detalles_Entradas.Estatus,Detalles_Entradas.Producto_ID,Detalles_Entradas.Cantidad,Detalles_Entradas.Faltante,Detalles_Entradas.Numero_Lote,Detalles_Entradas.Fecha_Caducidad,Detalles_Entradas.No_Acta,Cat_Productos.Nombre,Cat_Productos.Codigo"
    Mi_SQL = Mi_SQL & " FROM Ope_Entradas,Detalles_Entradas,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Ope_Entradas.No_Entrada=Detalles_Entradas.No_Entrada"
    Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Producto_ID=Cat_Productos.Producto_ID"
    If Chk_Producto.Value = 1 And Cmb_Producto_Venta.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Producto_ID='" & Format(Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.ListIndex), "00000") & "'"
    End If
    If Chk_Fechas.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Entradas.Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Historico_Ventas.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Historico_Ventas.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Numero_Lote.Value = 1 And Trim(Txt_Numero_Lote.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Numero_Lote LIKE '%" & Trim(Txt_Numero_Lote.Text) & "%'"
    End If
    If Chk_Fecha_Caducidad.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Fecha_Caducidad='" & Format(Dtp_Fecha_Caducidad.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Estatus_Materia_Prima.Value = 1 And Cmb_Estatus_Materia_Prima.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Estatus='" & Cmb_Estatus_Materia_Prima.Text & "'"
    End If
    If Chk_Numero_Acta.Value = 1 And Trim(Txt_Numero_Acta.Text) <> "" Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.No_Acta LIKE '%" & Trim(Txt_Numero_Acta.Text) & "%'"
    End If
    If InStr(1, Rol, "ALMACEN") Then
        Mi_SQL = Mi_SQL & " AND Detalles_Entradas.Faltante>0"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Entradas.Fecha"
    Set Rs_Consulta_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Entradas.EOF Then
        Call Encabezado_Reportes("REPORTE DE PRODUCTOS TERMINADOS", Format(DTP_Fecha_Inicio_Historico_Ventas, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Historico_Ventas, "dd/MMM/yyyy"), True)
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "  Fecha      Entrada    Estatus   Clave             Producto              Cantidad    No. Lote       Caducidad  No. Acta"
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Fecha" & Chr(9) & "Entrada" & Chr(9) & "Estatus" & Chr(9) & "Clave" & Chr(9) & "Producto" & Chr(9) & "Cantidad" & Chr(9) & "Numero Lote" & Chr(9) & "Fecha de Caducidad" & Chr(9) & "No. Acta"
        While Not Rs_Consulta_Entradas.EOF
            'Valida si tiene Lote y caducidad
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Numero_Lote")) Then
                Numero_Lote = Rs_Consulta_Entradas.rdoColumns("Numero_Lote")
            Else
                Numero_Lote = "             "
            End If
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad")) Then
                Fecha_Caducidad = Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy")
            Else
                Fecha_Caducidad = "           "
            End If
            If Not IsNull(Rs_Consulta_Entradas.rdoColumns("No_Acta")) Then
                Numero_Acta = Rs_Consulta_Entradas.rdoColumns("No_Acta")
            Else
                Numero_Acta = "               "
            End If
            If InStr(1, Rol, "ALMACEN") Then
                Existencia = Rs_Consulta_Entradas.rdoColumns("Faltante")
            Else
                Existencia = Rs_Consulta_Entradas.rdoColumns("Cantidad")
            End If
            Print #1, Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yy"); _
                Spc(1); Rs_Consulta_Entradas.rdoColumns("No_Entrada"); _
                Spc(1); Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 8); _
                Spc(10 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Estatus"), 1, 8))); Mid(Rs_Consulta_Entradas.rdoColumns("Codigo"), 1, 16); _
                Spc(18 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Codigo"), 1, 16))); Mid(Rs_Consulta_Entradas.rdoColumns("Nombre"), 1, 20); _
                Spc(21 - Len(Mid(Rs_Consulta_Entradas.rdoColumns("Nombre"), 1, 20))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Existencia, "#,###,##0.00"), 13); _
                Spc(2); Mid(Numero_Lote, 1, 13); _
                Spc(15 - Len(Mid(Numero_Lote, 1, 13))); Fecha_Caducidad; _
                Spc(2); Numero_Acta
            Grid_Reporte.AddItem Format(Rs_Consulta_Entradas.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("No_Entrada") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Estatus") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Codigo") _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Nombre") _
                & Chr(9) & Existencia _
                & Chr(9) & Rs_Consulta_Entradas.rdoColumns("Numero_Lote") _
                & Chr(9) & Format(Rs_Consulta_Entradas.rdoColumns("Fecha_Caducidad"), "dd/MMM/yyyy") _
                & Chr(9) & Numero_Acta
            Rs_Consulta_Entradas.MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay datos que consultar", vbInformation
    End If
    Rs_Consulta_Entradas.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Comparativo_Ventas
    'DESCRIPCIÓN: Hace el comparativo de ventas entre 2 años seleccionados por el
    '             usuario agrupados por mes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 05-Febrero-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Comparativo_Ventas()
Dim Encabezado As Boolean
Dim Total_Ventas As Double
Dim Total As Double
Dim Total_Ventas_1 As Double
Dim Total_1 As Double
Dim Porcentaje_Faltante As Double
Dim Pronostico_Ventas As Double
Dim Mi_SQL As String
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset
Dim Mes As Integer
    
    Call Encabezado_Reportes("REPORTE COMPARATIVO DE VENTAS", Format(Now, "dd/MMM/yyyy"), "", True)
    Print #1, Spc(48); "EMPRESA : "; Cmb_Empresa.Text
    Grid_Reporte.AddItem "EMPRESA : " & Chr(9) & Cmb_Empresa.Text
    If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "ZONA    : "; Cmb_Zona_Venta.Text
        Grid_Reporte.AddItem "ZONA : " & Chr(9) & Cmb_Zona_Venta.Text
    End If
    If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "ESTADO  : "; Cmb_Estado_Venta.Text
        Grid_Reporte.AddItem "ESTADO : " & Chr(9) & Cmb_Estado_Venta.Text
    End If
    If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "CIUDAD  : "; Cmb_Ciudad_Venta.Text
        Grid_Reporte.AddItem "CIUDAD : " & Chr(9) & Cmb_Ciudad_Venta.Text
    End If
    'FACTURAS
    Print #1,
    Print #1, Conectar_Ayudante.Centra("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "  Mes         Ventas "; Dtp_Año_Comparativo_1.Year; "        Ventas "; Dtp_Año_Comparativo_2.Year; "      % Avance    "; "     Pronostico"
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem "FACTURAS"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Mes" & Chr(9) & "Ventas " & Dtp_Año_Comparativo_1.Year & Chr(9) & "Ventas " & Dtp_Año_Comparativo_2.Year & Chr(9) & " % Avance" & Chr(9) & "Pronostico"
    'Hace el recorrido de los meses
    For Mes = 1 To 12
        'Consulta las ventas del primer año
        Mi_SQL = "SELECT SUM(Adm_Facturas_Clientes.Total) AS Total_Ventas"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND YEAR(Adm_Facturas_Clientes.Fecha)=" & Dtp_Año_Comparativo_1.Year
        Mi_SQL = Mi_SQL & " AND MONTH(Adm_Facturas_Clientes.Fecha)=" & Mes
        If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=" & Cmb_Empresa.ListIndex
        If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")) Then
                Total = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")
                Total_Ventas = Total_Ventas + Total
            Else
                Total = 0
            End If
        Else
            Total = 0
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        'Consulta las ventas del segundo año
        Mi_SQL = "SELECT SUM(Adm_Facturas_Clientes.Total) AS Total_Ventas"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND YEAR(Adm_Facturas_Clientes.Fecha)=" & Dtp_Año_Comparativo_2.Year
        Mi_SQL = Mi_SQL & " AND MONTH(Adm_Facturas_Clientes.Fecha)=" & Mes
        If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=" & Cmb_Empresa.ListIndex
        If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")) Then
                Total_1 = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")
                Total_Ventas_1 = Total_Ventas_1 + Total_1
            Else
                Total_1 = 0
            End If
        Else
            Total_1 = 0
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        'Obtiene el porcentaje faltante
        If Total <> 0 Then
            Porcentaje_Faltante = Val(Total_1 * 100 / Total)
        Else
            Porcentaje_Faltante = 0
        End If
        Print #1, Spc(1); Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM"); Spc(15 - Len(Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_1, "#,###,##0.00"), 14); _
            Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Porcentaje_Faltante, "#0.00"), 14)
        Grid_Reporte.AddItem Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM") & Chr(9) & Total & Chr(9) & Total_1 & Chr(9) & Porcentaje_Faltante
    Next Mes
    'Obtiene el porcentaje faltante
    If Total_Ventas <> 0 Then
        Porcentaje_Faltante = Val(Total_Ventas_1 * 100 / Total_Ventas)
    Else
        Porcentaje_Faltante = 0
    End If
    'Obtiene el pronóstico
    If Val(DateDiff("d", "01/01/" & Year(Dtp_Año_Comparativo_2.Value), Format(Now, "MM/dd") & "/" & Dtp_Año_Comparativo_2.Year)) > 0 Then
        Pronostico_Ventas = Val(300 * Total_Ventas_1 / DateDiff("d", "01/01/" & Year(Dtp_Año_Comparativo_2.Value), Format(Now, "MM/dd") & "/" & Dtp_Año_Comparativo_2.Year))
    Else
        Pronostico_Ventas = 0
    End If
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(1); "ANUAL"; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,##0.00"), 14); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas_1, "#,###,##0.00"), 14); _
        Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Porcentaje_Faltante, "#0.00"), 14); Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Pronostico_Ventas, "#,###,##0.00"), 14)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "ANUAL" & Chr(9) & Total_Ventas & Chr(9) & Total_Ventas_1 & Chr(9) & Porcentaje_Faltante & Chr(9) & Pronostico_Ventas
    'REMISIONES
    Total_Ventas = 0
    Total = 0
    Total_Ventas_1 = 0
    Porcentaje_Faltante = 0
    Pronostico_Ventas = 0
    Print #1,
    Print #1, Conectar_Ayudante.Centra("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "  Mes         Ventas "; Dtp_Año_Comparativo_1.Year; "        Ventas "; Dtp_Año_Comparativo_2.Year; "      % Avance    "; "     Pronostico"
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "REMISIONES"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Mes" & Chr(9) & "Ventas " & Dtp_Año_Comparativo_1.Year & Chr(9) & "Ventas " & Dtp_Año_Comparativo_2.Year & Chr(9) & " % Avance" & Chr(9) & "Pronostico"
    'Hace el recorrido de los meses
    For Mes = 1 To 12
        'Consulta las ventas del primer año
        Mi_SQL = "SELECT SUM(Adm_Remisiones_Clientes.Total) AS Total_Ventas"
        Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND YEAR(Adm_Remisiones_Clientes.Fecha)=" & Dtp_Año_Comparativo_1.Year
        Mi_SQL = Mi_SQL & " AND MONTH(Adm_Remisiones_Clientes.Fecha)=" & Mes
        If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Empresa=" & Cmb_Empresa.ListIndex
        If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")) Then
                Total = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")
                Total_Ventas = Total_Ventas + Total
            Else
                Total = 0
            End If
        Else
            Total = 0
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        'Consulta las ventas del segundo año
        Mi_SQL = "SELECT SUM(Adm_Remisiones_Clientes.Total) AS Total_Ventas"
        Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes"
        Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
        Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cancelada='N'"
        Mi_SQL = Mi_SQL & " AND YEAR(Adm_Remisiones_Clientes.Fecha)=" & Dtp_Año_Comparativo_2.Year
        Mi_SQL = Mi_SQL & " AND MONTH(Adm_Remisiones_Clientes.Fecha)=" & Mes
        If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Empresa=" & Cmb_Empresa.ListIndex
        If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
        If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
        Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")) Then
                Total_1 = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total_Ventas")
                Total_Ventas_1 = Total_Ventas_1 + Total_1
            Else
                Total_1 = 0
            End If
        Else
            Total_1 = 0
        End If
        Rs_Consulta_Adm_Facturas_Clientes.Close
        'Obtiene el porcentaje faltante
        If Total <> 0 Then
            Porcentaje_Faltante = Val(Total_1 * 100 / Total)
        Else
            Porcentaje_Faltante = 0
        End If
        Print #1, Spc(1); Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM"); Spc(15 - Len(Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_1, "#,###,##0.00"), 14); _
            Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Porcentaje_Faltante, "#0.00"), 14)
        Grid_Reporte.AddItem Format((Format(Mes, "00") & "/01/" & Dtp_Año_Comparativo_2.Year), "MMMM") & Chr(9) & Total & Chr(9) & Total_1 & Chr(9) & Porcentaje_Faltante
    Next Mes
    'Obtiene el porcentaje faltante
    If Total_Ventas <> 0 Then
        Porcentaje_Faltante = Val(Total_Ventas_1 * 100 / Total_Ventas)
    Else
        Porcentaje_Faltante = 0
    End If
    'Obtiene el pronóstico
    If Val(DateDiff("d", "01/01/" & Year(Dtp_Año_Comparativo_2.Value), Format(Now, "MM/dd") & "/" & Dtp_Año_Comparativo_2.Year)) > 0 Then
        Pronostico_Ventas = Val(300 * Total_Ventas_1 / DateDiff("d", "01/01/" & Year(Dtp_Año_Comparativo_2.Value), Format(Now, "MM/dd") & "/" & Dtp_Año_Comparativo_2.Year))
    Else
        Pronostico_Ventas = 0
    End If
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(1); "ANUAL"; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,##0.00"), 14); Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas_1, "#,###,##0.00"), 14); _
        Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Porcentaje_Faltante, "#0.00"), 14); Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Pronostico_Ventas, "#,###,##0.00"), 14)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "ANUAL" & Chr(9) & Total_Ventas & Chr(9) & Total_Ventas_1 & Chr(9) & Porcentaje_Faltante & Chr(9) & Pronostico_Ventas
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Detallado_Ventas
    'DESCRIPCIÓN: Muestra el reporte de ventas según los filtros seleccionados
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 11-Febrero-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Detallado_Ventas()
Dim Total_Ventas As Double
Dim Empresa As String
Dim Estatus As String
Dim Mi_SQL As String
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset
    
    Call Encabezado_Reportes("REPORTE DETALLADO DE VENTAS", Format(Dtp_Año_Comparativo_1.Value, "dd/MMM/yyyy"), Format(Dtp_Año_Comparativo_2.Value, "dd/MMM/yyyy"), True)
    Print #1, Spc(48); "EMPRESA : "; Cmb_Empresa.Text
    Grid_Reporte.AddItem "EMPRESA : " & Chr(9) & Cmb_Empresa.Text
    If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "ZONA    : "; Cmb_Zona_Venta.Text
        Grid_Reporte.AddItem "ZONA : " & Chr(9) & Cmb_Zona_Venta.Text
    End If
    If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "ESTADO  : "; Cmb_Estado_Venta.Text
        Grid_Reporte.AddItem "ESTADO : " & Chr(9) & Cmb_Estado_Venta.Text
    End If
    If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then
        Print #1, Spc(48); "CIUDAD  : "; Cmb_Ciudad_Venta.Text
        Grid_Reporte.AddItem "CIUDAD : " & Chr(9) & Cmb_Ciudad_Venta.Text
    End If
    'Consulta las facturas del periodo seleccionado
    Mi_SQL = "SELECT Adm_Facturas_Clientes.*,Cat_Clientes.Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha BETWEEN '" & Format(Dtp_Año_Comparativo_1.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Año_Comparativo_2.Value, "MM/dd/yyyy") & "'"
    If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=" & Cmb_Empresa.ListIndex
    If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
    If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
    If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Facturas_Clientes.Fecha"
    Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Print #1,
    Print #1, Conectar_Ayudante.Centra("FACTURAS", Len("------------------------------------------------------------------------------------------------------------------------"))
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "     Fecha       No. Factura       Estatus             Cliente                                                Total     "
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Fecha" & Chr(9) & "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Cliente" & Chr(9) & "Total"
    If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
        With Rs_Consulta_Adm_Facturas_Clientes
            While Not .EOF
                If .rdoColumns("Empresa") = 1 Then  'Valida la empresa
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                If .rdoColumns("Cancelada") = "S" Then  'Valida el estatus
                    Estatus = "CANCELADA"
                Else
                    If .rdoColumns("Pagada") = "S" Then
                        Estatus = "PAGADA   "
                    Else
                        Estatus = "PENDIENTE"
                    End If
                    Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                End If
                Print #1, Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(4); .rdoColumns("No_Factura") & " " & Empresa; _
                    Spc(4); Estatus; _
                    Spc(4); Mid(.rdoColumns("Nombre"), 1, 50); _
                    Spc(55 - Len(Mid(.rdoColumns("Nombre"), 1, 50))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14)
                Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("No_Factura") & Chr(9) & Empresa _
                    & Chr(9) & Estatus _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Total")
                .MoveNext
            Wend
        End With
    End If
    Rs_Consulta_Adm_Facturas_Clientes.Close
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(60); "TOTAL VENTAS"; _
        Spc(29); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,###,##0.00"), 16)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "" _
        & Chr(9) & "" & Chr(9) & "" _
        & Chr(9) & "" _
        & Chr(9) & "TOTAL DE VENTAS" _
        & Chr(9) & Total_Ventas
    'Consulta las remisiones del periodo seleccionado
    Total_Ventas = 0
    Mi_SQL = "SELECT Adm_Remisiones_Clientes.*,Cat_Clientes.Nombre"
    Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Adm_Remisiones_Clientes.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Cancelada='N'"
    Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Fecha BETWEEN '" & Format(Dtp_Año_Comparativo_1.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Año_Comparativo_2.Value, "MM/dd/yyyy") & "'"
    If Cmb_Empresa.ListIndex > 0 Then Mi_SQL = Mi_SQL & " AND Adm_Remisiones_Clientes.Empresa=" & Cmb_Empresa.ListIndex
    If Chk_Comparativo_Zona.Value = 1 And Cmb_Zona_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Zona_ID='" & Format(Cmb_Zona_Venta.ItemData(Cmb_Zona_Venta.ListIndex), "00000") & "'"
    If Chk_Comparativo_Estado.Value = 1 And Cmb_Estado_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Estado_ID='" & Format(Cmb_Estado_Venta.ItemData(Cmb_Estado_Venta.ListIndex), "00000") & "'"
    If Chk_Comparativo_Ciudad.Value = 1 And Cmb_Ciudad_Venta.ListIndex > -1 Then Mi_SQL = Mi_SQL & " AND Cat_Clientes.Ciudad_ID='" & Format(Cmb_Ciudad_Venta.ItemData(Cmb_Ciudad_Venta.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Adm_Remisiones_Clientes.Fecha"
    Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Print #1,
    Print #1, Conectar_Ayudante.Centra("REMISIONES", Len("------------------------------------------------------------------------------------------------------------------------"))
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, "     Fecha       No. Remision      Estatus             Cliente                                                Total     "
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Fecha" & Chr(9) & "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Estatus" & Chr(9) & "Cliente" & Chr(9) & "Total"
    If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
        With Rs_Consulta_Adm_Facturas_Clientes
            While Not .EOF
                If .rdoColumns("Empresa") = 1 Then  'Valida la empresa
                    Empresa = "NHE"
                Else
                    Empresa = "PNC"
                End If
                If .rdoColumns("Cancelada") = "S" Then  'Valida el estatus
                    Estatus = "CANCELADA"
                Else
                    If .rdoColumns("Pagada") = "S" Then
                        Estatus = "PAGADA   "
                    Else
                        Estatus = "PENDIENTE"
                    End If
                    Total_Ventas = Total_Ventas + Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                End If
                Print #1, Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(4); .rdoColumns("No_Remision") & " " & Empresa; _
                    Spc(4); Estatus; _
                    Spc(4); Mid(.rdoColumns("Nombre"), 1, 50); _
                    Spc(55 - Len(Mid(.rdoColumns("Nombre"), 1, 50))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14)
                Grid_Reporte.AddItem Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("No_Remision") & Chr(9) & Empresa _
                    & Chr(9) & Estatus _
                    & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Total")
                .MoveNext
            Wend
        End With
    End If
    Rs_Consulta_Adm_Facturas_Clientes.Close
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(60); "TOTAL VENTAS"; _
        Spc(29); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,###,##0.00"), 16)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "" _
        & Chr(9) & "" & Chr(9) & "" _
        & Chr(9) & "" _
        & Chr(9) & "TOTAL DE VENTAS" _
        & Chr(9) & Total_Ventas
    Call Encabezado_Reportes("", "", "", False)
End Sub

Private Sub Btn_Reporte_Envios_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Medios As rdoResultset          'Manejo de registro
Dim Rs_Consulta_Operadores As rdoResultset      'Manejo de registro
Dim Medio As String                             'Almacena el nombre del medio
Dim Estatus As String                           'Almacena el estatus de la factura
Dim Operador As String                          'Almacena el nombre del operador
Dim Gastos As Double                            'Alamcena el valor de los gastos adicionales
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Empresa As String                           'Almacena el nombre de la empresa
Dim Total_General As Double                     'Almacena el total del producto vendido
Dim Total_Producto As Double                    'Almacena el total de los productos pedidos
Dim Piezas_Pedido As Long
Dim Piezas_Factura As Long
Dim Rs_Consulta_Ope_Pedidos As rdoResultset     'Manejo de registro

    'Consulta la tabla de pedidos
    Mi_SQL = "SELECT DISTINCT Ope_Pedidos.No_Pedido,Ope_Pedidos.Fecha_Pedido,Ope_Salidas.No_Salida,Adm_Facturas_Clientes.No_Factura,Adm_Facturas_Clientes.Empresa,Adm_Facturas_Clientes.Fecha AS Fecha_Factura,Adm_Facturas_Clientes.Total,Adm_Facturas_Clientes.Pagada,Adm_Facturas_Clientes.Cancelada"
    Mi_SQL = Mi_SQL & " ,Ope_Pedidos.Cliente_ID,Cat_Clientes.Nombre AS Cliente,Ope_Salidas.Fecha AS Fecha_Salida,Ope_Salidas.Nombre_Surtio,Ope_Salidas.Medio_ID,Ope_Salidas.Costo_Envio,Ope_Salidas.Fecha_Envio,Ope_Salidas.Hora_Envio,Ope_Salidas.Gastos_Adicionales,Ope_Salidas.Operador_ID,Ope_Salidas.Lugar_Envio,Ope_Salidas.Observaciones"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Ope_Salidas,Detalles_Salidas,Adm_Facturas_Clientes,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.No_Pedido=Ope_Salidas.No_Pedido"
    Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Cliente_ID=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Detalles_Salidas.No_Factura=Adm_Facturas_Clientes.No_Factura"
    Mi_SQL = Mi_SQL & " AND Detalles_Salidas.Empresa_Factura=Adm_Facturas_Clientes.Empresa"
    Mi_SQL = Mi_SQL & " AND Ope_Salidas.No_Salida=Detalles_Salidas.No_Salida"
    If Chk_Cliente.Value = 1 And Cmb_Cliente_Envio.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Cliente_ID='" & Format(Cmb_Cliente_Envio.ItemData(Cmb_Cliente_Envio.ListIndex), "00000") & "'"
    End If
    If Chk_No_Pedido_Envio.Value = 1 And Txt_No_Pedido_Envio.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND Ope_Pedidos.No_Pedido='" & Format(Txt_No_Pedido_Envio.Text, "0000000000") & "'"
    End If
    If Chk_No_Factura_Envio.Value = 1 And Txt_No_Factura_Envio.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.No_Factura='" & Format(Txt_No_Factura_Envio.Text, "0000000000") & "'"
    End If
    If Chk_Medio_Envio.Value = 1 And Cmb_Medio_Envio.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.Medio_ID='" & Format(Cmb_Medio_Envio.ItemData(Cmb_Medio_Envio.ListIndex), "00000") & "'"
    End If
    If Chk_Fecha_Envio.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.Fecha_Envio BETWEEN '" & Format(Dtp_Fecha_Envio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Envio_Fin.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Fecha_Pedido.Value = 1 Then
        Mi_SQL = Mi_SQL & " AND Ope_Pedidos.Fecha_Pedido BETWEEN '" & Format(Dtp_Fecha_Inicio_Envio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Fin_Envio.Value, "MM/dd/yyyy") & "'"
    End If
    If Chk_Lugar_Envio.Value = 1 And Txt_Lugar_Envio.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND Ope_Salidas.Lugar_Envio LIKE '%" & Trim(Txt_Lugar_Envio.Text) & "%'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.No_Pedido"
    Set Rs_Consulta_Ope_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Ope_Pedidos.EOF Then
        'Coloca el encabezado de los reportes
        Grid_Reporte.Rows = 0
        Grid_Reporte.Cols = 20
        Call Encabezado_Reportes("REPORTE ENVIO DE PEDIDOS", Format(Dtp_Fecha_Inicio_Envio.Value, "dd/MMM/yyyy"), "", True)
        Grid_Reporte.AddItem "Pedido" & Chr(9) & "Fecha Pedido" & Chr(9) & "# Piezas" & Chr(9) & "Cliente" & Chr(9) & "Fecha Surtio" & Chr(9) & "Nombre Surtio" & Chr(9) & "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha Factura" & Chr(9) & "# Facturadas" & Chr(9) & "Importe" & Chr(9) & "Estatus" & Chr(9) & "Medio Envio" & Chr(9) & "Costo Envio" & Chr(9) & "Fecha Envio" & Chr(9) & "Hora Envio" & Chr(9) & "Gastos Adicionales" & Chr(9) & "Operador" & Chr(9) & "Lugar" & Chr(9) & "Observaciones"
        With Rs_Consulta_Ope_Pedidos
            While Not .EOF
                'Estatus Factura
                If Rs_Consulta_Ope_Pedidos.rdoColumns("Cancelada") = "S" Then
                    Estatus = "Cancelada"
                Else
                    If Rs_Consulta_Ope_Pedidos.rdoColumns("Pagada") = "S" Then
                        Estatus = "Pagada"
                    Else
                        Estatus = "Sin Pagar"
                    End If
                End If
                'Empresa
                If Not IsNull(.rdoColumns("Empresa")) Then
                    If .rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                Else
                    Empresa = ""
                End If
                If Not IsNull(.rdoColumns("Gastos_Adicionales")) Then
                    Gastos = Val(.rdoColumns("Gastos_Adicionales"))
                Else
                    Gastos = 0
                End If
                'Medio de envío
                If Not IsNull(.rdoColumns("Medio_ID")) Then
                    Mi_SQL = "SELECT Nombre,Medio_ID FROM Cat_Medios_Envio"
                    Mi_SQL = Mi_SQL & " WHERE Medio_ID='" & Rs_Consulta_Ope_Pedidos.rdoColumns("Medio_ID") & "'"
                    Set Rs_Consulta_Medios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Medios.EOF Then
                        Medio = Rs_Consulta_Medios.rdoColumns("Nombre")
                    Else
                        Medio = ""
                    End If
                    Rs_Consulta_Medios.Close
                Else
                    Medio = ""
                End If
                'Operador
                If Not IsNull(.rdoColumns("Operador_ID")) Then
                    Mi_SQL = "SELECT Nombre,Operador_ID FROM Cat_Operadores"
                    Mi_SQL = Mi_SQL & " WHERE Operador_ID='" & Rs_Consulta_Ope_Pedidos.rdoColumns("Operador_ID") & "'"
                    Set Rs_Consulta_Operadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Operadores.EOF Then
                        Operador = Rs_Consulta_Operadores.rdoColumns("Nombre")
                    Else
                        Operador = ""
                    End If
                    Rs_Consulta_Operadores.Close
                Else
                    Operador = ""
                End If
                'Obtiene la cantidad de piezas del pedido
                Mi_SQL = "SELECT SUM(Cantidad) AS Total_Unidades FROM Detalles_Pedidos"
                Mi_SQL = Mi_SQL & " WHERE No_Pedido='" & .rdoColumns("No_Pedido") & "'"
                Set Rs_Consulta_Medios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Medios.EOF Then
                    If Not IsNull(Rs_Consulta_Medios.rdoColumns("Total_Unidades")) Then
                        Piezas_Pedido = Val(Rs_Consulta_Medios.rdoColumns("Total_Unidades"))
                    Else
                        Piezas_Pedido = 0
                    End If
                Else
                    Piezas_Pedido = 0
                End If
                Rs_Consulta_Medios.Close
                'Obtiene la cantidad de piezas de la factura
                Mi_SQL = "SELECT SUM(Cantidad) AS Total_Unidades FROM Adm_Detalles_Facturas"
                Mi_SQL = Mi_SQL & " WHERE No_Factura='" & .rdoColumns("No_Factura") & "'"
                Mi_SQL = Mi_SQL & " AND Empresa=" & .rdoColumns("Empresa")
                Set Rs_Consulta_Medios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Medios.EOF Then
                    If Not IsNull(Rs_Consulta_Medios.rdoColumns("Total_Unidades")) Then
                        Piezas_Factura = Val(Rs_Consulta_Medios.rdoColumns("Total_Unidades"))
                    Else
                        Piezas_Factura = 0
                    End If
                Else
                    Piezas_Factura = 0
                End If
                Rs_Consulta_Medios.Close
                If No_Pedido <> .rdoColumns("No_Pedido") Then
                    'Imprime los pedidos
                    Print #1,
                    Print #1, "______________________________________________________________________________________________________________________"
                    Print #1, "NO PEDIDO:"; .rdoColumns("No_Pedido"); Spc(5); "FECHA PEDIDO: "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy"); Spc(3); "CLIENTE: "; Spc(1); .rdoColumns("Cliente")
                    Print #1, "______________________________________________________________________________________________________________________"
                    No_Pedido = .rdoColumns("No_Pedido")
                End If
                Print #1, "No.Factura:"; Conectar_Ayudante.Alinea_Derecha_Derecha(.rdoColumns("No_Factura"), 7); Spc(4); Empresa; Spc(4); "Importe:"; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Total")), "#,###,##0.00"), 7); Spc(2); Estatus
                Print #1, "Medio Envio:"; Medio; Spc(30 - Len(Medio)); "Costo Envio:"; Rs_Consulta_Ope_Pedidos.rdoColumns("Costo_Envio"); Spc(3); "Fecha Envio:"; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Envio"), "dd/MMM/yyyy"); Spc(3); "Hora Envio:"; Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Hora_Envio"), "HH:mm am/pm")
                Print #1, "Gastos Adicionales:"; Gastos; Spc(3); "Operador:"; Operador; Spc(25 - Len(Operador)); "Lugar:"; Trim(Rs_Consulta_Ope_Pedidos.rdoColumns("Lugar_Envio"))
                Print #1, "Observaciones:"; Rs_Consulta_Ope_Pedidos.rdoColumns("Observaciones")
                Print #1, "----------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem .rdoColumns("No_Pedido") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy") _
                    & Chr(9) & Piezas_Pedido _
                    & Chr(9) & .rdoColumns("Cliente") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Salida"), "dd/MMM/yyyy") _
                    & Chr(9) & .rdoColumns("Nombre_Surtio") _
                    & Chr(9) & .rdoColumns("No_Factura") _
                    & Chr(9) & Empresa _
                    & Chr(9) & Format(.rdoColumns("Fecha_Factura"), "dd/MMM/yyyy") _
                    & Chr(9) & Piezas_Factura _
                    & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") _
                    & Chr(9) & Estatus _
                    & Chr(9) & Medio _
                    & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Costo_Envio"), "#,###,##0.00") _
                    & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Fecha_Envio"), "dd/MM/yyyy") _
                    & Chr(9) & Format(Rs_Consulta_Ope_Pedidos.rdoColumns("Hora_Envio"), "HH:mm am/pm") _
                    & Chr(9) & Format(Gastos, "#,###,##0.00") _
                    & Chr(9) & Operador _
                    & Chr(9) & Trim(Rs_Consulta_Ope_Pedidos.rdoColumns("Lugar_Envio")) _
                    & Chr(9) & Conectar_Ayudante.Quitar_Caracter(Conectar_Ayudante.Quitar_Caracter(Trim(Rs_Consulta_Ope_Pedidos.rdoColumns("Observaciones")), Chr(13)), Chr(10))
                Total_General = Total_General + .rdoColumns("Total")
                .MoveNext
            Wend
        End With
        Rs_Consulta_Ope_Pedidos.Close
        Print #1, "------------------------------------------------------------------------------------------------------------------------------------"
        Print #1, Spc(30); "GRAN TOTAL:       "; Format(Total_General, "#,###,##0.00")
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "GRAN TOTAL:     " & Chr(9) & Format(Total_General, "#,###,##0.00")
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay datos con los criterios buscados", vbInformation
    End If
End Sub

Private Sub Reporte_Facturas_Directas()
Dim Rs_Consulta_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Detalles_Salidas As rdoResultset     'Manejo de registro
Dim Rs_Detalles_Entradas As rdoResultset    'Manejo de registro
Dim Rs_Productos As rdoResultset            'Manejo de registro
Dim Cliente As String                       'Almacena el nombre del cliente
Dim Total_Costos As Double                  'Almacena el total de los costos por salida
Dim Total_Venta As Double                   'Almacena el total de la venta por salida
Dim Total_Utilidad As Double                'Almacena el total de la utilidad por salida
Dim Gran_Total_Costos As Double             'Almacena el gran total de los costos de las salidas
Dim Gran_Total_Venta As Double              'Almacena el gran total de la venta por las salidas
Dim Gran_Total_Utilidad As Double           'Almacena el gran total de la utilidades de las salidas
Dim Empresa As String                       'Almacena el nombre de la empresa

    'Consulta las salidas de acuerdo a las fechas proporcionadas
    Mi_SQL = "SELECT No_Salida,Fecha,Tipo,No_Pedido,Ope_Salidas.Cliente_ID,Cat_Clientes.Cliente_ID,Cat_Clientes.Nombre as Cliente,Ope_Salidas.Observaciones"
    Mi_SQL = Mi_SQL & " FROM Ope_Salidas,Cat_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Ope_Salidas.Cliente_Id=Cat_Clientes.Cliente_ID"
    Mi_SQL = Mi_SQL & " AND Tipo<>'PEDIDO'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Salida"
    Set Rs_Consulta_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reportes("REPORTE DE FACTURAS DIRECTAS", Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem "No Salida" & Chr(9) & "Codigo Barras" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Costo Compra" & Chr(9) & " Total Compra" & Chr(9) & "Precio Venta" & Chr(9) & "Total Venta"
    Gran_Total_Costos = 0
    Gran_Total_Venta = 0
    Gran_Total_Utilidad = 0
    While Not Rs_Consulta_Salidas.EOF
        If Rs_Consulta_Salidas.rdoColumns("Cliente_ID") Then
            Cliente = Rs_Consulta_Salidas.rdoColumns("Cliente")
        Else
            Cliente = ""
        End If
        Print #1,
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Print #1, "Fecha Salida : "; Format(Rs_Consulta_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy"); Spc(3); "Cliente:"; Cliente
        Print #1, "------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem "Fecha Salida : " & Chr(9) & Format(Rs_Consulta_Salidas.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & "Cliente:" & Chr(9) & Cliente
        'Consulta la descripcion de las salidas
        Mi_SQL = "SELECT * FROM Adm_Detalles_Facturas"
        Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Trim(Mid(Rs_Consulta_Salidas.rdoColumns("Observaciones"), 22, 11)) & "'"
        Set Rs_Detalles_Salidas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Total_Costos = 0
        Total_Venta = 0
        With Rs_Detalles_Salidas
            Print #1, "No Factura   Codigo Barras    Descripcion                      Empresa     Cantidad        Importe"
            Print #1, "__________________________________________________________________________________________________________________"
            Print #1,
            While Not .EOF
            
            If Rs_Detalles_Salidas.rdoColumns("Empresa") = 1 Then
                Empresa = "NHE"
            Else
                Empresa = "PNC"
            End If
                'Consulta  la descripcion del producto
'                Mi_SQL = "SELECT Producto_ID,Codigo,Nombre,Costo"
'                Mi_SQL = Mi_SQL & " FROM Cat_Productos"
'                Mi_SQL = Mi_SQL & " WHERE Nombre='" & .rdoColumns("Producto") & "'"
'                Set Rs_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1, Trim(Mid(Rs_Consulta_Salidas.rdoColumns("Observaciones"), 22, 11)); Spc(3); Mid(.rdoColumns("Codigo"), 1, 17); Spc(18 - Len(Mid(.rdoColumns("Codigo"), 1, 17))); Mid(.rdoColumns("Producto"), 1, 32); Spc(33 - Len(Mid(.rdoColumns("Producto"), 1, 32))); Empresa; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "#,###"), 5); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 10)
                Grid_Reporte.AddItem Trim(Mid(Rs_Consulta_Salidas.rdoColumns("Observaciones"), 22, 11)) & Chr(9) & .rdoColumns("Codigo") & Chr(9) & .rdoColumns("Producto") & Chr(9) & Empresa & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Importe")

                'Calculas los totales de cada salida
                Total_Costos = Total_Costos + Val(.rdoColumns("Importe"))
'                Total_venta = Total_venta + Val(.rdoColumns("Precio_Venta") * .rdoColumns("Cantidad"))
'                Rs_Productos.Close
                .MoveNext
            Wend
        End With
        Print #1, Spc(60); " __________________________________________________________ "
        Print #1, Spc(60); "Total Salidas: "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Costos, "#,###,###.00"), 12)
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Salidas: " & Chr(9) & Total_Costos
        Grid_Reporte.AddItem ""
        'Calcula el gran total de todas las salidas
        Gran_Total_Costos = Gran_Total_Costos + Total_Costos
'        Gran_Total_Venta = Gran_Total_Venta + Total_venta
        Rs_Detalles_Salidas.Close
        Rs_Consulta_Salidas.MoveNext
    Wend
    Print #1,
    Print #1, "------------------------------------------------------------------------------------------------------------------------"
    Print #1, Spc(60); "Gran Total   : "; Spc(10); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total_Costos, "#,###,###.00"), 12)
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem Chr(9) & Chr(9) & "Gran Total : " & Chr(9) & Gran_Total_Costos
    Rs_Consulta_Salidas.Close
    Call Encabezado_Reportes("", "", "", False)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Presencia_Productos
    'DESCRIPCIÓN: Muestra un reporte en el que se checa el porcentaje de
    '             presencia de productos con el total de clientes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Agosto-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Reporte_Presencia_Productos()
Dim Mi_SQL As String
Dim Rs_Total_Clientes As rdoResultset
Dim Rs_Cat_Productos As rdoResultset
Dim Rs_Consulta_Ventas_Productos As rdoResultset
Dim Total_Clientes As Long
Dim Porcentaje_Presencia As Double

    'Calcula el total de clientes en el sistema con compra
    Mi_SQL = "SELECT COUNT(DISTINCT(Cliente_ID)) AS Total_Clientes FROM Adm_Facturas_Clientes"
    Mi_SQL = Mi_SQL & " WHERE Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
    Set Rs_Total_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Total_Clientes.EOF Then
        If Not IsNull(Rs_Total_Clientes.rdoColumns("Total_Clientes")) Then
            Total_Clientes = Rs_Total_Clientes.rdoColumns("Total_Clientes")
        Else
            Total_Clientes = 0
            Rs_Total_Clientes.Close
            MsgBox "No tiene clientes registrados", vbExclamation
            Exit Sub
        End If
    Else
        Total_Clientes = 0
        Rs_Total_Clientes.Close
        MsgBox "No tiene clientes registrados", vbExclamation
        Exit Sub
    End If
    Rs_Total_Clientes.Close
    'Recorre el catálogo de productos
    Mi_SQL = "SELECT Producto_ID,Empresa,Codigo,Nombre FROM Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Tipo_Producto_ID<>'00016'"
    Mi_SQL = Mi_SQL & " ORDER BY Empresa,Nombre"
    Set Rs_Cat_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Cat_Productos.EOF Then
        Call Encabezado_Reportes("REPORTE DE PRESENCIA DE PRODUCTOS", Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "dd/MMM/yyyy"), True)
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, "                                                                                   Clientes            Porcentaje        "
        Print #1, "   Codigo                 Descripcion                                           Total      Compra      Presencia         "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Total Clientes" & Chr(9) & "Total Compra" & Chr(9) & "Porcentaje de Presencia"
        Grid_Reporte.AddItem ""
        While Not Rs_Cat_Productos.EOF
            'Revisa en cuantas ventas existe el producto separado por cliente
            Mi_SQL = "SELECT COUNT(DISTINCT Adm_Facturas_Clientes.Cliente_ID) AS Presencia"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Adm_Detalles_Facturas"
            Mi_SQL = Mi_SQL & " WHERE Adm_Facturas_Clientes.No_Factura=Adm_Detalles_Facturas.No_Factura"
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Empresa=Adm_Detalles_Facturas.Empresa"
            Mi_SQL = Mi_SQL & " AND Adm_Detalles_Facturas.Producto_ID='" & Rs_Cat_Productos.rdoColumns("Producto_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Fecha<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Facturas_Directas.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Cancelada='N'"
            Set Rs_Consulta_Ventas_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Ventas_Productos.EOF Then
                If Not IsNull(Rs_Consulta_Ventas_Productos.rdoColumns("Presencia")) Then
                    Porcentaje_Presencia = Rs_Consulta_Ventas_Productos.rdoColumns("Presencia") / Total_Clientes * 100
                Else
                    Porcentaje_Presencia = 0
                End If
            Else
                Porcentaje_Presencia = 0
            End If
            Print #1, Spc(1); Mid(Rs_Cat_Productos.rdoColumns("Codigo"), 1, 20); Spc(22 - Len(Mid(Rs_Cat_Productos.rdoColumns("Codigo"), 1, 20))); Mid(Rs_Cat_Productos.rdoColumns("Nombre"), 1, 50); Spc(52 - Len(Mid(Rs_Cat_Productos.rdoColumns("Nombre"), 1, 50))); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Clientes, "####0"), 5); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Ventas_Productos.rdoColumns("Presencia"), "####0"), 5); Spc(7); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Porcentaje_Presencia, "#0.00"), 7)
            Grid_Reporte.AddItem Rs_Cat_Productos.rdoColumns("Codigo") & Chr(9) & Rs_Cat_Productos.rdoColumns("Nombre") & Chr(9) & Total_Clientes & Chr(9) & Rs_Consulta_Ventas_Productos.rdoColumns("Presencia") & Chr(9) & Format(Porcentaje_Presencia, "#0.00")
            Rs_Cat_Productos.MoveNext
        Wend
        Call Encabezado_Reportes("", "", "", False)
    Else
        MsgBox "No hay productos que mostrar", vbExclamation
    End If
    Rs_Cat_Productos.Close
End Sub

Private Sub Btn_Reporte_Facturas_Directas_Click()
    If Fra_Facturas_Directas.Caption = "Facturas Directas" Then
        Reporte_Facturas_Directas
    Else
        If Fra_Facturas_Directas.Caption = "Presencia de Productos en el Mercado" Then
            Reporte_Presencia_Productos
        Else
            If Cmb_Factura_Remision.ListIndex = 0 Then
                Cartera_Vencida
            Else
                Cartera_Vencida_Remisiones
            End If
        End If
    End If
End Sub

Private Sub Btn_Reporte_Facturas_Vencidas_Click()
    Select Case Fra_Facturas_Vencidas.Caption
        Case "Facturas Vencidas"
            Reporte_Facturas_Vencidas
        Case "Facturas por Vencer"
            Reporte_Facturas_Vencer
        Case "Facturas por Tipo"
            Reporte_Tipo_Factura
    End Select
End Sub

Private Sub Btn_Reporte_Gastos_Click()
    If Cmb_Banco.ListIndex > -1 And Txt_No_Cheque.Text <> "" Then
        Reporte_Gastos
    Else
        MsgBox "Falta seleccionar el banco de la lista o proporcinar un número de cheque", vbExclamation
    End If
End Sub

Private Sub Btn_Reporte_Historico_Facturas_Click()
    Select Case Fra_Historico_Facturas.Caption
        Case "General Facturas"
            Reporte_General_Facturas
        Case "General Remisiones"
            Reporte_General_Remisiones
        Case "Facturas Pagadas"
            Reporte_Facturas_Pagadas
        Case "Facturas Canceladas"
            Reporte_Facturas_Canceladas
        Case "Facturas Cliente"
            Reporte_Facturas_Clientes
        Case "Remisiones Cliente"
            Reporte_Remisiones_Clientes
        Case "General Facturas Proveedor"
            Reporte_Facturas_Proveedores
    End Select
End Sub

Private Sub Btn_Reporte_Historico_Ventas_Click()
    Select Case Fra_Reporte_Ventas.Caption
        Case "Ventas por Cliente"
            Reporte_Ventas_Clientes
        Case "Distribución de Producto"
            Reporte_Ventas_Productos
        Case "Consulta de Materias Primas"
            Reporte_Materia_Prima
        Case "Consulta de Productos Terminados"
            Reporte_Producto_Terminado
    End Select
End Sub

Private Sub Reporte_Notas_Credito()
Dim Proveedor As String                                  'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset          'Consulta el proveedor al cual pertenece la nota de credito emitida
Dim Rs_Consulta_Adm_Notas_Credito_Proveedores As rdoResultset    'Consulta la nota de credito que fueron emitidas en las fechas dadas por el usuario

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Facturas_Proveedores.ListIndex > -1 Then
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
        Mi_SQL = "SELECT No_Nota_Credito,Proveedor_ID,Fecha,Importe,No_Factura FROM Adm_Notas_Credito_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & Format(Cmb_Facturas_Proveedores.ItemData(Cmb_Facturas_Proveedores.ListIndex), "00000") & "'"
        Mi_SQL = Mi_SQL & " AND Fecha >= " & Par_Fecha & Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "MM/dd/yyyy") & Par_Fecha
        Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "MM/dd/yyyy") & Par_Fecha
        Set Rs_Consulta_Adm_Notas_Credito_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Notas_Credito_Proveedores.EOF Then
            With Rs_Consulta_Adm_Notas_Credito_Proveedores
                Call Encabezado_Reportes("REPORTE NOTAS CRÉDITO POR PROVEEDOR", "", "", True)
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(1); "Nota Credito"; Spc(15 - Len("Nota Credito")); "Proveedor"; _
                          Spc(40 - Len("Proveedor")); "No. Factura"; Spc(15 - Len("No. Factura")); "Fecha"; Spc(15 - Len("Fecha")); "Importe"; Spc(6 - Len("Importe"))
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Nota Credito" & Chr(9) & "Proveedor" & Chr(9) & "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Importe"
                Grid_Reporte.AddItem ""
                While Not .EOF
                    If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Nombre FROM Cat_Proveedores"
                        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & .rdoColumns("Proveedor_ID") & "'"
                        Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Proveedores.EOF Then
                            Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                        End If
                        Rs_Consulta_Cat_Proveedores.Close
                    End If
                    Print #1, Spc(2); .rdoColumns("No_Nota_Credito"); Spc(2); Mid(Proveedor, 1, 40); _
                    Spc(45 - Len(Mid(Proveedor, 1, 40))); .rdoColumns("No_Factura"); Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                    Spc(15 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "##,##0.00"), 9); _
                    Spc(15 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "##,##0.00"), 9)))
                    Grid_Reporte.AddItem .rdoColumns("No_Nota_Credito") & Chr(9) & "Proveedor" & Chr(9) & _
                    Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                    .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Importe"), "##,##0.00") & Chr(13)
                    .MoveNext
                Wend
                Call Encabezado_Reportes("", "", "", False)
            End With
        Else
            MsgBox "No hay Notas de Credito en las fechas que proporcionó", vbInformation
        End If
        Rs_Consulta_Adm_Notas_Credito_Proveedores.Close
    Else
        MsgBox "Seleccione un Proveedor", vbInformation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Notas_Credito_Clientes
    'DESCRIPCIÓN: Consulta las notas de crédito de clientes
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 04-Julio-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Btn_Reporte_Notas_Credito_Clientes_Click()
Dim Cliente As String                                           'Almacena el nombre del cliente
Dim Empresa_NC As String                                        'Almacena el nombre de la Empresa
Dim Nota_Credito As String
Dim Factura_Origen As String                                    'Almacena el numero de factura
Dim Rs_Consulta_Cat_Clientes As rdoResultset                    'Consulta el proveedor al cual pertenece la nota de credito emitida
Dim Rs_Consulta_Clientes As rdoResultset                        'Manejo de registro
Dim Rs_Tipos_Notas_Credito As rdoResultset
Dim Rs_Consulta_Adm_Notas_Credito_Clientes As rdoResultset      'Consulta la nota de credito que fueron emitidas en las fechas dadas por el usuario
Dim Total_NC As Double
Dim Gran_Total As Double

    Grid_Reporte.Cols = 10
    Grid_Reporte.Rows = 0
    If Fra_Reporte_Notas_Credito.Caption = "Reporte de Notas de Crédito" Then
        'Conceptos de notas de crédito válidos
        Mi_SQL = "SELECT * FROM Cat_Tipos_Notas_Credito"
        'Filtra por el tipo de Nota de Crédito
        If Cmb_Tipo_Nota_Credito.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " WHERE Tipo_Nota_Credito_ID='" & Format(Cmb_Tipo_Nota_Credito.ItemData(Cmb_Tipo_Nota_Credito.ListIndex), "00000") & "'"
            'Valida que exista la ruta si no la crea
            If Len(Dir("C:\REP", vbDirectory)) = 0 Then
                MkDir "C:\REP"
            End If
            Open "C:\REP\Reportes.txt" For Output As #1
            Print #1, "************************************************************************************************************************"
            Print #1,
            If Cmb_Empresa_NC.ListIndex = 1 Then
                Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            Else
                Print #1, Conectar_Ayudante.Centra("PRODUCTOS NACIONALES DEL CENTRO S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "PRODUCTOS NACIONALES DEL CENTRO S.A. DE C.V."
            End If
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, Conectar_Ayudante.Centra("REPORTE DE NOTAS CREDITO", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "REPORTE DE NOTAS CREDITO"
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, Conectar_Ayudante.Centra("DEL " & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMM/yyyy")
            Print #1, "************************************************************************************************************************"
            Print #1,
        Else
            If MsgBox("¿Desea ver el reporte agrupado por tipos?", vbQuestion + vbYesNo) = vbYes Then
                Call Encabezado_Reportes("REPORTE DE NOTAS CREDITO", Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMMM/yyyy"), Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMMM/yyyy"), True)
            Else    'Reporte de notas de crédito sin agrupar
                Call Encabezado_Reportes("REPORTE DE NOTAS CREDITO", Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMMM/yyyy"), Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMMM/yyyy"), True)
                Mi_SQL = "SELECT * FROM Adm_Notas_Credito_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito_Clientes.Fecha BETWEEN '" & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Nota_Credito_Fecha_Final.Value, "MM/dd/yyyy") & "'"
                If Cmb_Nota_Credito_Cliente.ListIndex > -1 Then
                    Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Cliente_ID='" & Format(Cmb_Nota_Credito_Cliente.ItemData(Cmb_Nota_Credito_Cliente.ListIndex), "00000") & "'"
                End If
                If Cmb_Empresa_NC.ListIndex > 0 Then
                    Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Empresa_Destino=" & Cmb_Empresa_NC.ListIndex
                End If
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Notas_Credito_Clientes.Empresa,Adm_Notas_Credito_Clientes.Nota_Credito"
                Set Rs_Consulta_Adm_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1,
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, " Nota Credito      Fecha      Origen     Destino     Cliente                              Total  Apl  Consecutivo        "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Nota Credito" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Factura Origen" & Chr(9) & "Factura Destino" & Chr(9) & "Cliente" & Chr(9) & "Total" & Chr(9) & "Aplicada" & Chr(9) & "Consecutivo"
                Grid_Reporte.AddItem ""
                Total_NC = 0
                While Not Rs_Consulta_Adm_Notas_Credito_Clientes.EOF
                    With Rs_Consulta_Adm_Notas_Credito_Clientes
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Clientes.EOF Then
                            Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                        End If
                        Rs_Consulta_Cat_Clientes.Close
                        If Not IsNull(.rdoColumns("Nota_Credito")) Then
                            Nota_Credito = .rdoColumns("Nota_Credito")
                        Else
                            Nota_Credito = "SIN FOLIO "
                        End If
                        If Not IsNull(.rdoColumns("Empresa")) Then
                            If .rdoColumns("Empresa") = 1 Then
                                Empresa_NC = "NHE"
                            Else
                                Empresa_NC = "PNC"
                            End If
                        Else
                            Empresa_NC = "  "
                        End If
                        If Not IsNull(.rdoColumns("Factura_Origen")) Then
                            Factura_Origen = .rdoColumns("Factura_Origen")
                        Else
                            Factura_Origen = "          "
                        End If
                        If Trim(.rdoColumns("Comentarios")) <> "Cancelada" Then
                            Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                                      Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito")
                            Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito")
                            Total_NC = Total_NC + Val(.rdoColumns("Importe"))
                        Else
                            Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                                  Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito") & " Cancel"
                            Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito") & Chr(9) & "Cancelada"
                        End If
                        .MoveNext
                    End With
                Wend
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, "                                                                 T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
                Gran_Total = Gran_Total + Total_NC
                Rs_Consulta_Adm_Notas_Credito_Clientes.Close
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, "                                                      G R A  N   T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "GRAN TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
                Call Encabezado_Reportes("", "", "", False)
                Exit Sub
            End If
        End If
        Set Rs_Tipos_Notas_Credito = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Tipos_Notas_Credito.EOF Then
            While Not Rs_Tipos_Notas_Credito.EOF
                'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
                Mi_SQL = "SELECT Adm_Notas_Credito_Clientes.*,Adm_Detalles_Notas_Credito_Clientes.Producto_ID AS Producto_ID,Adm_Detalles_Notas_Credito_Clientes.Importe AS Importe_Concepto"
                Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito_Clientes,Adm_Detalles_Notas_Credito_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito_Clientes.No_Nota_Credito=Adm_Detalles_Notas_Credito_Clientes.No_Nota_Credito"
                Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Empresa=Adm_Detalles_Notas_Credito_Clientes.Empresa"
                Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Fecha>='" & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Fecha<='" & Format(Dtp_Nota_Credito_Fecha_Final.Value, "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Detalles_Notas_Credito_Clientes.Producto_ID='" & Rs_Tipos_Notas_Credito.rdoColumns("Tipo_Nota_Credito_ID") & "'"
                If Cmb_Nota_Credito_Cliente.ListIndex > -1 Then
                    Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Cliente_ID='" & Format(Cmb_Nota_Credito_Cliente.ItemData(Cmb_Nota_Credito_Cliente.ListIndex), "00000") & "'"
                End If
                If Cmb_Empresa_NC.ListIndex > 0 Then
                    Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Empresa_Destino=" & Cmb_Empresa_NC.ListIndex
                End If
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Notas_Credito_Clientes.Empresa,Adm_Notas_Credito_Clientes.Nota_Credito"
                Set Rs_Consulta_Adm_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1,
                Print #1, Spc(30); "TIPO DE NOTA DE CREDITO:     " & Rs_Tipos_Notas_Credito.rdoColumns("Descripcion")
                Print #1,
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, " Nota Credito      Fecha      Origen     Destino     Cliente                              Total  Apl  Consecutivo        "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "TIPO DE NOTA DE CREDITO" & Chr(9) & Rs_Tipos_Notas_Credito.rdoColumns("Descripcion")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Nota Credito" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Factura Origen" & Chr(9) & "Factura Destino" & Chr(9) & "Cliente" & Chr(9) & "Total" & Chr(9) & "Aplicada" & Chr(9) & "Consecutivo"
                Grid_Reporte.AddItem ""
                Total_NC = 0
                While Not Rs_Consulta_Adm_Notas_Credito_Clientes.EOF
                    With Rs_Consulta_Adm_Notas_Credito_Clientes
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                        Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Clientes.EOF Then
                            Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                        End If
                        Rs_Consulta_Cat_Clientes.Close
                        If Not IsNull(.rdoColumns("Nota_Credito")) Then
                            Nota_Credito = .rdoColumns("Nota_Credito")
                        Else
                            Nota_Credito = "SIN FOLIO "
                        End If
                        If Not IsNull(.rdoColumns("Empresa")) Then
                            If .rdoColumns("Empresa") = 1 Then
                                Empresa_NC = "NHE"
                            Else
                                Empresa_NC = "PNC"
                            End If
                        Else
                            Empresa_NC = "  "
                        End If
                        If Not IsNull(.rdoColumns("Factura_Origen")) Then
                            Factura_Origen = .rdoColumns("Factura_Origen")
                        Else
                            Factura_Origen = "          "
                        End If
                        If Trim(.rdoColumns("Comentarios")) <> "Cancelada" Then
                            Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                                      Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe_Concepto"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito")
                            Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe_Concepto"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito")
                            Total_NC = Total_NC + Val(.rdoColumns("Importe_Concepto"))
                        Else
                            Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                                  Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe_Concepto"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito") & " Cancel"
                            Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe_Concepto"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito") & Chr(9) & "Cancelada"
                        End If
                        .MoveNext
                    End With
                Wend
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, "                                                                 T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
                Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
                Gran_Total = Gran_Total + Total_NC
                Rs_Consulta_Adm_Notas_Credito_Clientes.Close
                Rs_Tipos_Notas_Credito.MoveNext
            Wend
        End If
        Rs_Tipos_Notas_Credito.Close
        If Cmb_Tipo_Nota_Credito.ListIndex > -1 Then
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "                                                      G R A  N   T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "GRAN TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13305
            Me.Height = 8340
            Me.Left = 0
            Me.Top = 0
        Else
            'Valida si hay notas de crédito de otros tipos para ser mostrados en pantalla
            Mi_SQL = "SELECT Adm_Notas_Credito_Clientes.*,Adm_Detalles_Notas_Credito_Clientes.Producto_ID AS Producto_ID,Adm_Detalles_Notas_Credito_Clientes.Importe AS Importe_Concepto"
            Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito_Clientes,Adm_Detalles_Notas_Credito_Clientes"
            Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito_Clientes.No_Nota_Credito=Adm_Detalles_Notas_Credito_Clientes.No_Nota_Credito"
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Empresa=Adm_Detalles_Notas_Credito_Clientes.Empresa"
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Tipo<>'CONCEPTO'"
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Fecha BETWEEN '" & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Nota_Credito_Fecha_Final.Value, "MM/dd/yyyy") & "'"
            If Cmb_Nota_Credito_Cliente.ListIndex > -1 Then
                Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Cliente_ID='" & Format(Cmb_Nota_Credito_Cliente.ItemData(Cmb_Nota_Credito_Cliente.ListIndex), "00000") & "'"
            End If
            If Cmb_Empresa_NC.ListIndex > 0 Then
                Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Clientes.Empresa_Destino=" & Cmb_Empresa_NC.ListIndex
            End If
            Mi_SQL = Mi_SQL & " ORDER BY Adm_Notas_Credito_Clientes.Empresa,Adm_Notas_Credito_Clientes.Nota_Credito"
            Set Rs_Consulta_Adm_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Print #1,
            Print #1, Spc(30); "TIPO DE NOTA DE CREDITO:     " & "PRODUCTOS"
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " Nota Credito      Fecha      Origen     Destino     Cliente                              Total  Apl  Consecutivo        "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TIPO DE NOTA DE CREDITO" & Chr(9) & "PRODUCTOS"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "Nota Credito" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Factura Origen" & Chr(9) & "Factura Destino" & Chr(9) & "Cliente" & Chr(9) & "Total" & Chr(9) & "Aplicada" & Chr(9) & "Consecutivo"
            Grid_Reporte.AddItem ""
            Total_NC = 0
            While Not Rs_Consulta_Adm_Notas_Credito_Clientes.EOF
                With Rs_Consulta_Adm_Notas_Credito_Clientes
                    'Consulta el nombre del cliente
                    Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                    If Not IsNull(.rdoColumns("Nota_Credito")) Then
                        Nota_Credito = .rdoColumns("Nota_Credito")
                    Else
                        Nota_Credito = "SIN FOLIO "
                    End If
                    If Not IsNull(.rdoColumns("Empresa")) Then
                        If .rdoColumns("Empresa") = 1 Then
                            Empresa_NC = "NHE"
                        Else
                            Empresa_NC = "PNC"
                        End If
                    Else
                        Empresa_NC = "  "
                    End If
                    If Not IsNull(.rdoColumns("Factura_Origen")) Then
                        Factura_Origen = .rdoColumns("Factura_Origen")
                    Else
                        Factura_Origen = "          "
                    End If
                    If Trim(.rdoColumns("Comentarios")) <> "Cancelada" Then
                        Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                                  Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe_Concepto"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito")
                        Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe_Concepto"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito")
                        Total_NC = Total_NC + Val(.rdoColumns("Importe_Concepto"))
                    Else
                        Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); Factura_Origen; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                              Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe_Concepto"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito") & " Cancel"
                        Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe_Concepto"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito") & Chr(9) & "Cancelada"
                    End If
                    .MoveNext
                End With
            Wend
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "                                                                 T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
            Gran_Total = Gran_Total + Total_NC
            Rs_Consulta_Adm_Notas_Credito_Clientes.Close
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "                                                      G R A  N   T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "GRAN TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Gran_Total, "#,###,##0.00"), 14)
            Call Encabezado_Reportes("", "", "", False)
        End If
    Else
        Mi_SQL = "SELECT Adm_Notas_Credito_Remisiones.*"
        Mi_SQL = Mi_SQL & " FROM Adm_Notas_Credito_Remisiones,Adm_Notas_Credito_Remisiones_Detalles"
        Mi_SQL = Mi_SQL & " WHERE Adm_Notas_Credito_Remisiones.No_Nota_Credito=Adm_Notas_Credito_Remisiones_Detalles.No_Nota_Credito"
        Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones.Empresa=Adm_Notas_Credito_Remisiones_Detalles.Empresa"
        Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones.Fecha>='" & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
        Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones.Fecha<='" & Format(Dtp_Nota_Credito_Fecha_Final.Value, "MM/dd/yyyy") & "'"
        If Cmb_Nota_Credito_Cliente.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones.Cliente_ID='" & Format(Cmb_Nota_Credito_Cliente.ItemData(Cmb_Nota_Credito_Cliente.ListIndex), "00000") & "'"
        End If
        If Cmb_Empresa_NC.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones.Empresa_Remision=" & Cmb_Empresa_NC.ListIndex
        End If
        'Filtra por el tipo de Nota de Crédito
        If Cmb_Tipo_Nota_Credito.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND Adm_Notas_Credito_Remisiones_Detalles.Producto_ID='" & Format(Cmb_Tipo_Nota_Credito.ItemData(Cmb_Tipo_Nota_Credito.ListIndex), "00000") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Adm_Notas_Credito_Remisiones.Empresa,Adm_Notas_Credito_Remisiones.No_Nota_Credito"
        Set Rs_Consulta_Adm_Notas_Credito_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Cmb_Empresa_NC.ListIndex > 0 Then
            'Valida que exista la ruta si no la crea
            If Len(Dir("C:\REP", vbDirectory)) = 0 Then
                MkDir "C:\REP"
            End If
            Open "C:\REP\Reportes.txt" For Output As #1
            Print #1, "************************************************************************************************************************"
            Print #1,
            If Cmb_Empresa_NC.ListIndex = 1 Then
                Print #1, Conectar_Ayudante.Centra("NATURAL HEALTH S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "NATURAL HEALTH S.A. DE C.V."
            Else
                Print #1, Conectar_Ayudante.Centra("PRODUCTOS NACIONALES DEL CENTRO S.A. DE C.V.", Len("************************************************************************************************************************"))
                Grid_Reporte.AddItem "PRODUCTOS NACIONALES DEL CENTRO S.A. DE C.V."
            End If
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, Conectar_Ayudante.Centra("REPORTE DE NOTAS CREDITO DE REMISION", Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "REPORTE DE NOTAS CREDITO DE REMISION"
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, Conectar_Ayudante.Centra("DEL " & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMM/yyyy"), Len("************************************************************************************************************************"))
            Grid_Reporte.AddItem "DEL " & Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMM/yyyy") & " AL " & Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMM/yyyy")
            Print #1, "************************************************************************************************************************"
            Print #1,
        Else
            Call Encabezado_Reportes("REPORTE DE NOTAS CREDITO DE REMISION", Format(Dtp_Nota_Credito_Fecha_Inicio.Value, "dd/MMMM/yyyy"), Format(Dtp_Nota_Credito_Fecha_Final.Value, "dd/MMMM/yyyy"), True)
        End If
        If Cmb_Tipo_Nota_Credito.ListIndex > -1 Then
            Print #1,
            Print #1, Spc(30); "TIPO DE NOTA DE CREDITO:     " & Cmb_Tipo_Nota_Credito.Text
            Print #1,
        End If
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, " Nota Credito      Fecha      Origen     Destino     Cliente                              Total  Apl  Consecutivo        "
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Grid_Reporte.AddItem ""
        Grid_Reporte.AddItem "Nota Credito" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Factura Origen" & Chr(9) & "Factura Destino" & Chr(9) & "Cliente" & Chr(9) & "Total" & Chr(9) & "Aplicada" & Chr(9) & "Consecutivo"
        Grid_Reporte.AddItem ""
        While Not Rs_Consulta_Adm_Notas_Credito_Clientes.EOF
            With Rs_Consulta_Adm_Notas_Credito_Clientes
                'Consulta el nombre del cliente
                Mi_SQL = "SELECT Nombre FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID = '" & .rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Clientes.EOF Then
                    Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Clientes.Close
                Nota_Credito = .rdoColumns("No_Nota_Credito")
                If .rdoColumns("Empresa") = 1 Then
                    Empresa_NC = "NHE"
                Else
                    Empresa_NC = "PNC"
                End If
                If Trim(.rdoColumns("Comentarios")) <> "Cancelada" Then
                    Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); "          "; Spc(2); .rdoColumns("No_Remision"); Spc(2); _
                              Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito")
                    Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("No_Remision") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito")
                Else
                    Print #1, Spc(1); Nota_Credito; Spc(1); Empresa_NC; Spc(2); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); Spc(2); "          "; Spc(2); .rdoColumns("Factura_Destino"); Spc(2); _
                          Mid(Cliente, 1, 30); Spc(32 - Len(Mid(Cliente, 1, 30))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,###,##0.00"), 12); Spc(2); .rdoColumns("Aplicada"); Spc(2); .rdoColumns("No_Nota_Credito") & " Cancel"
                    Grid_Reporte.AddItem Nota_Credito & Chr(9) & Empresa_NC & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yy") & Chr(9) & Factura_Origen & Chr(9) & .rdoColumns("Factura_Destino") & Chr(9) & Cliente & Chr(9) & Format(.rdoColumns("Importe"), "##,##0.00") & Chr(9) & .rdoColumns("Aplicada") & Chr(9) & .rdoColumns("No_Nota_Credito") & Chr(9) & "Cancelada"
                End If
                Total_NC = Total_NC + Val(.rdoColumns("Importe"))
                .MoveNext
            End With
        Wend
        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
        Print #1, "                                                                 T O T A L        " & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
        Grid_Reporte.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_NC, "#,###,##0.00"), 14)
        If Cmb_Empresa_NC.ListIndex > 0 Then
            Print #1,
            Print #1, "------------------------------------------------------------------------------------------------------------------------"
            Close #1
            Me.Edm_Reporte.Font.Size = 8
            Me.Edm_Reporte.Font = "Courier New"
            Me.Edm_Reporte.FileName = "C:\REP\Reportes.txt"
            Me.Edm_Reporte.Visible = True
            Ancho_Frame = Me.Width
            Largo_Frame = Me.Height
            Me.Width = 13305
            Me.Height = 8340
            Me.Left = 0
            Me.Top = 0
        Else
            Call Encabezado_Reportes("", "", "", False)
        End If
        Rs_Consulta_Adm_Notas_Credito_Clientes.Close
    End If
End Sub

Private Sub Btn_Reporte_Pedido_Factura_Click()
    If Chk_Pedido_Facturado.Value = 1 Or Chk_Pedido_Sin_Factura.Value = 1 Then
        Reporte_Pedidos_Facturas
    Else
        MsgBox "Debe seleccionar al menos una opción", vbExclamation
        Chk_Pedido_Facturado.SetFocus
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Reporte_Pedidos_Cancelados
    'DESCRIPCIÓN: Muestra el reporte de los pedidos que han sido cancelados
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 08-Mayo-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
    'PARÁMETROS:
'*******************************************************************************
Private Sub Reporte_Pedidos_Cancelados()
Dim Rs_Consulta_Pedidos As rdoResultset         'Manejo de registro
Dim Rs_Consulta_Clientes As rdoResultset        'Manejo de registro
Dim Rs_Consulta_Ciudades As rdoResultset        'Manejo de registro
Dim No_Pedido As String                         'Almacena el numero de pedido
Dim Ciudad As String                            'Almacena el nombre de la ciudad
Dim Total As Double                             'Almacena el total del pedido
Dim Total_Producto As Double                    'Almacena el total de productos
Dim Productos_Pedido As Double                  'Almacena el total del producto por pedido
Dim Total_Pedido As Double
Dim Usuario_Cancelo As String
Dim Fecha_Cancelo As String
Dim Motivo_Cancelo As String

    'Coloca el encabezado de los reportes
    Call Encabezado_Reportes("REPORTE DE PEDIDOS CANCELADOS", Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "dd/MMM/yyyy"), True)
    Grid_Reporte.AddItem ""
    'Consulta los pedidos que han sido cancelados
    Mi_SQL = "SELECT Ope_Pedidos.No_Pedido,Fecha_Pedido,Ope_Pedidos.Cliente_ID,Detalles_Pedidos.Producto_ID,Detalles_Pedidos.Descripcion,Detalles_Pedidos.Cantidad,Detalles_Pedidos.Importe,Cat_Productos.Codigo,Detalles_Pedidos.Estatus,Cancelo,Fecha_Cancelo,Comentarios_Cancelo"
    Mi_SQL = Mi_SQL & " FROM Ope_Pedidos,Detalles_Pedidos,Cat_Productos"
    Mi_SQL = Mi_SQL & " WHERE Ope_Pedidos.No_Pedido=Detalles_Pedidos.No_Pedido"
    Mi_SQL = Mi_SQL & " AND Cat_Productos.Producto_ID=Detalles_Pedidos.Producto_ID"
    Mi_SQL = Mi_SQL & " AND Detalles_Pedidos.Estatus='Cancelado'"
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido>=" & Par_Fecha & Format(Dtp_Fecha_Inicio_Pedido_Producto.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " AND Fecha_Pedido<=" & Par_Fecha & Format(Dtp_Fecha_Fin_Pedidos_Productos.Value, "MM/dd/yyyy") & Par_Fecha
    Mi_SQL = Mi_SQL & " ORDER BY Ope_Pedidos.Fecha_Pedido,Ope_Pedidos.No_Pedido"
    Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Pedidos
        While Not .EOF
            If No_Pedido <> .rdoColumns("No_Pedido") Then
                'Consulta los datos del cliente
                Mi_SQL = "SELECT Cliente_ID,Nombre,Direccion,Colonia,CP,Ciudad_ID"
                Mi_SQL = Mi_SQL & " FROM Cat_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Pedidos.rdoColumns("Cliente_ID") & "'"
                Set Rs_Consulta_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Clientes.EOF Then
                    If Not IsNull(Rs_Consulta_Clientes.rdoColumns("Ciudad_ID")) Then
                        'Consulta la ciudad del cliente
                        Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                        Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                        Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                        Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Clientes.rdoColumns("Ciudad_ID") & "'"
                        Set Rs_Consulta_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Ciudades.EOF Then
                            Ciudad = Rs_Consulta_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Ciudades.rdoColumns("Nombre_Corto")
                        Else
                            Ciudad = ""
                        End If
                        Rs_Consulta_Ciudades.Close
                    Else
                        Ciudad = ""
                    End If
                    Print #1,
                    If Productos_Pedido > 0 Then
                        Print #1, Spc(70); "TOTAL PEDIDO :    "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Productos_Pedido, "###,###"), 7); Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pedido, "#,###,##0.00"), 12)
                        Print #1, Spc(2); "CANCELO       : "; .rdoColumns("Cancelo")
                        Print #1, Spc(2); "FECHA CANCELO : "; Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy")
                        Print #1, Spc(2); "MOTIVO CANCELO: "; .rdoColumns("Comentarios_Cancelo")
                        Print #1,
                        Grid_Reporte.AddItem "" & Chr(9) & "TOTAL PEDIDO:" & Chr(9) & Format(Productos_Pedido, "###,###") & Chr(9) & Format(Total_Pedido, "#,###,##0.00")
                        Grid_Reporte.AddItem "CANCELO: " & Chr(9) & .rdoColumns("Cancelo")
                        Grid_Reporte.AddItem "FECHA CANCELO: " & Chr(9) & Format(.rdoColumns("Fecha_Cancelo"), "dd/MMM/yyyy")
                        Grid_Reporte.AddItem "MOTIVO CANCELO: " & Chr(9) & .rdoColumns("Comentarios_Cancelo")
                        Grid_Reporte.AddItem ""
                        Productos_Pedido = 0
                        Total_Pedido = 0
                    End If
                    Print #1, "________________________________________________________________________________________________________________________"
                    Print #1,
                    Print #1, Spc(2); "Cliente:   "; Rs_Consulta_Clientes.rdoColumns("Nombre")
                    Print #1, Spc(2); "Direccion: "; Rs_Consulta_Clientes.rdoColumns("Direccion")
                    Print #1, Spc(2); "Colonia:   "; Rs_Consulta_Clientes.rdoColumns("Colonia"); Spc(2); "CP: "; Rs_Consulta_Clientes.rdoColumns("CP")
                    Print #1, Spc(2); "Ciudad:    "; Ciudad
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "Cliente:" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Nombre")
                    Grid_Reporte.AddItem "Direccion:" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Direccion")
                    Grid_Reporte.AddItem "Colonia: " & Chr(9) & Rs_Consulta_Clientes.rdoColumns("Colonia") & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Clientes.rdoColumns("CP")
                    Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad
                End If
                Rs_Consulta_Clientes.Close
                'Imprime los pedidos
                Print #1,
                Grid_Reporte.AddItem ""
                Print #1, Spc(2); "No. Pedido : "; .rdoColumns("No_Pedido"); Spc(5); "Fecha : "; Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                Grid_Reporte.AddItem "Pedido :" & Chr(9) & .rdoColumns("No_Pedido") & Chr(9) & "Fecha :" & Chr(9) & Format(.rdoColumns("Fecha_Pedido"), "dd/MMM/yyyy")
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Print #1, "   Codigo               Descripcion                                                      Cantidad        Importe        "
                Print #1, "------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Importe"
                No_Pedido = .rdoColumns("No_Pedido")
                Usuario_Cancelo = .rdoColumns("Cancelo")
                Fecha_Cancelo = .rdoColumns("Fecha_Cancelo")
                Motivo_Cancelo = .rdoColumns("Comentarios_Cancelo")
            End If
            Print #1, Spc(2); .rdoColumns("Codigo"); Spc(21 - Len(.rdoColumns("Codigo"))); Mid(.rdoColumns("Descripcion"), 1, 59); Spc(62 - Len(Mid(.rdoColumns("Descripcion"), 1, 59))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Cantidad"), "###,###"), 10); Spc(8); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Val(.rdoColumns("Importe")), "###,##0.00"), 10)
            Grid_Reporte.AddItem .rdoColumns("Codigo") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Cantidad") & Chr(9) & .rdoColumns("Importe")
            Productos_Pedido = Productos_Pedido + .rdoColumns("Cantidad")
            Total_Producto = Total_Producto + .rdoColumns("Cantidad")
            Total_Pedido = Total_Pedido + .rdoColumns("Importe")
            Total = Total + .rdoColumns("Importe")
            .MoveNext
        Wend
    End With
    Rs_Consulta_Pedidos.Close
    Print #1,
    Print #1, Spc(70); "TOTAL PEDIDO :    "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Productos_Pedido, "###,###"), 7); Spc(6); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Pedido, "#,###,##0.00"), 12)
    Print #1, Spc(2); "CANCELO       : "; Usuario_Cancelo
    Print #1, Spc(2); "FECHA CANCELO : "; Format(Fecha_Cancelo, "dd/MMM/yyyy")
    Print #1, Spc(2); "MOTIVO CANCELO: "; Motivo_Cancelo
    Print #1,
    Grid_Reporte.AddItem "" & Chr(9) & "TOTAL PEDIDO:" & Chr(9) & Format(Productos_Pedido, "###,###") & Chr(9) & Format(Total_Pedido, "#,###,##0.00")
    Grid_Reporte.AddItem "CANCELO: " & Chr(9) & Usuario_Cancelo
    Grid_Reporte.AddItem "FECHA CANCELO: " & Chr(9) & Format(Fecha_Cancelo, "dd/MMM/yyyy")
    Grid_Reporte.AddItem "MOTIVO CANCELO: " & Chr(9) & Motivo_Cancelo
    Grid_Reporte.AddItem ""
    Productos_Pedido = 0
    Total_Pedido = 0
    Print #1, "________________________________________________________________________________________________________________________"
    Print #1,
    Print #1, Spc(70); "GRAN TOTAL:     "; Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Producto, "#,###,###"), 9); Spc(4); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 14)
    Grid_Reporte.AddItem "" & Chr(9) & "GRAN TOTAL:" & Chr(9) & Format(Total_Producto, "###,###") & Chr(9) & Format(Total, "#,###,##0.00")
    Call Encabezado_Reportes("", "", "", False)
End Sub

Private Sub Btn_Reporte_Pedidos_Cliente_Click()
    Select Case Fra_Reporte_Pedidos_Clientes.Caption
        Case "Pedidos por Producto"
            Reporte_Pedidos_Productos
        Case "Facturas por Producto"
            Reporte_Facturas_Productos
        Case "Remisiones por Producto"
            Reporte_Remisiones_Productos
        Case "Concentrado de Pedidos"
            Reporte_Concentrado_Pedidos
        Case "Reporte de Pedidos"
            Reporte_Pedidos
        Case "Reporte de Pedidos Cancelados"
            Reporte_Pedidos_Cancelados
    End Select
End Sub

Private Sub Reporte_Facturas_Por_Proveedor()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset        'Manejo de registro para las facturas
Dim Rs_Consulta_Proveedores As rdoResultset                 'Manejo de registro de la tabla Clientes
Dim Total_Facturas As Double
Dim Estatus As String

    If Cmb_Facturas_Proveedores.ListIndex > -1 Then
        Call Encabezado_Reportes("REPORTE DE FACTURAS POR PROVEEDOR", Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "dd/MMM/yyyy"), True)
        'Consulta el cliente seleccionado
        Mi_SQL = "SELECT Proveedor_ID,Nombre,Direccion,Colonia,CP,Ciudad_Estado,RFC,Telefono1"
        Mi_SQL = Mi_SQL & " FROM Cat_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Facturas_Proveedores.ItemData(Cmb_Facturas_Proveedores.ListIndex), "00000") & "'"
        Set Rs_Consulta_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Proveedores.EOF Then
            Print #1, Spc(30); "Proveedor : "; Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Grid_Reporte.AddItem "Proveedor" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Print #1, Spc(30); "R.F.C.    : "; Rs_Consulta_Proveedores.rdoColumns("RFC")
            Grid_Reporte.AddItem "R.F.C." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("RFC")
            Print #1, Spc(30); "Direccion : "; Mid(Rs_Consulta_Proveedores.rdoColumns("Direccion"), 1, 70)
            Grid_Reporte.AddItem "Direccion" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Direccion")
            Print #1, Spc(30); "Colonia   : "; Rs_Consulta_Proveedores.rdoColumns("Colonia")
            Grid_Reporte.AddItem "Colonia" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Colonia") & Chr(9) & "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("CP")
            Print #1, Spc(30); "Ciudad    : "; Rs_Consulta_Proveedores.rdoColumns("Ciudad_Estado")
            Grid_Reporte.AddItem "Ciudad" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Ciudad_Estado")
            Print #1, Spc(30); "C.P.      : "; Rs_Consulta_Proveedores.rdoColumns("CP")
            Grid_Reporte.AddItem "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("CP")
            Print #1, Spc(30); "Telefono  : "; Rs_Consulta_Proveedores.rdoColumns("Telefono1")
            Grid_Reporte.AddItem "Telefono" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Telefono1")
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Factura              Fecha         Tipo                     Total      Estatus                                     "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Tipo" & Chr(9) & "Total" & Chr(9) & "Estatus"
            'Consulta las facturas registradas del proveedor
            Mi_SQL = "SELECT * FROM Adm_Facturas_Proveedores"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & "'"
            Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Facturas_Proveedores.EOF
                With Rs_Consulta_Facturas_Proveedores
                    If .rdoColumns("Cancelada") = "S" Then
                        Estatus = "Cancelada"
                    Else
                        If .rdoColumns("Pagada") = "S" Then
                            Estatus = "Pagada"
                        Else
                            Estatus = "Pendiente"
                        End If
                    End If
                    If .rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                End With
                Rs_Consulta_Facturas_Proveedores.MoveNext
            Wend
            Rs_Consulta_Facturas_Proveedores.Close
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(30); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End If
    Else
        MsgBox "Seleccione un proveedor", vbInformation
    End If
End Sub

Private Sub Reporte_Actas_Liberacion()
Dim Rs_Consulta_Facturas_Proveedores As rdoResultset        'Manejo de registro para las facturas
Dim Rs_Consulta_Proveedores As rdoResultset                 'Manejo de registro de la tabla Clientes
Dim Rs_Consulta_Compras As rdoResultset
Dim Rs_Consulta_Pedidos As rdoResultset
Dim Total_Facturas As Double
Dim Estatus As String
Dim Cantidad_Pedida As Double

    If Cmb_Facturas_Proveedores.ListIndex > -1 Then
        Call Encabezado_Reportes("REPORTE DE ACTAS DE LIBERACION POR PROVEEDOR", Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "dd/MMM/yyyy"), True)
        'Consulta el cliente seleccionado
        Mi_SQL = "SELECT Proveedor_ID,Nombre,Direccion,Colonia,CP,Ciudad_Estado,RFC,Telefono1"
        Mi_SQL = Mi_SQL & " FROM Cat_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Facturas_Proveedores.ItemData(Cmb_Facturas_Proveedores.ListIndex), "00000") & "'"
        Set Rs_Consulta_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Proveedores.EOF Then
            Print #1, Spc(30); "Proveedor : "; Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Grid_Reporte.AddItem "Proveedor" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Nombre")
            Print #1, Spc(30); "R.F.C.    : "; Rs_Consulta_Proveedores.rdoColumns("RFC")
            Grid_Reporte.AddItem "R.F.C." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("RFC")
            Print #1, Spc(30); "Direccion : "; Mid(Rs_Consulta_Proveedores.rdoColumns("Direccion"), 1, 70)
            Grid_Reporte.AddItem "Direccion" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Direccion")
            Print #1, Spc(30); "Colonia   : "; Rs_Consulta_Proveedores.rdoColumns("Colonia")
            Grid_Reporte.AddItem "Colonia" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Colonia") & Chr(9) & "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("CP")
            Print #1, Spc(30); "Ciudad    : "; Rs_Consulta_Proveedores.rdoColumns("Ciudad_Estado")
            Grid_Reporte.AddItem "Ciudad" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Ciudad_Estado")
            Print #1, Spc(30); "C.P.      : "; Rs_Consulta_Proveedores.rdoColumns("CP")
            Grid_Reporte.AddItem "C.P." & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("CP")
            Print #1, Spc(30); "Telefono  : "; Rs_Consulta_Proveedores.rdoColumns("Telefono1")
            Grid_Reporte.AddItem "Telefono" & Chr(9) & Rs_Consulta_Proveedores.rdoColumns("Telefono1")
            Print #1,
            Grid_Reporte.AddItem ""
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, "  No. Factura              Fecha         Tipo                     Total      Estatus                                     "
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Fecha" & Chr(9) & "Tipo" & Chr(9) & "Total" & Chr(9) & "Estatus"
            'Consulta las facturas registradas del proveedor
            Mi_SQL = "SELECT * FROM Adm_Facturas_Proveedores"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Rs_Consulta_Proveedores.rdoColumns("Proveedor_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "MM/dd/yyyy") & "' AND '" & Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "MM/dd/yyyy") & "'"
            If Trim(Txt_No_Factura.Text) <> "" Then
                Mi_SQL = Mi_SQL & " AND No_Factura LIKE '%" & Trim(Txt_No_Factura.Text) & "%'"
            End If
            Mi_SQL = Mi_SQL & " ORDER BY Fecha"
            Set Rs_Consulta_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Consulta_Facturas_Proveedores.EOF
                With Rs_Consulta_Facturas_Proveedores
                    'Estatus
                    If .rdoColumns("Cancelada") = "S" Then
                        Estatus = "Cancelada"
                    Else
                        If .rdoColumns("Pagada") = "S" Then
                            Estatus = "Pagada"
                        Else
                            Estatus = "Pendiente"
                        End If
                    End If
                    If .rdoColumns("Cancelada") = "N" Then
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00") & Chr(9) & Estatus
                        Total_Facturas = Total_Facturas + .rdoColumns("Total")
                    Else
                        Print #1, Spc(1); Trim(.rdoColumns("No_Factura")); _
                            Spc(25 - Len(Trim(.rdoColumns("No_Factura")))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(5); UCase(.rdoColumns("Tipo")); _
                            Spc(20 - Len(.rdoColumns("Tipo"))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 14); _
                            Spc(3); Estatus
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & .rdoColumns("Tipo") & Chr(9) & "0.00" & Chr(9) & Estatus
                    End If
                    'Consulta la factura de compras para mostrar en pantalla la información
                    Mi_SQL = "SELECT Tmp_Facturas_Proveedores.No_Control,Tmp_Facturas_Proveedores.No_Orden_Compra,Alm_Entradas_Materia_Prima.No_Pedido_Proveedor,Alm_Entradas_Materia_Prima_Detalles.No_Acta,Alm_Entradas_Materia_Prima_Detalles.Estatus,Alm_Entradas_Materia_Prima_Detalles.Cantidad,Alm_Entradas_Materia_Prima_Detalles.Usuario_Libero,Alm_Entradas_Materia_Prima_Detalles.Fecha_Libero,Alm_Entradas_Materia_Prima_Detalles.Comentarios_Libero,Alm_Entradas_Materia_Prima_Detalles.Materia_Prima_ID,Alm_Entradas_Materia_Prima_Detalles.Numero_Lote"
                    Mi_SQL = Mi_SQL & " FROM Tmp_Facturas_Proveedores,Alm_Entradas_Materia_Prima,Alm_Entradas_Materia_Prima_Detalles"
                    Mi_SQL = Mi_SQL & " WHERE Tmp_Facturas_Proveedores.No_Control=Alm_Entradas_Materia_Prima.No_Control"
                    Mi_SQL = Mi_SQL & " AND Alm_Entradas_Materia_Prima.No_Entrada=Alm_Entradas_Materia_Prima_Detalles.No_Entrada"
                    Mi_SQL = Mi_SQL & " AND Tmp_Facturas_Proveedores.Proveedor_ID='" & Format(Cmb_Facturas_Proveedores.ItemData(Cmb_Facturas_Proveedores.ListIndex), "00000") & "'"
                    Mi_SQL = Mi_SQL & " AND Tmp_Facturas_Proveedores.No_Factura_Proveedor='" & Trim(.rdoColumns("No_Factura")) & "'"
                    'Mi_SQL = Mi_SQL & " AND Tmp_Facturas_Proveedores.No_Factura='" & Trim(.rdoColumns("No_Factura")) & "'"
                    Set Rs_Consulta_Compras = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Compras.EOF Then
                        Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                        Print #1, "     No. Acta   F. Liberó    Usuario Libero              Cantidad   No. Lote      Estatus   Comentarios           "
                        While Not Rs_Consulta_Compras.EOF
                            'Consulta la cantidad pedida
                            Mi_SQL = "SELECT Cantidad FROM Alm_Pedidos_Proveedor_Detalles"
                            Mi_SQL = Mi_SQL & " WHERE No_Pedido_Proveedor='" & Rs_Consulta_Compras.rdoColumns("No_Pedido_Proveedor") & "'"
                            Mi_SQL = Mi_SQL & " AND Materia_Prima_ID='" & Rs_Consulta_Compras.rdoColumns("Materia_Prima_ID") & "'"
                            Set Rs_Consulta_Pedidos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Pedidos.EOF Then
                                Cantidad_Pedida = Rs_Consulta_Pedidos.rdoColumns("Cantidad")
                            Else
                                Cantidad_Pedida = 0
                            End If
                            Rs_Consulta_Pedidos.Close
                            Print #1, Spc(6); Rs_Consulta_Compras.rdoColumns("No_Acta"); _
                                Spc(10 - Len(Rs_Consulta_Compras.rdoColumns("No_Acta"))); Format(Rs_Consulta_Compras.rdoColumns("Fecha_Libero"), "dd/MMM/yy"); _
                                Spc(2); Mid(Rs_Consulta_Compras.rdoColumns("Usuario_Libero"), 1, 20); _
                                Spc(22 - Len(Mid(Rs_Consulta_Compras.rdoColumns("Usuario_Libero"), 1, 20))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Compras.rdoColumns("Cantidad"), "#,##0.00"), 14); _
                                Spc(4); Mid(Rs_Consulta_Compras.rdoColumns("Numero_Lote"), 1, 13); _
                                Spc(15 - Len(Mid(Rs_Consulta_Compras.rdoColumns("Numero_Lote"), 1, 13))); Rs_Consulta_Compras.rdoColumns("Estatus"); _
                                Spc(10 - Len(Rs_Consulta_Compras.rdoColumns("Estatus"))); Mid(Rs_Consulta_Compras.rdoColumns("Comentarios_Libero"), 1, 30)
                            Grid_Reporte.AddItem "ACTA " & Rs_Consulta_Compras.rdoColumns("No_Acta") _
                                & Chr(9) & Format(Rs_Consulta_Compras.rdoColumns("Fecha_Libero"), "dd/MMM/yyyy") _
                                & Chr(9) & Rs_Consulta_Compras.rdoColumns("Usuario_Libero") _
                                & Chr(9) & Rs_Consulta_Compras.rdoColumns("Cantidad") _
                                & Chr(9) & Rs_Consulta_Compras.rdoColumns("Numero_Lote") _
                                & Chr(9) & Rs_Consulta_Compras.rdoColumns("Estatus") _
                                & Chr(9) & Rs_Consulta_Compras.rdoColumns("Comentarios_Libero")
                            Grid_Reporte.FixedRows = 1
                            Rs_Consulta_Compras.MoveNext
                        Wend
                        Print #1, "_________________________________________________________________________________________________________________________"
                    End If
                    Rs_Consulta_Compras.Close
                End With
                Rs_Consulta_Facturas_Proveedores.MoveNext
            Wend
            Rs_Consulta_Facturas_Proveedores.Close
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, Spc(15); "T O T A L E S "; Spc(30); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Facturas, "#,###,##0.00"), 14)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "" & Chr(9) & "T O T A L E S" & Chr(9) & "" & Chr(9) & Format(Total_Facturas, "#,###,##0.00")
            Call Encabezado_Reportes("", "", "", False)
        End If
    Else
        MsgBox "Seleccione un proveedor", vbInformation
    End If
End Sub

Private Sub Reporte_Cambio_Ubicaciones()
Dim Mi_SQL As String
Dim Rs_Consulta_Ubicaciones As rdoResultset
Dim Rs_Consulta_Productos As rdoResultset

    Call Encabezado_Reportes("REPORTE DE CAMBIO DE UBICACIONES", Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "dd/MMM/yyyy"), Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "dd/MMM/yyyy"), True)
    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
    Print #1, " Movimiento        Fecha            Usuario         Codigo         Descripcion           Cantidad  Anterior    Nueva     "
    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
    Grid_Reporte.AddItem ""
    Grid_Reporte.AddItem "Movimiento" & Chr(9) & "Fecha" & Chr(9) & "Usuario" & Chr(9) & "Codigo" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Ubicacion Anterior" & Chr(9) & "Ubicacion Nueva"
    'Consulta los datos de cambios de ubicaciones
    Mi_SQL = "SELECT * FROM Alm_Cambio_Ubicaciones"
    Mi_SQL = Mi_SQL & " WHERE Fecha_Cambio BETWEEN '" & Format(DTP_Fecha_Inicio_Facturas_Proveedores.Value, "MM/dd/yyyy") & " 00:00:00' AND '" & Format(DTP_Fecha_Final_Facturas_Proveedores.Value, "MM/dd/yyyy") & " 23:59:59'"
    If Cmb_Facturas_Proveedores.ListIndex > -1 Then
        Mi_SQL = Mi_SQL & " Usuario_Cambio='" & Cmb_Facturas_Proveedores.Text & "'"
    End If
    Set Rs_Consulta_Ubicaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Ubicaciones.EOF
        With Rs_Consulta_Ubicaciones
            If .rdoColumns("Tipo_Producto") = "PRODUCTO TERMINADO" Then
                Mi_SQL = "SELECT Producto_ID,Codigo,Nombre FROM Cat_Productos"
                Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & .rdoColumns("ID_Producto") & "'"
            Else
                Mi_SQL = "SELECT Materia_Prima_ID,Codigo,Nombre FROM Cat_Materias_Primas"
                Mi_SQL = Mi_SQL & " WHERE Materia_Prima_ID='" & .rdoColumns("ID_Producto") & "'"
            End If
            Set Rs_Consulta_Productos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Productos.EOF Then
                Print #1, Trim(.rdoColumns("No_Movimiento")); _
                    Spc(2); Format(.rdoColumns("Fecha_Cambio"), "dd/MMM/yy HH:mm"); _
                    Spc(2); Mid(.rdoColumns("Usuario_Cambio"), 1, 20); _
                    Spc(22 - Len(Mid(.rdoColumns("Usuario_Cambio"), 1, 20))); Mid(Rs_Consulta_Productos.rdoColumns(1), 1, 15); _
                    Spc(16 - Len(Mid(Rs_Consulta_Productos.rdoColumns(1), 1, 15))); Mid(Rs_Consulta_Productos.rdoColumns(2), 1, 22); _
                    Spc(24 - Len(Mid(Rs_Consulta_Productos.rdoColumns(2), 1, 22))); Conectar_Ayudante.Alinea_Derecha_Derecha(.rdoColumns("Cantidad"), 6); _
                    Spc(2); Mid(.rdoColumns("Almacen_Actual"), 1, 10); _
                    Spc(12 - Len(Mid(.rdoColumns("Almacen_Actual"), 1, 10))); .rdoColumns("Almacen_Nuevo")
                Grid_Reporte.AddItem .rdoColumns("No_Movimiento") _
                    & Chr(9) & Format(.rdoColumns("Fecha_Cambio"), "dd/MMMM/yyyy") _
                    & Chr(9) & .rdoColumns("Usuario_Cambio") _
                    & Chr(9) & Rs_Consulta_Productos.rdoColumns(1) _
                    & Chr(9) & Rs_Consulta_Productos.rdoColumns(2) _
                    & Chr(9) & .rdoColumns("Almacen_Actual") _
                    & Chr(9) & .rdoColumns("Almacen_Nuevo")
            End If
        End With
        Rs_Consulta_Ubicaciones.MoveNext
    Wend
    Call Encabezado_Reportes("", "", "", False)
End Sub

Private Sub Btn_Reporte_Productos_Facturados_Cliente_Click()
    Select Case Fra_Facturas_Proveedor.Caption
        Case "Facturas por Proveedor"
            Reporte_Facturas_Por_Proveedor
        Case "Notas de Credito por Proveedor"
            Reporte_Notas_Credito
        Case "Reporte de Cambio de Ubicaciones"
            Reporte_Cambio_Ubicaciones
        Case "Actas de Liberación"
            Reporte_Actas_Liberacion
    End Select
End Sub

Private Sub Btn_Reporte_Saldos_Click()
    Select Case Fra_Saldos.Caption
        Case "Reporte de Saldo por Cliente"
            Reporte_Saldos_Cliente
        Case "Reporte de Saldo por Proveedor"
            Reporte_Saldos_Proveedores
    End Select
End Sub

Private Sub Reporte_Saldos_Proveedores()
Dim Proveedor As String                                     'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Proveedores As rdoResultset             'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Proveedores As rdoResultset    'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Estatus As String                                       'Almacena el estatus de la factura
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Cliente_Saldo.ListIndex > -1 Then
        'Consulta las facturas que fueron emitidas al proveedor seleccionado por el usuario
        Mi_SQL = "SELECT No_Factura,Proveedor_ID,Tipo,Empresa,Fecha,Importe,IVA,Total,Saldo,Cancelada"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
        If Chk_Fechas_Saldos.Value = 1 Then
            Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Fecha_Saldo_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Fecha_Saldo_Fin.Value, "MM/dd/yyyy") & "'"
        End If
        Mi_SQL = Mi_SQL & " ORDER BY Fecha"
        Set Rs_Consulta_Adm_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Adm_Facturas_Proveedores.EOF Then
            Mi_SQL = "SELECT Proveedor_ID,Nombre,Direccion,Colonia,CP,Ciudad_Estado"
            Mi_SQL = Mi_SQL & " FROM Cat_Proveedores"
            Mi_SQL = Mi_SQL & " WHERE Proveedor_ID='" & Format(Cmb_Cliente_Saldo.ItemData(Cmb_Cliente_Saldo.ListIndex), "00000") & "'"
            Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Cat_Proveedores
                If Chk_Fechas_Saldos.Value = 0 Then
                    Call Encabezado_Reportes("REPORTE DE SALDO POR PROVEEDOR", Format(Now, "dd/MMM/yyyy HH:mm:ss"), "", True)
                Else
                    Call Encabezado_Reportes("REPORTE DE SALDO POR PROVEEDOR", Format(Dtp_Fecha_Saldo_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Fecha_Saldo_Fin.Value, "dd/MMM/yyyy"), True)
                End If
                Print #1, Spc(2); "Proveedor:   "; Cmb_Cliente_Saldo.Text
                Print #1, Spc(2); "Direccion:   "; .rdoColumns("Direccion")
                Print #1, Spc(2); "Colonia:     "; .rdoColumns("Colonia"); Spc(2); "CP: "; .rdoColumns("CP")
                Print #1, Spc(2); "Ciudad:      "; .rdoColumns("Ciudad_Estado")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Proveedor:" & Chr(9) & Cmb_Cliente_Saldo.Text
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & .rdoColumns("Direccion")
                Grid_Reporte.AddItem "Colonia:" & Chr(9) & .rdoColumns("Colonia") & Chr(9) & "CP:" & Chr(9) & .rdoColumns("CP")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & .rdoColumns("Ciudad_Estado")
            End With
            Grid_Reporte.AddItem ""
            With Rs_Consulta_Adm_Facturas_Proveedores
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, "No. Factura          Empresa  Tipo           Fecha             Importe           IVA           Total          Saldo      "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Tipo" & Chr(9) & "Fecha" & Chr(9) & "Importe" & _
                                     Chr(9) & "IVA" & Chr(9) & "Total" & Chr(9) & "Saldo"
                While Not .EOF
                    If Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Empresa") = 1 Then
                        Empresa = "NHE"
                    ElseIf Rs_Consulta_Adm_Facturas_Proveedores.rdoColumns("Empresa") = 2 Then
                        Empresa = "PNC"
                    Else
                        Empresa = "GNH"
                    End If
                    If .rdoColumns("Cancelada") = "N" Then
                        Print #1, Mid(Trim(.rdoColumns("No_Factura")), 1, 20); _
                            Spc(22 - Len(Mid(Trim(.rdoColumns("No_Factura")), 1, 20))); Empresa; _
                            Spc(2); Mid(Trim(.rdoColumns("Tipo")), 1, 15); _
                            Spc(17 - Len(Mid(Trim(.rdoColumns("Tipo")), 1, 15))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "#,##0.00"), 14); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "#,##0.00"), 14); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,##0.00"), 14); _
                            Spc(2); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,##0.00"), 14)
                        Grid_Reporte.AddItem .rdoColumns("No_Factura") _
                            & Chr(9) & Empresa _
                            & Chr(9) & .rdoColumns("Tipo") _
                            & Chr(9) & Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") _
                            & Chr(9) & Format(.rdoColumns("Importe"), "#0.00") _
                            & Chr(9) & Format(.rdoColumns("IVA"), "#0.00") _
                            & Chr(9) & Format(.rdoColumns("Total"), "#0.00") _
                            & Chr(9) & Format(.rdoColumns("Saldo"), "#0.00")
                    End If
                    .MoveNext
                Wend
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Call Encabezado_Reportes("", "", "", False)
            End With
        Else
            MsgBox "No hay facturas del Proveedor seleccionado", vbInformation
        End If
        Rs_Consulta_Adm_Facturas_Proveedores.Close
    Else
        MsgBox "Seleccione un proveedor de la lista", vbExclamation
    End If
End Sub

Private Sub Reporte_Tipo_Factura()
Dim Proveedor As String                                 'Obtiene el nombre del proveedor
Dim Rs_Consulta_Cat_Proveedores As rdoResultset          'Consulta el Proveedor al cual pertenece la factura emitida
Dim Rs_Consulta_Cat_Facturas_Proveedores As rdoResultset 'Consulta el tipo de facturas

    Grid_Reporte.Cols = 8
    Grid_Reporte.Rows = 0
    If Cmb_Empresa_Factura_Vencida.ListIndex > -1 Then
        'Consulta las facturas que fueron emitidas durante el periodo de tiempo seleccionado por el usuario
        Mi_SQL = "SELECT No_Factura, Proveedor_ID, Fecha, Importe,IVA, Total"
        Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE Tipo='" & Cmb_Empresa_Factura_Vencida.Text & "'"
        Set Rs_Consulta_Cat_Facturas_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Facturas_Proveedores.EOF Then
            With Rs_Consulta_Cat_Facturas_Proveedores
                Call Encabezado_Reportes("REPORTE DE FACTURAS POR TIPO " & UCase(Cmb_Empresa_Factura_Vencida.Text), Format(Now, "MM/dd/yyyy"), "", True)
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(2); "No. Factura"; Spc(12 - Len("No. Factura")); "Proveedor"; _
                          Spc(42 - Len("Proveedor")); "Fecha"; Spc(15 - Len("Fecha")); "Importe"; Spc(16 - Len("Importe")); "IVA"; _
                          Spc(10 - Len("IVA")); "Total"; Spc(6 - Len("Total"));
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1,
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "IVA" & Chr(9) & "Total"
                Grid_Reporte.AddItem ""
                While Not .EOF
                    If Not IsNull(.rdoColumns("Proveedor_ID")) Then
                        'Consulta el nombre del cliente
                        Mi_SQL = "SELECT Nombre FROM Cat_Proveedores"
                        Mi_SQL = Mi_SQL & " WHERE Proveedor_ID = '" & .rdoColumns("Proveedor_ID") & "'"
                        Set Rs_Consulta_Cat_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Proveedores.EOF Then
                            Proveedor = Rs_Consulta_Cat_Proveedores.rdoColumns("Nombre")
                        End If
                        Rs_Consulta_Cat_Proveedores.Close
                    End If
                    Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(2); Mid(Proveedor, 1, 40); _
                        Spc(42 - Len(Mid(Proveedor, 1, 40))); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                        Spc(14 - Len(Format(.rdoColumns("Fecha"), "dd/MMM/yy"))); _
                        Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "##,##0.00"), 9); _
                        Spc(13 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Importe"), "##,##0.00"), 9))); _
                        Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9); _
                        Spc(11 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("IVA"), "##,##0.00"), 9))); _
                        Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9); _
                        Spc(10 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "##,##0.00"), 9)))
                    Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & "Proveedor" & Chr(9) & _
                        Format(.rdoColumns("Fecha"), "dd/MMMM/yyyy") & Chr(9) & _
                        Format(.rdoColumns("Importe"), "###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("IVA"), "###,##0.00") & Chr(9) & _
                        Format(.rdoColumns("Total"), "##,##0.00") & Chr(13)
                    .MoveNext
                Wend
                Call Encabezado_Reportes("", "", "", False)
            End With
        Else
            MsgBox "No hay facturas del Tipo que proporcionó", vbInformation
        End If
        Rs_Consulta_Cat_Facturas_Proveedores.Close
    Else
        MsgBox "Seleccione un Tipo de Factura", vbInformation
    End If
End Sub

Private Sub Ventas_Vendedor()
Dim Cliente As String                                       'Almacena el nombre del cliente
Dim Rs_Consulta_Cat_Clientes As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Cat_Ciudades As rdoResultset                'Manejo de registro
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Estatus_Facturas As rdoResultset            'Manejo de registro
Dim Rs_Consulta_Detalles_Facturas As rdoResultset
Dim Ciudad As String                                        'Almacena el nombre de la ciudad
Dim Empresa As String                                       'Almacena el nombre de la empresa
Dim Total_Factura As Double                                 'Almacena el total de la facura
Dim Total_Saldo As Double                                   'Almacena el total del saldo
Dim Total_Abono As Double                                   'Almacena el total de los abonos
Dim Total_Cliente As Double
Dim Cliente_ID As String

    Grid_Reporte.Cols = 9
    Grid_Reporte.Rows = 0
    If Cmb_Vendedor_Comision.ListIndex > -1 Then
        'Consulta los clientes del vendedor seleccionado
        Mi_SQL = "SELECT Cliente_ID,Cat_Clientes.Nombre,Cat_Clientes.Direccion,Cat_Clientes.Colonia,Cat_Clientes.CP,Cat_Clientes.Ciudad_ID,Cat_Clientes.Telefono,Cat_Clientes.Vendedor_ID"
        Mi_SQL = Mi_SQL & " FROM Cat_Clientes,Cat_Vendedores"
        Mi_SQL = Mi_SQL & " WHERE Cat_Clientes.Vendedor_ID=Cat_Vendedores.Vendedor_ID"
        Mi_SQL = Mi_SQL & " AND Cat_Clientes.Vendedor_ID='" & Format(Cmb_Vendedor_Comision.ItemData(Cmb_Vendedor_Comision.ListIndex), "00000") & "'"
        Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Clientes.EOF Then
            Call Encabezado_Reportes("REPORTE DE VENTAS POR VENDEDOR", Format(Dtp_Comision_Fecha_Inicio.Value, "dd/MMM/yyyy"), Format(Dtp_Comision_Fecha_Fin.Value, "dd/MMM/yyyy"), True)
            Print #1,
            Print #1, Conectar_Ayudante.Centra("VENDEDOR: " & Cmb_Vendedor_Comision.Text, Len("-------------------------------------------------------------------------------------------------------------------------"))
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "VENDEDOR" & Chr(9) & Cmb_Vendedor_Comision.Text
            While Not Rs_Consulta_Cat_Clientes.EOF
                'Consulta la ciudad del cliente
                If Not IsNull(Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID")) Then
                    Mi_SQL = "SELECT Ciudad_ID,Cat_Ciudades.Nombre,Cat_Ciudades.Estado_ID,Cat_Estados.Nombre_Corto"
                    Mi_SQL = Mi_SQL & " FROM Cat_Ciudades,Cat_Estados"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Ciudades.Estado_ID=Cat_Estados.Estado_ID"
                    Mi_SQL = Mi_SQL & " AND Ciudad_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Ciudad_ID") & "'"
                    Set Rs_Consulta_Cat_Ciudades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Ciudades.EOF Then
                        Ciudad = Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre") & ", " & Rs_Consulta_Cat_Ciudades.rdoColumns("Nombre_Corto")
                    Else
                        Ciudad = ""
                    End If
                    Rs_Consulta_Cat_Ciudades.Close
                Else
                    Ciudad = ""
                End If
                If Cliente_ID <> "" Then
                    Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                    Print #1, " TOTALES CLIENTE "; Spc(53); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 15)
                    Print #1,
                    Grid_Reporte.AddItem ""
                    Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00")
                    Grid_Reporte.AddItem ""
                End If
                Print #1,
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, Spc(2); "Cliente:   "; Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                Print #1, Spc(2); "Direccion: "; Rs_Consulta_Cat_Clientes.rdoColumns("Direccion"); " "; Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                Print #1, Spc(2); "Ciudad:    "; Ciudad; Spc(2); "CP: "; Rs_Consulta_Cat_Clientes.rdoColumns("CP"); Spc(2); "Telefono: "; Rs_Consulta_Cat_Clientes.rdoColumns("Telefono")
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "Cliente:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                Grid_Reporte.AddItem "Dirección:" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Direccion") & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Colonia")
                Grid_Reporte.AddItem "Ciudad:" & Chr(9) & Ciudad & Chr(9) & "CP" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("CP") & Chr(9) & "Telefono" & Chr(9) & Rs_Consulta_Cat_Clientes.rdoColumns("Telefono")
                'Consulta las facturas que fueron emitidas al cliente del vendedor seleccionado
                Mi_SQL = "SELECT No_Factura,Empresa,Cliente_ID,Fecha,Total,Saldo,Abono,Cancelada"
                Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Comision_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Facturas_Clientes.Cliente_ID,Fecha"
                Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, " No Factura  Empresa    Fecha              Abonos            Saldos          Totales                                 "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Factura" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Abonos" & Chr(9) & "Saldos" & Chr(9) & "Totales"
                Cliente_ID = Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID")
                Total_Cliente = 0
                While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                    With Rs_Consulta_Adm_Facturas_Clientes
                        If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                        
                        
                        If Chk_Detalle_Vendededor.Value = 0 Then
                            Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(3); Empresa; _
                                Spc(5); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15)
                                Total_Factura = Total_Factura + .rdoColumns("Total")
                                Total_Cliente = Total_Cliente + .rdoColumns("Total")
                                Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                                Total_Abono = Total_Abono + .rdoColumns("Abono")
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                        Else
                            Print #1, Spc(1); .rdoColumns("No_Factura"); Spc(3); Empresa; _
                                Spc(5); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                                Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15)
                                Total_Factura = Total_Factura + .rdoColumns("Total")
                                Total_Cliente = Total_Cliente + .rdoColumns("Total")
                                Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                                Total_Abono = Total_Abono + .rdoColumns("Abono")
                            Grid_Reporte.AddItem .rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & _
                                Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                                Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                            Print #1, "------------------------------------------------------------------------------------------------------------------------"
                            'Consulta el detalle de la factura
                            Mi_SQL = "SELECT * FROM Adm_Detalles_Facturas"
                            Mi_SQL = Mi_SQL & " WHERE No_Factura='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & "'"
                            Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")
                            Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            While Not Rs_Consulta_Detalles_Facturas.EOF
                                Print #1, Spc(1); Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"); Spc(15 - Len(Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo"))); Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50); Spc(55 - Len(Mid(Conectar_Ayudante.Quitar_Caracter(Rs_Consulta_Detalles_Facturas.rdoColumns("Producto"), Chr(13)), 1, 50))); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0"), 9); Spc(5); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe") + Rs_Consulta_Detalles_Facturas.rdoColumns("IVA"), "#,###,##0.00"), 13)
                                Grid_Reporte.AddItem Rs_Consulta_Detalles_Facturas.rdoColumns("Codigo") & Chr(9) & Rs_Consulta_Detalles_Facturas.rdoColumns("Producto") & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Cantidad"), "#,###,##0") & Chr(9) & Format(Rs_Consulta_Detalles_Facturas.rdoColumns("Importe") + Rs_Consulta_Detalles_Facturas.rdoColumns("IVA"), "#,###,##0.00")
                                Rs_Consulta_Detalles_Facturas.MoveNext
                            Wend
                            Rs_Consulta_Detalles_Facturas.Close
                        End If
                    End With
                    Print #1,
                    Grid_Reporte.AddItem ""
                    Rs_Consulta_Adm_Facturas_Clientes.MoveNext
                Wend
                Rs_Consulta_Adm_Facturas_Clientes.Close
                'Consulta las remisiones que fueron emitidas al cliente del vendedor seleccionado
                Mi_SQL = "SELECT No_Remision,Empresa,Cliente_ID,Fecha,Total,Saldo,Abono,Cancelada"
                Mi_SQL = Mi_SQL & " FROM Adm_Remisiones_Clientes"
                Mi_SQL = Mi_SQL & " WHERE Cancelada='N'"
                Mi_SQL = Mi_SQL & " AND Cliente_ID='" & Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Comision_Fecha_Fin.Value, "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Remisiones_Clientes.Cliente_ID,Fecha"
                Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Print #1,
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Print #1, "No Remision  Empresa    Fecha              Abonos            Saldos          Totales                                 "
                Print #1, "-------------------------------------------------------------------------------------------------------------------------"
                Grid_Reporte.AddItem ""
                Grid_Reporte.AddItem "No. Remision" & Chr(9) & "Empresa" & Chr(9) & "Fecha" & Chr(9) & "Abonos" & Chr(9) & "Saldos" & Chr(9) & "Totales"
                Cliente_ID = Rs_Consulta_Cat_Clientes.rdoColumns("Cliente_ID")
                Total_Cliente = 0
                While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                    With Rs_Consulta_Adm_Facturas_Clientes
                        If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = 1 Then
                            Empresa = "NHE"
                        Else
                            Empresa = "PNC"
                        End If
                        Print #1, Spc(1); .rdoColumns("No_Remision"); Spc(3); Empresa; _
                            Spc(5); Format(.rdoColumns("Fecha"), "dd/MMM/yy"); _
                            Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Abono"), "#,###,##0.00"), 15); _
                            Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Saldo"), "#,###,##0.00"), 15); _
                            Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(.rdoColumns("Total"), "#,###,##0.00"), 15)
                            Total_Factura = Total_Factura + .rdoColumns("Total")
                            Total_Cliente = Total_Cliente + .rdoColumns("Total")
                            Total_Saldo = Total_Saldo + .rdoColumns("Saldo")
                            Total_Abono = Total_Abono + .rdoColumns("Abono")
                        Grid_Reporte.AddItem .rdoColumns("No_Remision") & Chr(9) & Empresa & Chr(9) & _
                            Format(.rdoColumns("Fecha"), "dd/MMM/yyyy") & Chr(9) & Format(.rdoColumns("Abono"), "#,###,##0.00") & Chr(9) & _
                            Format(.rdoColumns("Saldo"), "#,###,##0.00") & Chr(9) & Format(.rdoColumns("Total"), "#,###,##0.00")
                    End With
                    Rs_Consulta_Adm_Facturas_Clientes.MoveNext
                Wend
                Rs_Consulta_Adm_Facturas_Clientes.Close
                Rs_Consulta_Cat_Clientes.MoveNext
            Wend
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " TOTALES CLIENTE "; Spc(53); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Cliente, "#,###,##0.00"), 15)
            Print #1,
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES CLIENTE" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Format(Total_Cliente, "#,###,##0.00")
            Grid_Reporte.AddItem ""
            Print #1,
            Print #1, "-------------------------------------------------------------------------------------------------------------------------"
            Print #1, " TOTALES GENERALES "; Spc(15); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Abono, "#,###,##0.00"), 15); _
                        Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Saldo, "#,###,##0.00"), 15); _
                        Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Factura, "#,###,##0.00"), 15)
            Call Encabezado_Reportes("", "", "", False)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem "TOTALES GENERALES" & Chr(9) & Chr(9) & Chr(9) & Format(Total_Abono, "#,###,##0.00") & Chr(9) & Format(Total_Saldo, "#,###,##0.00") & Chr(9) & Format(Total_Factura, "#,###,##0.00")
        Else
            MsgBox "No hay datos que consultar", vbExclamation
        End If
        Rs_Consulta_Cat_Clientes.Close
    Else
        MsgBox "Seleccione un vendedor de la lista", vbExclamation
    End If
End Sub

Private Sub Btn_Salir_Click()
Dim Mi_Frame As Object    'Indica que objeto se esta consultando
Dim Frame_Visible As Boolean
    For Each Mi_Frame In Me.Controls
        If TypeOf Mi_Frame Is Frame Then
            If Mi_Frame.Visible = True Then
                Frame_Visible = True
            End If
        End If
    Next
    If Frame_Visible = True Then
        Edm_Reporte.Visible = False
        'Btn_Imprimir.Visible = False
        'Btn_Exportar_Excel.Visible = False
        'Btn_Guardar_Disco.Visible = False
        'Btn_Salir.Visible = False
        Pic_Reportes.Visible = True
        If Me.WindowState <> 2 Then
            Me.Width = Ancho_Frame
            Me.Height = Largo_Frame
            Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
            Me.Top = 100
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Chk_Estatus_Materia_Prima_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Estatus_Entradas As rdoResultset
    If Chk_Estatus_Materia_Prima.Value = 1 Then
        Mi_SQL = "SELECT DISTINCT Estatus FROM Alm_Entradas_Materia_Prima_Detalles"
        Set Rs_Consulta_Estatus_Entradas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Estatus_Materia_Prima.Clear
        Cmb_Estatus_Materia_Prima.AddItem "TODOS"
        While Not Rs_Consulta_Estatus_Entradas.EOF
            Cmb_Estatus_Materia_Prima.AddItem Rs_Consulta_Estatus_Entradas.rdoColumns("Estatus")
            Rs_Consulta_Estatus_Entradas.MoveNext
        Wend
        Rs_Consulta_Estatus_Entradas.Close
    End If
End Sub

Private Sub Chk_Fechas_Saldos_Click()
    If Chk_Fechas_Saldos.Value = 1 Then
        Dtp_Fecha_Saldo_Inicio.Enabled = True
        Dtp_Fecha_Saldo_Inicio.Value = Now
        Dtp_Fecha_Saldo_Fin.Enabled = True
        Dtp_Fecha_Saldo_Fin.Value = Now
    Else
        Dtp_Fecha_Saldo_Inicio.Enabled = False
        Dtp_Fecha_Saldo_Fin.Enabled = False
    End If
End Sub

Private Sub Chk_General_Click()
    If Chk_General.Value = 0 Then
        Dtp_Fecha_Inicio_Facturas_Directas.Enabled = True
        Dtp_Fecha_Fin_Facturas_Directas.Enabled = True
    Else
        Dtp_Fecha_Inicio_Facturas_Directas.Enabled = False
        Dtp_Fecha_Fin_Facturas_Directas.Enabled = False
    End If
End Sub

Public Sub Cmb_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Banco, 1, "Nombre")
        Cmb_Banco.Text = ""
    End If
End Sub

Private Sub Cmb_Ciudad_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Ciudad_ID,Nombre", "Cat_Ciudades", Cmb_Ciudad_Venta, 1, "Nombre")
        Cmb_Ciudad_Venta.Text = ""
    End If
End Sub

Public Sub Cmb_Cliente_Anticipo_KeyPress(KeyAscii As Integer)
Dim Mi_SQL As String
Dim Rs_Adm_Cuenta_Ajustes As rdoResultset

    If KeyAscii = 13 Then
        Select Case Fra_Anticipos_Clientes.Caption
            Case "Anticipos Clientes"
                Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Anticipo, 1, "Nombre")
            Case "Anticipos Proveedores"
                Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Cliente_Anticipo, 1, "Nombre")
            Case "Reporte de Pago de Saldos Pequeños"
                Mi_SQL = "SELECT No_Cuenta,Empresa,Cliente FROM Adm_Cuentas_Ajustes"
                Mi_SQL = Mi_SQL & " WHERE Cliente LIKE '%" & Cmb_Cliente_Anticipo.Text & "%'"
                Set Rs_Adm_Cuenta_Ajustes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Cliente_Anticipo.Clear
                While Not Rs_Adm_Cuenta_Ajustes.EOF
                    If Rs_Adm_Cuenta_Ajustes.rdoColumns("Empresa") = 1 Then
                        Cmb_Cliente_Anticipo.AddItem "NHE - " & Rs_Adm_Cuenta_Ajustes.rdoColumns("Cliente")
                    Else
                        Cmb_Cliente_Anticipo.AddItem "PNC - " & Rs_Adm_Cuenta_Ajustes.rdoColumns("Cliente")
                    End If
                    Cmb_Cliente_Anticipo.ItemData(Cmb_Cliente_Anticipo.NewIndex) = Rs_Adm_Cuenta_Ajustes.rdoColumns("No_Cuenta")
                    Rs_Adm_Cuenta_Ajustes.MoveNext
                Wend
                Rs_Adm_Cuenta_Ajustes.Close
        End Select
        Cmb_Cliente_Anticipo.Text = ""
    End If
End Sub

Private Sub Cmb_Cliente_Devoluciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Devoluciones, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Cliente_Envio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Envio, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Cliente_Estadistico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fra_Estadistico_Cliente.Caption = "Estadistico Cliente" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Estadistico, 1, "Nombre")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Ciudad_ID,Nombre", "Cat_Ciudades", Cmb_Cliente_Estadistico, 1, "Nombre")
        End If
        Cmb_Cliente_Estadistico.Text = ""
    End If
End Sub

Private Sub Cmb_Cliente_Factura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_Id,Nombre", "Cat_Clientes", Cmb_Cliente_Factura, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Cliente_Saldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Fra_Saldos.Caption
            Case "Reporte de Saldo por Cliente"
                If Cmb_Vendedor_Saldo.ListIndex > -1 Then
                    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Saldo, 1, "Nombre", " AND Vendedor_ID='" & Format(Cmb_Vendedor_Saldo.ItemData(Cmb_Vendedor_Saldo.ListIndex), "00000") & "'")
                Else
                    Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Cliente_Saldo, 1, "Nombre")
                End If
            Case "Reporte de Saldo por Proveedor"
                Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Cliente_Saldo, 1, "Nombre")
        End Select
        Cmb_Cliente_Saldo.Text = ""
    End If
End Sub

Private Sub Cmb_Cuenta_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Banco_ID,Nombre", "Cat_Bancos", Cmb_Cuenta_Banco, 1, "Nombre")
    End If
End Sub

Public Sub Cmb_Empresa_Cobranza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Lbl_Empresa_Cobranza.Caption = "Cliente" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Empresa_Cobranza, 1, "Nombre")
            Cmb_Empresa_Cobranza.Text = ""
        End If
    End If
End Sub

Private Sub Cmb_Estado_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Estado_ID,Nombre", "Cat_Estados", Cmb_Estado_Venta, 1, "Nombre")
        Cmb_Estado_Venta.Text = ""
    End If
End Sub

Private Sub Cmb_Facturas_Proveedores_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Lbl_Proveedor_Reporte.Caption = "Proveedor" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Facturas_Proveedores, 1, "Nombre")
        Else
            Call Conectar_Ayudante.Llena_Combo("Nombre", "Cat_Usuarios", Cmb_Facturas_Proveedores, 1, "Nombre")
        End If
        Cmb_Facturas_Proveedores.Text = ""
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Medio_Envio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Medio_ID,Nombre", "Cat_Medios_Envio", Cmb_Medio_Envio, 1, "Nombre")
    End If
End Sub


Private Sub Cmb_Nota_Credito_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Cliente_Id,Nombre", "Cat_Clientes", Cmb_Nota_Credito_Cliente, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Producto_KeyPress(KeyAscii As Integer)
Dim Mi_SQL As String
Dim Rs_Consulta_Producto As rdoResultset    'Manejo de registro
    
    If KeyAscii = 13 Then
        'Verifica que tipo de inventario es para hacer la consulta a la tabla correspondiente
        If Fra_Kardex.Caption = "Kardex de Producto Terminado" Then
            'Consulta el campo
            Mi_SQL = "SELECT Producto_ID,Codigo,Nombre FROM Cat_Productos"
            Mi_SQL = Mi_SQL & " WHERE (Nombre LIKE '%" & Cmb_Producto.Text & "%'"
            Mi_SQL = Mi_SQL & " OR Codigo LIKE '%" & Cmb_Producto.Text & "%')"
            Mi_SQL = Mi_SQL & " AND Tipo_Producto_ID<>'00016'"
            Mi_SQL = Mi_SQL & " ORDER BY Codigo"
        Else
            Mi_SQL = "SELECT Materia_Prima_ID,Codigo,Nombre FROM Cat_Materias_Primas"
            Mi_SQL = Mi_SQL & " WHERE (Nombre LIKE '%" & Cmb_Producto.Text & "%'"
            Mi_SQL = Mi_SQL & " OR Codigo LIKE '%" & Cmb_Producto.Text & "%')"
            Mi_SQL = Mi_SQL & " ORDER BY Codigo"
        End If
        Set Rs_Consulta_Producto = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena el ComboBox de la forma
        If Not Rs_Consulta_Producto.EOF Then
            Cmb_Producto.Clear
            While Not Rs_Consulta_Producto.EOF
                Cmb_Producto.AddItem Rs_Consulta_Producto.rdoColumns("Codigo") & " - " & Rs_Consulta_Producto.rdoColumns("Nombre")
                Cmb_Producto.ItemData(Cmb_Producto.NewIndex) = Rs_Consulta_Producto.rdoColumns(0)
                Rs_Consulta_Producto.MoveNext
            Wend
        End If
        Rs_Consulta_Producto.Close
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Producto_Materia_Prima_KeyPress(KeyAscii As Integer)
Dim Rs_Consulta_Productos As rdoResultset
    If KeyAscii = 13 Then
        If Lbl_Producto_Materia_Prima.Caption = "Producto" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Producto_Materia_Prima, 1, "Nombre")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Materia_Prima_ID,Nombre", "Cat_Materias_Primas", Cmb_Producto_Materia_Prima, 1, "Nombre")
        End If
        Cmb_Producto_Materia_Prima.Text = ""
    End If
End Sub

Private Sub Cmb_Producto_Venta_KeyPress(KeyAscii As Integer)
Dim Rs_Materias_Primas As rdoResultset
    If KeyAscii = 13 Then
        Select Case Chk_Producto.Caption
            Case "Producto"
                Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Producto_Venta, 1, "Nombre", "AND Tipo_Producto_ID<>'00016'")
            Case "Clientes"
                Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Producto_Venta, 1, "Nombre")
            Case "M. Prima"
                Mi_SQL = "SELECT Materia_Prima_ID,Codigo,Nombre FROM Cat_Materias_Primas"
                Mi_SQL = Mi_SQL & " WHERE Estatus='ACTIVO'"
                Mi_SQL = Mi_SQL & " AND (Nombre LIKE '%" & Cmb_Producto_Venta.Text & "%'"
                Mi_SQL = Mi_SQL & " OR Codigo LIKE '%" & Cmb_Producto_Venta.Text & "%')"
                Mi_SQL = Mi_SQL & " ORDER BY Codigo"
                Set Rs_Materias_Primas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                Cmb_Producto_Venta.Clear
                While Not Rs_Materias_Primas.EOF
                    Cmb_Producto_Venta.AddItem Rs_Materias_Primas.rdoColumns("Nombre")
                    Cmb_Producto_Venta.ItemData(Cmb_Producto_Venta.NewIndex) = Rs_Materias_Primas.rdoColumns("Materia_Prima_ID")
                    Rs_Materias_Primas.MoveNext
                Wend
                Rs_Materias_Primas.Close
                Cmb_Producto_Venta.Text = ""
        End Select
        Cmb_Producto_Venta.Text = ""
    End If
End Sub

Private Sub Cmb_Proveedor_Entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Proveedor_ID,Nombre", "Cat_Proveedores", Cmb_Proveedor_Entrada, 1, "Nombre")
        Cmb_Proveedor_Entrada.Text = ""
    End If
End Sub

Private Sub Cmb_Reporte_Pedido_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Lbl_Reporte_Pedido_Producto.Caption = "Producto" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Producto_ID,Nombre", "Cat_Productos", Cmb_Reporte_Pedido_Producto, 1, "Nombre")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Cliente_ID,Nombre", "Cat_Clientes", Cmb_Reporte_Pedido_Producto, 1, "Nombre")
        End If
        Cmb_Reporte_Pedido_Producto.Text = ""
    End If
End Sub

Public Sub Cmb_Tipo_Entrada_KeyPress(KeyAscii As Integer)
    Select Case Fra_Reporte_Entradas_Salidas.Caption
        Case "Reporte de Entradas Materia Prima"
            Cmb_Tipo_Entrada.Clear
            Cmb_Tipo_Entrada.AddItem "ORDEN PRODUCCION"
            Cmb_Tipo_Entrada.AddItem "TRASPASO"
            Cmb_Tipo_Entrada.AddItem "MUESTRAS"
            Cmb_Tipo_Entrada.AddItem "FACTURA"
            Cmb_Tipo_Entrada.AddItem "REMISION"
            Cmb_Tipo_Entrada.AddItem "MAQUILA"
            Cmb_Tipo_Entrada.AddItem "ORDEN COMPRA"
            Cmb_Tipo_Entrada.AddItem "AJUSTE"
            Cmb_Tipo_Entrada.AddItem "OTROS"
        Case "Reporte de Salidas Producto Terminado"
            Call Conectar_Ayudante.Llena_Combo("DISTINCT Tipo", "Ope_Salidas", Cmb_Tipo_Entrada, 0, "")
        Case "Reporte de Salidas Materia Prima"
            Call Conectar_Ayudante.Llena_Combo("DISTINCT Tipo_Salida", "Alm_Salidas_Materia_Prima", Cmb_Tipo_Entrada, 0, "")
    End Select
End Sub

Public Sub Cmb_Tipo_Nota_Credito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Nota_Credito_ID,Descripcion", "Cat_Tipos_Notas_Credito", Cmb_Tipo_Nota_Credito, 1, "Descripcion")
        Cmb_Tipo_Nota_Credito.Text = ""
    End If
End Sub

Public Sub Cmb_Tipo_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Producto_ID,Nombre", "Cat_Tipos_Productos", Cmb_Tipo_Producto, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Vendedor_Comision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Vendedor_Id,Nombre", "Cat_Vendedores", Cmb_Vendedor_Comision, 1, "Nombre")
    End If
End Sub

Private Sub Cmb_Vendedor_Saldo_Click()
    Cmb_Cliente_Saldo.Clear
    Call Cmb_Cliente_Saldo_KeyPress(13)
End Sub

Private Sub Cmb_Vendedor_Saldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Vendedor_ID,Nombre", "Cat_Vendedores", Cmb_Vendedor_Saldo, 1, "Nombre")
        Cmb_Vendedor_Saldo.Text = ""
    End If
End Sub


Private Sub Cmb_Zona_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Zona_ID,Nombre", "Cat_Zonas", Cmb_Zona_Venta, 1, "Nombre")
        Cmb_Zona_Venta.Text = ""
    End If
End Sub



Private Sub Form_Load()
    Me.Width = 5025
    Me.Height = 3570
    Me.Left = (MDIFrm_Apl_Principal.Width - Me.Width) / 2
    Me.Top = 100
End Sub

Private Sub Comisiones_Vendedores()
Dim Encabezado As Boolean                               'Indica si fue impreso el encabezado del reporte
Dim Dia As Integer                                      'Indica que día es el que se esta consultando
Dim Dias As Integer                                     'Indica la cantidad de días a consultar
Dim Total_Ventas As Double                              'Indica el total de las ventas obtenidas
Dim Total As Double                                     'Consulta el total de las ventas realizadas durante el día
Dim Comision As Double                                  'Almacena el total de la comision por factura
Dim Total_Comision As Double                            'Almacena el total de las comisiones por empleado
Dim Rs_Consulta_Adm_Facturas_Clientes As rdoResultset   'Manejo de registro
Dim Rs_Consulta_Detalles_Facturas As rdoResultset       'Manejo de registro
Dim Rs_Consulta_Cat_Clientes As rdoResultset            'Manejo de registro
Dim Cliente As String                                   'Almacena el nombre del cliente
Dim Empresa As String                                   'Almacena el nombre de la empresa
    
    Encabezado = False
    Total_Ventas = 0
    If Cmb_Vendedor_Comision.ListIndex > -1 Then
        Dias = DateDiff("d", Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy"))
        For Dia = 0 To Dias
            If Dia > 0 Then Dtp_Comision_Fecha_Inicio.Value = DateAdd("d", 1, Dtp_Comision_Fecha_Fin)
            Total = 0
            'Consulta las facturas realizadas durante el día que se esta seleccionado
            Mi_SQL = "SELECT Adm_Facturas_Clientes.Empresa,Adm_Facturas_Clientes.No_Factura,Adm_Facturas_Clientes.Pagada,Adm_Facturas_Clientes.Cliente_ID,Adm_Facturas_Clientes.Total,Cat_Clientes.Vendedor_ID,Cat_Vendedores.Vendedor_ID,Cat_Clientes.Cliente_ID,Adm_Facturas_Clientes.Fecha"
            Mi_SQL = Mi_SQL & " FROM Adm_Facturas_Clientes,Cat_Clientes,Cat_Vendedores"
            Mi_SQL = Mi_SQL & " WHERE Cat_Vendedores.Vendedor_ID = '" & Format(Cmb_Vendedor_Comision.ItemData(Cmb_Vendedor_Comision.ListIndex), "00000") & "' AND Adm_Facturas_Clientes.Cliente_ID = Cat_Clientes.Cliente_ID"
            Mi_SQL = Mi_SQL & " AND Cat_Vendedores.Vendedor_ID = Cat_Clientes.Vendedor_ID "
            Mi_SQL = Mi_SQL & " AND Adm_Facturas_Clientes.Pagada = 'S'"
            Mi_SQL = Mi_SQL & " AND Fecha >= " & Par_Fecha & Format(Dtp_Comision_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(Dtp_Comision_Fecha_Fin.Value, "MM/dd/yyyy") & Par_Fecha
            Set Rs_Consulta_Adm_Facturas_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Adm_Facturas_Clientes.EOF Then
            While Not Rs_Consulta_Adm_Facturas_Clientes.EOF
                If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")) Then
                    Total = Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Total")
                End If
                'Consulta el cliente
                If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID")) Then
                    Mi_SQL = "SELECT Nombre,Cliente_ID FROM Cat_Clientes"
                    Mi_SQL = Mi_SQL & " WHERE Cliente_ID='" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Cliente_ID") & "'"
                    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Clientes.EOF Then
                        Cliente = Rs_Consulta_Cat_Clientes.rdoColumns("Nombre")
                    Else
                        Cliente = ""
                    End If
                    Rs_Consulta_Cat_Clientes.Close
                End If
                
                If Not IsNull(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa")) Then
                    If Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Empresa") = "1" Then
                        Empresa = "NHE"
                    Else
                        Empresa = "PNC"
                    End If
                End If
                Mi_SQL = " SELECT * FROM Adm_Detalles_Facturas"
                Mi_SQL = Mi_SQL & " WHERE No_Factura = '" & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & "'"
                Set Rs_Consulta_Detalles_Facturas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                While Not Rs_Consulta_Detalles_Facturas.EOF
                    If Rs_Consulta_Detalles_Facturas.rdoColumns("No_Oferta") > 0 Then
                        Comision = Total * PG_Comision_Promocion
                        Rs_Consulta_Detalles_Facturas.MoveLast
                    Else
                        Comision = Total * PG_Comision_Completa
                    End If
                    Rs_Consulta_Detalles_Facturas.MoveNext
                Wend
                Rs_Consulta_Detalles_Facturas.Close
                If Total > 0 Then
                    Total_Ventas = Total_Ventas + Total
                    If Encabezado = False Then
                        Call Encabezado_Reportes("REPORTE DE COMISION POR VENTAS", Format(Dtp_Comision_Fecha_Inicio, "dd/MMM/yyyy"), Format(Dtp_Comision_Fecha_Fin, "dd/MMM/yyyy"), True)
                        Print #1, Spc(50); "Vendedor:"; Cmb_Vendedor_Comision.Text
                        Grid_Reporte.AddItem ""
                        Grid_Reporte.AddItem "Vendedor:" & Chr(9) & Cmb_Vendedor_Comision.Text
                        Grid_Reporte.AddItem ""
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1, " No. Factura"; Spc(2); "Empresa"; Spc(3); "Cliente"; Spc(30 - Len("Cliente")); Spc(3); "Fecha Venta"; Spc(10 - Len("Fecha")); Spc(2); "Importe"; Spc(5); "Comision"
                        Print #1, "------------------------------------------------------------------------------------------------------------------------"
                        Print #1,
                        Grid_Reporte.AddItem "No_Factura" & Chr(9) & "Empresa" & Chr(9) & "Cliente" & Chr(9) & "Fecha" & Chr(9) & "Importe" & Chr(9) & "Comision"
                        Encabezado = True
                    End If
                    Print #1, Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura"); Spc(4); Empresa; Spc(6 - Len(Mid(Empresa, 1, 6))); Spc(2); Mid(Cliente, 1, 30); _
                    Spc(35 - Len(Mid(Cliente, 1, 30))); Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"); _
                    Spc(12 - Len(Format(Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha"), "dd/MMM/yyyy"))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 12); _
                    Spc(12 - Len(Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total, "#,###,##0.00"), 12))); _
                    Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Comision, "#,###,##0.00"), 12)
                    Grid_Reporte.AddItem Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("No_Factura") & Chr(9) & Empresa & Chr(9) & Cliente & Chr(9) & Rs_Consulta_Adm_Facturas_Clientes.rdoColumns("Fecha") & Chr(9) & Format(Total, "#,###,##0.00") & Chr(9) & Format(Comision, "#,###,##0.00")
                End If
                Total_Comision = Total_Comision + Comision
            Rs_Consulta_Adm_Facturas_Clientes.MoveNext
            Wend
            End If
            Rs_Consulta_Adm_Facturas_Clientes.Close
        Next Dia
        If Encabezado = True Then
            Print #1, Spc(71); "------------"
            Print #1, Spc(54); "Total Ventas"; Spc(3); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Ventas, "#,###,##0.00"), 12); Conectar_Ayudante.Alinea_Derecha_Derecha(Format(Total_Comision, "#,###,##0.00"), 12)
            Grid_Reporte.AddItem ""
            Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Venta : " & Chr(9) & Total_Ventas
            Grid_Reporte.AddItem Chr(9) & Chr(9) & "Total Comisiones :" & Chr(9) & Total_Comision
            Call Encabezado_Reportes("", "", "", False)
        Else
            MsgBox "No hay datos que consultar", vbInformation
        End If
    Else
        MsgBox "Seleccione un Vendedor de la lista", vbInformation
    End If
End Sub

Private Sub Txt_No_Cheque_Change()
    Btn_Reporte_Gastos.Default = True
End Sub

Private Sub Txt_No_Cheque_LostFocus()
    Btn_Reporte_Gastos.Default = False
End Sub

Private Sub Txt_No_Factura_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii)
End Sub

Private Sub Txt_Objetivo_Efectividad_1_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Objetivo_Efectividad_1.Text, True)
End Sub

Private Sub Txt_Objetivo_Efectividad_2_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Objetivo_Efectividad_2.Text, True)
End Sub

Private Sub Txt_Objetivo_Efectividad_3_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Objetivo_Efectividad_3.Text, True)
End Sub

Private Sub Txt_Objetivo_Efectividad_4_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Objetivo_Efectividad_4.Text, True)
End Sub

