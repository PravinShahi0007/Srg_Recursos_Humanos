VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_Cat_Generales_RH 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9270
   Begin VB.PictureBox Pic_Cat_Calendarios_Turnos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   11565
      TabIndex        =   197
      Top             =   0
      Width           =   11565
      Begin VB.Frame Fra_Calendarios_Turnos_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1425
         Left            =   75
         TabIndex        =   223
         Top             =   360
         Width           =   11370
         Begin VB.TextBox Txt_Calendario_Estatus 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   227
            Top             =   195
            Width           =   1305
         End
         Begin VB.TextBox Txt_Calendario_Comentarios 
            Height          =   330
            Left            =   1500
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   226
            Top             =   975
            Width           =   9765
         End
         Begin VB.TextBox Txt_Calendario_Nombre 
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   225
            Top             =   600
            Width           =   9765
         End
         Begin VB.TextBox Txt_Calendario_Turno_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   224
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   230
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   8
            Left            =   135
            TabIndex        =   229
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calendario ID"
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
            Index           =   7
            Left            =   135
            TabIndex        =   228
            Top             =   255
            Width           =   1170
         End
      End
      Begin TabDlg.SSTab Tab_Calendarios_Turnos 
         Height          =   4365
         Left            =   75
         TabIndex        =   198
         Top             =   1890
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   7699
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Calendarios"
         TabPicture(0)   =   "Frm_Cat_Generales_RH.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Calendarios_Turnos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Configuración"
         TabPicture(1)   =   "Frm_Cat_Generales_RH.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Grid_Calendarios_Configuracion_Turnos"
         Tab(1).Control(1)=   "Fra_Calendarios_Configuración_Turnos"
         Tab(1).ControlCount=   2
         Begin MSFlexGridLib.MSFlexGrid Grid_Calendarios_Configuracion_Turnos 
            Height          =   2550
            Left            =   -74790
            TabIndex        =   233
            Top             =   1560
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   4498
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
         Begin VB.Frame Fra_Calendarios_Turnos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Calendarios"
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
            Left            =   120
            TabIndex        =   218
            Top             =   360
            Width           =   11190
            Begin MSFlexGridLib.MSFlexGrid Grid_Calendarios_Turnos 
               Height          =   3570
               Left            =   75
               TabIndex        =   219
               Top             =   240
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   6297
               _Version        =   393216
               Rows            =   0
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Fra_Calendarios_Configuración_Turnos 
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
            Height          =   3945
            Left            =   -74880
            TabIndex        =   199
            Top             =   360
            Width           =   11190
            Begin VB.TextBox Txt_Calendario_Filtro_Empleados 
               Height          =   315
               Left            =   7920
               MaxLength       =   50
               TabIndex        =   232
               Top             =   1320
               Width           =   3105
            End
            Begin VB.ListBox Lst_Calendarios_Configuracion_Empleados 
               Height          =   2085
               ItemData        =   "Frm_Cat_Generales_RH.frx":0038
               Left            =   7920
               List            =   "Frm_Cat_Generales_RH.frx":003F
               Style           =   1  'Checkbox
               TabIndex        =   222
               Top             =   1680
               Width           =   3135
            End
            Begin VB.TextBox Txt_Calendarios_Configuracion_Turno 
               Height          =   315
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   221
               Top             =   120
               Width           =   6525
            End
            Begin VB.TextBox Txt_Calendario_Horas_Comida 
               Height          =   315
               Left            =   9555
               MaxLength       =   50
               TabIndex        =   203
               Top             =   810
               Width           =   1545
            End
            Begin VB.TextBox Txt_Calendario_Horas_Turno 
               Height          =   315
               Left            =   9555
               MaxLength       =   50
               TabIndex        =   202
               Top             =   480
               Width           =   1545
            End
            Begin VB.CommandButton Btn_Configuracion_Calendarios_Turnos_Limpiar_Datos 
               Caption         =   "Limpiar Datos"
               Height          =   315
               Left            =   7995
               TabIndex        =   201
               ToolTipText     =   "Limpiar Datos"
               Top             =   120
               Width           =   1545
            End
            Begin VB.CommandButton Btn_Configuracion_Calendarios_Turnos_Limpiar_Calendario 
               Caption         =   "Limpiar Calendario"
               Height          =   315
               Left            =   9555
               TabIndex        =   200
               ToolTipText     =   "Limpiar Calendario"
               Top             =   120
               Width           =   1545
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Hora_Inicio 
               Height          =   315
               Left            =   1320
               TabIndex        =   204
               Top             =   840
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Hora_Termino 
               Height          =   315
               Left            =   4080
               TabIndex        =   205
               Top             =   840
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Inicio_Comida 
               Height          =   315
               Left            =   6930
               TabIndex        =   206
               Top             =   480
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Termino_Comida 
               Height          =   315
               Left            =   6930
               TabIndex        =   207
               Top             =   840
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Fecha_Inicio 
               Height          =   315
               Left            =   1320
               TabIndex        =   208
               Top             =   480
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   123666435
               CurrentDate     =   39941
            End
            Begin MSComCtl2.DTPicker Dtp_Calendario_Fecha_Termino 
               Height          =   315
               Left            =   4080
               TabIndex        =   209
               Top             =   480
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy"
               Format          =   123666435
               CurrentDate     =   39941
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Index           =   14
               Left            =   120
               TabIndex        =   220
               Top             =   120
               Width           =   510
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Horas Comida"
               Height          =   195
               Left            =   7995
               TabIndex        =   217
               Top             =   810
               Width           =   990
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Horas Turno"
               Height          =   195
               Left            =   7995
               TabIndex        =   216
               Top             =   480
               Width           =   885
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio Comida"
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
               Index           =   13
               Left            =   5520
               TabIndex        =   215
               Top             =   450
               Width           =   1275
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Inicio"
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
               Index           =   12
               Left            =   120
               TabIndex        =   214
               Top             =   840
               Width           =   945
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Término"
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
               Index           =   11
               Left            =   2760
               TabIndex        =   213
               Top             =   840
               Width           =   1140
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Término Comida"
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
               Index           =   10
               Left            =   5520
               TabIndex        =   212
               Top             =   810
               Width           =   1350
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Inicio"
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
               Index           =   15
               Left            =   120
               TabIndex        =   211
               Top             =   480
               Width           =   1065
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Término"
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
               Left            =   2760
               TabIndex        =   210
               Top             =   480
               Width           =   1260
            End
         End
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CALENDARIOS TURNOS"
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
         Left            =   75
         TabIndex        =   231
         Top             =   0
         Width           =   11370
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Btn_Salir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   555
      Left            =   7140
      Picture         =   "Frm_Cat_Generales_RH.frx":006C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6375
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton Btn_Consultar 
      Caption         =   "Consultar"
      Height          =   555
      Left            =   5430
      Picture         =   "Frm_Cat_Generales_RH.frx":05F6
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "C"
      Top             =   6375
      Width           =   1305
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   555
      Left            =   1905
      Picture         =   "Frm_Cat_Generales_RH.frx":0B80
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "M"
      Top             =   6375
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   555
      Left            =   3675
      Picture         =   "Frm_Cat_Generales_RH.frx":110A
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "B"
      Top             =   6375
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   165
      Picture         =   "Frm_Cat_Generales_RH.frx":1694
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "A"
      Top             =   6375
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.PictureBox Pic_Cat_Empresas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8640
      TabIndex        =   39
      Top             =   0
      Width           =   8640
      Begin VB.Frame Fra_Cat_Empresas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   100
         TabIndex        =   40
         Top             =   3690
         Width           =   8220
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Empresas 
            Height          =   2220
            Left            =   105
            TabIndex        =   18
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   3916
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin TabDlg.SSTab Tab_Cat_Empresas 
         Height          =   3240
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   5715
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Generales"
         TabPicture(0)   =   "Frm_Cat_Generales_RH.frx":1C1E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Cat_Empresas_Datos_Generales"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Equipos Identificacion"
         TabPicture(1)   =   "Frm_Cat_Generales_RH.frx":1C3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Cat_Empresas_Equipos"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Eq. dentificacion Almacenes"
         TabPicture(2)   =   "Frm_Cat_Generales_RH.frx":1C56
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Fra_Cat_Empresas_Equipos_Almacenes"
         Tab(2).ControlCount=   1
         Begin VB.Frame Fra_Cat_Empresas_Equipos_Almacenes 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2820
            Left            =   -74940
            TabIndex        =   191
            Top             =   420
            Width           =   8100
            Begin VB.ComboBox Cmb_Cat_Empresas_Equipo_Almacenes 
               Height          =   315
               ItemData        =   "Frm_Cat_Generales_RH.frx":1C72
               Left            =   725
               List            =   "Frm_Cat_Generales_RH.frx":1C7C
               Style           =   2  'Dropdown List
               TabIndex        =   194
               Top             =   180
               Width           =   5340
            End
            Begin VB.CommandButton Btn_Cat_Empresas_Agregar_Equipo_Almacenes 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   6175
               TabIndex        =   193
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton Btn_Cat_Empresas_Eliminar_Equipo_Almacenes 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   7140
               TabIndex        =   192
               Top             =   180
               Width           =   855
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Empresas_Equipos_Almacenes 
               Height          =   2160
               Left            =   45
               TabIndex        =   195
               Top             =   585
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   3810
               _Version        =   393216
               Rows            =   0
               Cols            =   3
               FixedRows       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Equipo"
               Height          =   195
               Left            =   120
               TabIndex        =   196
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Fra_Cat_Empresas_Equipos 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2820
            Left            =   -74940
            TabIndex        =   55
            Top             =   420
            Width           =   8100
            Begin VB.CommandButton Btn_Cat_Empresas_Eliminar_Equipo 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   7140
               TabIndex        =   17
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton Btn_Cat_Empresas_Agregar_Equipo 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   6175
               TabIndex        =   15
               Top             =   180
               Width           =   855
            End
            Begin VB.ComboBox Cmb_Cat_Empresas_Equipo 
               Height          =   315
               ItemData        =   "Frm_Cat_Generales_RH.frx":1C94
               Left            =   725
               List            =   "Frm_Cat_Generales_RH.frx":1C9E
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   180
               Width           =   5340
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Empresas_Equipos 
               Height          =   2160
               Left            =   45
               TabIndex        =   16
               Top             =   585
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   3810
               _Version        =   393216
               Rows            =   0
               Cols            =   3
               FixedRows       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Equipo"
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Fra_Cat_Empresas_Datos_Generales 
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
            Height          =   2805
            Left            =   60
            TabIndex        =   42
            Top             =   420
            Width           =   8070
            Begin VB.CommandButton Btn_Ruta_Logo 
               Caption         =   "..."
               Height          =   255
               Left            =   7560
               TabIndex        =   173
               Top             =   1980
               Width           =   375
            End
            Begin VB.TextBox Txt_Logo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   6165
               TabIndex        =   172
               Top             =   1980
               Width           =   1335
            End
            Begin VB.TextBox Txt_Cat_Empresas_Estado 
               Height          =   285
               Left            =   6165
               MaxLength       =   50
               TabIndex        =   8
               Top             =   1287
               Width           =   1800
            End
            Begin VB.TextBox Txt_Cat_Empresas_Telefono 
               Height          =   285
               Left            =   6165
               MaxLength       =   20
               TabIndex        =   6
               Top             =   918
               Width           =   1800
            End
            Begin VB.TextBox Txt_Cat_Empresas_Empresa_ID 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   0
               Top             =   180
               Width           =   1395
            End
            Begin VB.TextBox Txt_Cat_Empresas_Nombre 
               Height          =   285
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   2
               Top             =   549
               Width           =   4000
            End
            Begin VB.TextBox Txt_Cat_Empresas_RFC 
               Height          =   285
               Left            =   6165
               MaxLength       =   20
               TabIndex        =   3
               Top             =   180
               Width           =   1800
            End
            Begin VB.TextBox Txt_Cat_Empresas_Ciudad 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   9
               Top             =   1656
               Width           =   4000
            End
            Begin VB.TextBox Txt_Cat_Empresas_CP 
               Height          =   285
               Left            =   6165
               MaxLength       =   20
               TabIndex        =   4
               Top             =   549
               Width           =   1800
            End
            Begin VB.TextBox Txt_Cat_Empresas_Direccion 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   5
               Top             =   918
               Width           =   4000
            End
            Begin VB.TextBox Txt_Cat_Empresas_Colonia 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   7
               Top             =   1287
               Width           =   4000
            End
            Begin VB.TextBox Txt_Cat_Empresas_Acronimo 
               Height          =   285
               Left            =   3330
               MaxLength       =   20
               TabIndex        =   1
               Top             =   180
               Width           =   1755
            End
            Begin VB.TextBox Txt_Cat_Empresas_Comentarios 
               Height          =   390
               Left            =   1080
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   2340
               Width           =   6900
            End
            Begin VB.ComboBox Cmb_Cat_Empresas_Tipo_Nomina 
               Height          =   315
               ItemData        =   "Frm_Cat_Generales_RH.frx":1CB6
               Left            =   6165
               List            =   "Frm_Cat_Generales_RH.frx":1CC0
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   1641
               Width           =   1800
            End
            Begin VB.TextBox Txt_Cat_Empresas_Noi_Coi_ID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1080
               MaxLength       =   2
               TabIndex        =   11
               Top             =   1980
               Visible         =   0   'False
               Width           =   4020
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000E&
               Caption         =   "Logo"
               Height          =   255
               Left            =   5190
               TabIndex        =   171
               Top             =   2025
               Width           =   855
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Estado"
               Height          =   195
               Left            =   5190
               TabIndex        =   54
               Top             =   1332
               Width           =   495
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono"
               Height          =   195
               Left            =   5190
               TabIndex        =   53
               Top             =   963
               Width           =   630
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Empresa ID                         Acronimo"
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
               Left            =   45
               TabIndex        =   52
               Top             =   225
               Width           =   3270
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   45
               TabIndex        =   51
               Top             =   604
               Width           =   660
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RFC"
               Height          =   195
               Left            =   5190
               TabIndex        =   50
               Top             =   225
               Width           =   315
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CP"
               Height          =   195
               Left            =   5190
               TabIndex        =   49
               Top             =   594
               Width           =   210
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ciudad"
               Height          =   195
               Left            =   45
               TabIndex        =   48
               Top             =   1741
               Width           =   495
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dirección"
               Height          =   195
               Left            =   45
               TabIndex        =   47
               Top             =   983
               Width           =   675
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Colonia"
               Height          =   195
               Left            =   45
               TabIndex        =   46
               Top             =   1362
               Width           =   525
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comentarios"
               Height          =   195
               Left            =   45
               TabIndex        =   45
               Top             =   2445
               Width           =   870
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Tipo Nómina"
               Height          =   195
               Left            =   5190
               TabIndex        =   44
               Top             =   1701
               Width           =   900
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Nomipaq ID"
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
               Left            =   45
               TabIndex        =   43
               Top             =   2025
               Visible         =   0   'False
               Width           =   1005
            End
         End
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESAS"
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
         Left            =   3120
         TabIndex        =   41
         Top             =   15
         Width           =   2145
      End
   End
   Begin VB.PictureBox Pic_Cat_Tipos_Faltas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8400
      TabIndex        =   31
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Tipos_Faltas_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   2130
         Left            =   120
         TabIndex        =   32
         Top             =   495
         Width           =   8250
         Begin VB.ComboBox Cmb_Clasificacion_Incidencias 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales_RH.frx":1CD8
            Left            =   5730
            List            =   "Frm_Cat_Generales_RH.frx":1CE2
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox Txt_Clave_SAP_Tipos_Faltas 
            Height          =   315
            Left            =   5735
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1035
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox Txt_Cat_Tipos_Faltas_Descripcion 
            Height          =   315
            Left            =   1500
            TabIndex        =   26
            Top             =   640
            Width           =   6500
         End
         Begin VB.TextBox Txt_Cat_Tipos_Faltas_Simbologia 
            Height          =   315
            Left            =   1500
            MaxLength       =   5
            TabIndex        =   27
            Top             =   1035
            Width           =   2295
         End
         Begin VB.TextBox Txt_Cat_Tipos_Faltas_Falta_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox Txt_Cat_Tipos_Faltas_Comentarios 
            Height          =   600
            Left            =   1500
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   1440
            Width           =   6500
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4095
            TabIndex        =   135
            Top             =   1095
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Simbología"
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
            Left            =   120
            TabIndex        =   38
            Top             =   1095
            Width           =   960
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Falta ID"
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
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Descripcion"
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
            Left            =   120
            TabIndex        =   33
            Top             =   700
            Width           =   1020
         End
         Begin VB.Label Lbl_Clasificacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clasificacion"
            Height          =   195
            Left            =   4095
            TabIndex        =   138
            Top             =   300
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.Frame Fra_Cat_Tipos_Faltas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipos Faltas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         Left            =   135
         TabIndex        =   36
         Top             =   2655
         Width           =   8250
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Tipos_Faltas 
            Height          =   3240
            Left            =   90
            TabIndex        =   30
            Top             =   240
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   5715
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "INCIDENCIAS EXTRAORDINARIAS"
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
         Left            =   1110
         TabIndex        =   37
         Top             =   0
         Width           =   6225
      End
   End
   Begin VB.PictureBox Pic_Cat_Puestos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8400
      TabIndex        =   120
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Puestos_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1725
         Left            =   120
         TabIndex        =   123
         Top             =   435
         Width           =   8250
         Begin VB.TextBox Txt_Clave_SAP_Puestos 
            Height          =   315
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   127
            Top             =   975
            Visible         =   0   'False
            Width           =   6700
         End
         Begin VB.TextBox Txt_Cat_Puestos_Abreviatura 
            Height          =   315
            Left            =   6150
            MaxLength       =   10
            TabIndex        =   125
            Top             =   240
            Width           =   1900
         End
         Begin VB.TextBox Txt_Cat_Puestos_Puesto_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   124
            Top             =   225
            Width           =   1900
         End
         Begin VB.TextBox Txt_Cat_Puestos_Nombre 
            Height          =   330
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   126
            Top             =   600
            Width           =   6700
         End
         Begin VB.TextBox Txt_Cat_Puestos_Comentarios 
            Height          =   315
            Left            =   1350
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   128
            Top             =   1335
            Width           =   6700
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   150
            TabIndex        =   136
            Top             =   1020
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abreviatura"
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
            Left            =   4980
            TabIndex        =   132
            Top             =   285
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puesto ID"
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
            Left            =   150
            TabIndex        =   131
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   150
            TabIndex        =   130
            Top             =   675
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   150
            TabIndex        =   129
            Top             =   1395
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Cat_Puestos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4110
         Left            =   120
         TabIndex        =   121
         Top             =   2160
         Width           =   8250
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Puestos 
            Height          =   3750
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   6615
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PUESTOS"
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
         Left            =   3345
         TabIndex        =   133
         Top             =   45
         Width           =   1875
      End
   End
   Begin VB.PictureBox Pic_Cat_Nivel_Estudio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8400
      TabIndex        =   109
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Nivel_Estudio 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Niveles de Estudio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   120
         TabIndex        =   117
         Top             =   2520
         Width           =   8205
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Nivel_Estudio 
            Height          =   3330
            Left            =   75
            TabIndex        =   118
            Top             =   240
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   5874
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Cat_Nivel_Estudio_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1905
         Left            =   120
         TabIndex        =   110
         Top             =   585
         Width           =   8205
         Begin VB.TextBox Txt_Cat_Nivel_Estudio_Descripcion 
            Height          =   735
            Left            =   1320
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   113
            Top             =   1065
            Width           =   6735
         End
         Begin VB.TextBox Txt_Cat_Nivel_Estudio_Nombre 
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   112
            Top             =   680
            Width           =   6735
         End
         Begin VB.TextBox Txt_Cat_Nivel_Estudio_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   240
            Width           =   2000
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   1065
            Width           =   840
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   115
            Top             =   740
            Width           =   660
         End
         Begin VB.Label Label92 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel ID"
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
            Left            =   120
            TabIndex        =   114
            Top             =   285
            Width           =   705
         End
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL DE ESTUDIOS"
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
         Left            =   2325
         TabIndex        =   119
         Top             =   0
         Width           =   3825
      End
   End
   Begin VB.PictureBox Pic_Cat_Motivos_Baja 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8400
      TabIndex        =   97
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Motivos_Baja_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1905
         Left            =   120
         TabIndex        =   100
         Top             =   585
         Width           =   8205
         Begin VB.TextBox Txt_Clave_SAP_Motivos_Baja 
            Height          =   315
            Left            =   4095
            MaxLength       =   50
            TabIndex        =   102
            Top             =   270
            Visible         =   0   'False
            Width           =   3960
         End
         Begin VB.TextBox Txt_Cat_Motivos_Baja_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   270
            Width           =   1485
         End
         Begin VB.TextBox Txt_Cat_Motivos_Baja_Nombre 
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   103
            Top             =   680
            Width           =   6735
         End
         Begin VB.TextBox Txt_Cat_Motivos_Baja_Descripcion 
            Height          =   735
            Left            =   1320
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   1065
            Width           =   6735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   2970
            TabIndex        =   137
            Top             =   330
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo ID"
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
            Left            =   120
            TabIndex        =   107
            Top             =   330
            Width           =   840
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   106
            Top             =   740
            Width           =   660
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1065
            Width           =   840
         End
      End
      Begin VB.Frame Fra_Cat_Motivos_Baja 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Motivos de Bajas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   120
         TabIndex        =   98
         Top             =   2520
         Width           =   8205
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Motivos_Baja 
            Height          =   3330
            Left            =   75
            TabIndex        =   99
            Top             =   240
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   5874
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
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOTIVOS DE BAJAS"
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
         Left            =   2370
         TabIndex        =   108
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox Pic_Cat_Equipos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleWidth      =   8640
      TabIndex        =   82
      Top             =   0
      Width           =   8640
      Begin VB.Frame Fra_Cat_Equipos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Equipos de Identificacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         Left            =   120
         TabIndex        =   83
         Top             =   2280
         Width           =   8250
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Equipos 
            Height          =   3660
            Left            =   75
            TabIndex        =   84
            Top             =   240
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   6456
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Cat_Equipos_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   85
         Top             =   480
         Width           =   8295
         Begin VB.CommandButton Btn_Configuracion_Equipo 
            Caption         =   "Configuracion Equipo"
            Height          =   315
            Left            =   6120
            TabIndex        =   174
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cat_Equipos_Descripcion 
            Height          =   315
            Left            =   1305
            MaxLength       =   200
            ScrollBars      =   2  'Vertical
            TabIndex        =   90
            Top             =   1080
            Width           =   6825
         End
         Begin VB.TextBox Txt_Cat_Equipos_No_Equipo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6120
            TabIndex        =   87
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cat_Equipos_Puerto_IP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6105
            TabIndex        =   89
            Text            =   "4370"
            Top             =   660
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cat_Equipos_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   240
            Width           =   2000
         End
         Begin MSMask.MaskEdBox Txt_Cat_Equipos_Direccion_IP 
            Height          =   315
            Left            =   1305
            TabIndex        =   88
            Top             =   660
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            Format          =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicacion"
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
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto IP"
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
            Left            =   4800
            TabIndex        =   94
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion IP"
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
            Left            =   120
            TabIndex        =   93
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo ID"
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
            Left            =   120
            TabIndex        =   92
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Equipo"
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
            Left            =   4800
            TabIndex        =   91
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Label Label89 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EQUIPO DE IDENTIFICACION"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   96
         Top             =   0
         Width           =   5355
      End
   End
   Begin VB.PictureBox Pic_Cat_Turnos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8595
      TabIndex        =   139
      Top             =   0
      Width           =   8595
      Begin VB.Frame Fra_Cat_Turnos_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1425
         Left            =   105
         TabIndex        =   162
         Top             =   360
         Width           =   8250
         Begin VB.TextBox Txt_Cat_Turnos_Turno_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   195
            Width           =   1065
         End
         Begin VB.TextBox Txt_Cat_Turnos_Nombre 
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   164
            Top             =   600
            Width           =   6645
         End
         Begin VB.TextBox Txt_Cat_Turnos_Comentarios 
            Height          =   330
            Left            =   1500
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   163
            Top             =   975
            Width           =   6645
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Turno ID"
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
            Left            =   135
            TabIndex        =   168
            Top             =   255
            Width           =   765
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   1
            Left            =   135
            TabIndex        =   167
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Lbl_Turnos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   166
            Top             =   1050
            Width           =   870
         End
      End
      Begin TabDlg.SSTab Tab_Turnos 
         Height          =   4365
         Left            =   75
         TabIndex        =   140
         Top             =   1890
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   7699
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Turnos"
         TabPicture(0)   =   "Frm_Cat_Generales_RH.frx":1CFC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Cat_Turnos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Configuración"
         TabPicture(1)   =   "Frm_Cat_Generales_RH.frx":1D18
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Turnos_Detalles"
         Tab(1).ControlCount=   1
         Begin VB.Frame Fra_Cat_Turnos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Turnos"
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
            Left            =   60
            TabIndex        =   160
            Top             =   405
            Width           =   8190
            Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Turnos 
               Height          =   3570
               Left            =   75
               TabIndex        =   161
               Top             =   225
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   6297
               _Version        =   393216
               Rows            =   0
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Fra_Turnos_Detalles 
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
            Height          =   3945
            Left            =   -74925
            TabIndex        =   141
            Top             =   330
            Width           =   8190
            Begin VB.CommandButton Btn_Eliminar_Dia 
               Caption         =   "Eliminar"
               Height          =   285
               Left            =   6960
               TabIndex        =   147
               Top             =   1215
               Width           =   1140
            End
            Begin VB.CommandButton Btn_Agregar 
               Caption         =   "Agregar"
               Height          =   285
               Left            =   5790
               TabIndex        =   146
               Top             =   1215
               Width           =   1140
            End
            Begin VB.CheckBox Chk_Dia_Descanso 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "Dia de Descanso"
               Height          =   285
               Left            =   5955
               TabIndex        =   145
               Top             =   225
               Width           =   1875
            End
            Begin VB.ComboBox Cmb_Dias_Semana 
               Height          =   315
               ItemData        =   "Frm_Cat_Generales_RH.frx":1D34
               Left            =   1455
               List            =   "Frm_Cat_Generales_RH.frx":1D4D
               Style           =   2  'Dropdown List
               TabIndex        =   144
               Top             =   210
               Width           =   4470
            End
            Begin VB.TextBox Txt_Horas_Turno 
               Height          =   315
               Left            =   7155
               MaxLength       =   50
               TabIndex        =   143
               Top             =   540
               Width           =   945
            End
            Begin VB.TextBox Txt_Horas_Comida 
               Height          =   315
               Left            =   7155
               MaxLength       =   50
               TabIndex        =   142
               Top             =   870
               Width           =   945
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Turnos_Hora_Inicio 
               Height          =   315
               Left            =   1455
               TabIndex        =   148
               Top             =   540
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Turnos_Hora_Termino 
               Height          =   315
               Left            =   4470
               TabIndex        =   149
               Top             =   540
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Turnos_Comida_Inicio 
               Height          =   315
               Left            =   1455
               TabIndex        =   150
               Top             =   870
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSComCtl2.DTPicker Dtp_Cat_Turnos_Comida_Termino 
               Height          =   315
               Left            =   4470
               TabIndex        =   151
               Top             =   870
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "HH:mm:ss"
               Format          =   123666435
               UpDown          =   -1  'True
               CurrentDate     =   39863
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Detalles_Turnos 
               Height          =   2310
               Left            =   90
               TabIndex        =   152
               Top             =   1530
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   4075
               _Version        =   393216
               Rows            =   0
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Termino Comida"
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
               Index           =   5
               Left            =   3000
               TabIndex        =   159
               Top             =   930
               Width           =   1365
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Termino"
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
               Left            =   3000
               TabIndex        =   158
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label Lbl_Dia_Semana 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dia Semana"
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
               Index           =   7
               Left            =   135
               TabIndex        =   157
               Top             =   270
               Width           =   1035
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Inicio"
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
               Left            =   135
               TabIndex        =   156
               Top             =   600
               Width           =   945
            End
            Begin VB.Label Lbl_Turnos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio Comida"
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
               Index           =   4
               Left            =   135
               TabIndex        =   155
               Top             =   930
               Width           =   1155
            End
            Begin VB.Label Lbl_Horas_Turno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Horas Turno"
               Height          =   195
               Left            =   5955
               TabIndex        =   154
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Lbl_Horas_Comida 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Horas Comida"
               Height          =   195
               Left            =   5955
               TabIndex        =   153
               Top             =   930
               Width           =   990
            End
         End
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TURNOS"
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
         Left            =   3405
         TabIndex        =   169
         Top             =   0
         Width           =   1665
      End
   End
   Begin VB.PictureBox Pic_Cat_Dias_No_Laborales 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8520
      TabIndex        =   71
      Top             =   0
      Width           =   8520
      Begin VB.Frame Fra_Cat_Dias_No_Laborales_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1950
         Left            =   135
         TabIndex        =   74
         Top             =   585
         Width           =   8205
         Begin VB.TextBox Txt_Cat_Dias_No_Laborales_Comentarios 
            Height          =   735
            Left            =   1725
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   1125
            Width           =   6375
         End
         Begin VB.TextBox Txt_Cat_Dias_No_Laborales_Dia_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1725
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   240
            Width           =   2000
         End
         Begin MSComCtl2.DTPicker Dtp_Cat_Dias_No_Laborales_Fecha 
            Height          =   315
            Left            =   1725
            TabIndex        =   77
            Top             =   682
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   123666435
            CurrentDate     =   39863
            MinDate         =   32874
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1125
            Width           =   870
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia No Laboral ID"
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
            Left            =   120
            TabIndex        =   79
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia"
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
            Left            =   120
            TabIndex        =   78
            Top             =   742
            Width           =   300
         End
      End
      Begin VB.Frame Fra_Cat_Dias_No_Laborales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Días No Laborales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   120
         TabIndex        =   72
         Top             =   2565
         Width           =   8205
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Dias_No_Laborales 
            Height          =   3375
            Left            =   75
            TabIndex        =   73
            Top             =   240
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIAS NO LABORALES"
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
         Left            =   2265
         TabIndex        =   81
         Top             =   0
         Width           =   3945
      End
   End
   Begin VB.PictureBox Pic_Cat_Equipos_Almacenes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8400
      TabIndex        =   175
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cat_Equipos_Almacenes_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   178
         Top             =   480
         Width           =   8295
         Begin VB.CommandButton Btn_Configuracion_Equipo_Almacenes 
            Caption         =   "Configuracion Equipo"
            Height          =   315
            Left            =   6120
            TabIndex        =   179
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cat_Equipos_Almacenes_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   182
            Top             =   240
            Width           =   2000
         End
         Begin VB.TextBox Txt_Cat_Equipos_Almacenes_Puerto_IP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6105
            TabIndex        =   181
            Text            =   "4370"
            Top             =   660
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cat_Equipos_Almacenes_No_Equipo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6120
            TabIndex        =   180
            Top             =   240
            Width           =   2055
         End
         Begin MSMask.MaskEdBox Txt_Cat_Equipos_Almacenes_Direccion_IP 
            Height          =   315
            Left            =   1305
            TabIndex        =   183
            Top             =   660
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            Format          =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txt_Cat_Equipos_Almacenes_Descripcion 
            Height          =   315
            Left            =   1320
            MaxLength       =   200
            ScrollBars      =   2  'Vertical
            TabIndex        =   184
            Top             =   1080
            Width           =   6825
         End
         Begin VB.Label Lbl_No_Equipo_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Equipo"
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
            Left            =   4800
            TabIndex        =   189
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Lbl_Equipo_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo ID"
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
            Left            =   120
            TabIndex        =   188
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Lbl_Direccion_IP_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion IP"
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
            Left            =   120
            TabIndex        =   187
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Lbl_Puerto_IP_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto IP"
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
            Left            =   4800
            TabIndex        =   186
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Lbl_Ubicacion_Almacenes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicacion"
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
            Left            =   120
            TabIndex        =   185
            Top             =   1080
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Cat_Equipos_Almacenes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Equipos de Identificacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         Left            =   120
         TabIndex        =   176
         Top             =   2265
         Width           =   8250
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Equipos_Almacenes 
            Height          =   3660
            Left            =   75
            TabIndex        =   177
            Top             =   240
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   6456
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label89 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EQUIPO DE IDENTIFICACION ALMACENES"
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
         Index           =   1
         Left            =   465
         TabIndex        =   190
         Top             =   0
         Width           =   7785
      End
   End
   Begin VB.PictureBox Pic_Cat_Departamentos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   8505
      TabIndex        =   57
      Top             =   0
      Width           =   8505
      Begin VB.Frame Fra_Departamentos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Departamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   120
         TabIndex        =   68
         Top             =   2640
         Width           =   8250
         Begin MSFlexGridLib.MSFlexGrid Grid_Departamentos 
            Height          =   3210
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   5662
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Departamento_Generales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales"
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
         Height          =   1935
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   8250
         Begin VB.TextBox Txt_Clave_SAP_Departamentos 
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   61
            Top             =   990
            Visible         =   0   'False
            Width           =   6375
         End
         Begin VB.TextBox Txt_Departamento_Comentarios 
            Height          =   450
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   1380
            Width           =   6375
         End
         Begin VB.TextBox Txt_Departamento_Nombre 
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   60
            Top             =   615
            Width           =   6375
         End
         Begin VB.TextBox Txt_Departamento_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   232
            Width           =   1500
         End
         Begin VB.TextBox Txt_Departamento_Clave 
            Height          =   315
            Left            =   5670
            MaxLength       =   50
            TabIndex        =   59
            Top             =   225
            Width           =   2385
         End
         Begin VB.Label Lbl_Clave_SAP_Departamentos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   134
            Top             =   1035
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   66
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento ID"
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
            Left            =   120
            TabIndex        =   65
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clave"
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
            Left            =   4905
            TabIndex        =   64
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTAMENTOS"
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
         Left            =   2475
         TabIndex        =   70
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Image Logo_Temp 
      Height          =   375
      Left            =   7680
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Nomipaq ID"
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
      Left            =   0
      TabIndex        =   170
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "Frm_Cat_Generales_RH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios
Public Catalogo As String          'Indicar que formulario se va a abrir
Dim Horas_Iguales_Confirmadas As Boolean
Private Temp_MouseRow As Long
Private Temp_MouseCol As Long
Private Copia_Configuracion_Calendario As Boolean
Private Ejecutando_MouseUp As Boolean
Private Ejecutando_LostFocus As Boolean
Private Ejecutando_Grid_Calendarios_Configuracion_Turnos_DblClick As Boolean
Private Temp_Col As Integer
Private Temp_Row As Integer
Private Temp_Enter_Cell As Boolean

Private Sub Btn_Agregar_Click()
Dim Descanso As String
Dim Cont_Fila As Integer

On Error GoTo HANDLER
    If Cmb_Dias_Semana.ListIndex > -1 Then
        ''se valida si no esta ya agregado el dia
        For Cont_Fila = 0 To Grid_Detalles_Turnos.Rows - 1
            If Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 1) = Cmb_Dias_Semana.Text Then
                MsgBox "El dia ya esta agregado en la lista", vbInformation
                Exit Sub
            End If
        Next
        ''Se da de alta el registro
        If Grid_Detalles_Turnos.Rows < 1 Then
            ''Se agrega encabezado
             Grid_Detalles_Turnos.Cols = 9
             Grid_Detalles_Turnos.AddItem _
             "Turno ID" _
             & Chr(9) & "Dia" _
             & Chr(9) & "H.Inicio" _
             & Chr(9) & "H.Termino" _
             & Chr(9) & "C.Inicio" _
             & Chr(9) & "C.Termino" _
             & Chr(9) & "H.Turnos" _
             & Chr(9) & "H.Comida" _
             & Chr(9) & "Descanso"
        End If
        If Chk_Dia_Descanso.Value = 1 Then
            Descanso = "SI"
        Else
            Descanso = "NO"
        End If
        'Horas del Turno
        If ((Val(DateDiff("n", Dtp_Cat_Turnos_Hora_Inicio.Value, Dtp_Cat_Turnos_Comida_Inicio.Value)) + Val(DateDiff("n", Dtp_Cat_Turnos_Comida_Termino.Value, Dtp_Cat_Turnos_Hora_Termino.Value))) / 60) > 0 Then
            Txt_Horas_Turno.Text = (Val(DateDiff("n", Dtp_Cat_Turnos_Hora_Inicio.Value, Dtp_Cat_Turnos_Comida_Inicio.Value)) + Val(DateDiff("n", Dtp_Cat_Turnos_Comida_Termino.Value, Dtp_Cat_Turnos_Hora_Termino.Value))) / 60
        Else
            Txt_Horas_Turno.Text = 24 + ((Val(DateDiff("n", Dtp_Cat_Turnos_Hora_Inicio.Value, Dtp_Cat_Turnos_Comida_Inicio.Value)) + Val(DateDiff("n", Dtp_Cat_Turnos_Comida_Termino.Value, Dtp_Cat_Turnos_Hora_Termino.Value))) / 60)
        End If
        'Horas de comida
        Txt_Horas_Comida.Text = Val(DateDiff("n", Dtp_Cat_Turnos_Comida_Inicio.Value, Dtp_Cat_Turnos_Comida_Termino.Value)) / 60
        'Se agrega el registro
        Grid_Detalles_Turnos.AddItem Txt_Cat_Turnos_Turno_ID.Text _
            & Chr(9) & Cmb_Dias_Semana.Text _
            & Chr(9) & Format(Dtp_Cat_Turnos_Hora_Inicio.Value, "HH:mm") _
            & Chr(9) & Format(Dtp_Cat_Turnos_Hora_Termino.Value, "HH:mm") _
            & Chr(9) & Format(Dtp_Cat_Turnos_Comida_Inicio.Value, "HH:mm") _
            & Chr(9) & Format(Dtp_Cat_Turnos_Comida_Termino.Value, "HH:mm") _
            & Chr(9) & Txt_Horas_Turno.Text _
            & Chr(9) & Txt_Horas_Comida.Text _
            & Chr(9) & Descanso
        ''Se formatea el grid
        With Grid_Detalles_Turnos
            .FixedRows = 1
            .FixedCols = 2
            .ColWidth(0) = 0  'Dia
            .ColWidth(1) = 800  'H.Inicio
            .ColAlignment(1) = flexAlignCenterTop
            .ColWidth(2) = 1000   'H.Termino
            .ColAlignment(2) = flexAlignCenterTop
            .ColWidth(3) = 1000  'C.Inicio
            .ColAlignment(3) = flexAlignCenterTop
            .ColWidth(4) = 1000     'C.Termino
            .ColAlignment(4) = flexAlignCenterTop
            .ColWidth(5) = 1000     'H.Turnos
            .ColAlignment(5) = flexAlignCenterTop
            .ColWidth(6) = 1000     'H.Comida
            .ColAlignment(6) = flexAlignCenterTop
            .ColWidth(7) = 1000     'Descanso
            .ColAlignment(7) = flexAlignCenterTop
            .ColWidth(8) = 870     'Descanso
            .ColAlignment(8) = flexAlignCenterTop
        End With
    Else
        MsgBox "Seleccione un dia", vbInformation
        Cmb_Dias_Semana.SetFocus
    End If
    Exit Sub
HANDLER:
    MsgBox Err.Description
End Sub

Private Sub Btn_Cat_Empresas_Agregar_Equipo_Almacenes_Click()
Dim Cont_Fila As Integer
    'Valida los datos del dependiente
    If Cmb_Cat_Empresas_Equipo_Almacenes.ListIndex > -1 Then
        'Agrega el dependiente a la lista
        'Busca si el equipo ya ha sido agregado
        For Cont_Fila = 1 To Grid_Empresas_Equipos_Almacenes.Rows - 1
            If Format(Cmb_Cat_Empresas_Equipo_Almacenes.ItemData(Cmb_Cat_Empresas_Equipo_Almacenes.ListIndex), "00000") = _
              Trim(Grid_Empresas_Equipos_Almacenes.TextMatrix(Cont_Fila, 0)) Then
                MsgBox "El equipo ya se agregó", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
        Next
        Grid_Empresas_Equipos_Almacenes.Cols = 2
        If Grid_Empresas_Equipos_Almacenes.Rows = 0 Then
            Grid_Empresas_Equipos_Almacenes.AddItem "Equipo_ID" & Chr(9) & "Equipo"
            Grid_Empresas_Equipos_Almacenes.ColWidth(0) = 0    'Equipo_ID
            Grid_Empresas_Equipos_Almacenes.ColWidth(1) = 3500 'Equipo
            Grid_Empresas_Equipos_Almacenes.ColAlignment(1) = flexAlignLeftCenter
        End If
        Grid_Empresas_Equipos_Almacenes.AddItem Format(Cmb_Cat_Empresas_Equipo_Almacenes.ItemData(Cmb_Cat_Empresas_Equipo_Almacenes.ListIndex), "00000") & Chr(9) & _
            Trim(Cmb_Cat_Empresas_Equipo_Almacenes.Text)
        Cmb_Cat_Empresas_Equipo_Almacenes.ListIndex = -1
        Grid_Empresas_Equipos_Almacenes.FixedRows = 1
    End If
End Sub

Private Sub Btn_Cat_Empresas_Agregar_Equipo_Click()
Dim Cont_Fila As Integer
    'Valida los datos del dependiente
    If Cmb_Cat_Empresas_Equipo.ListIndex > -1 Then
        'Agrega el dependiente a la lista
        'Busca si el equipo ya ha sido agregado
        For Cont_Fila = 1 To Grid_Empresas_Equipos.Rows - 1
            If Format(Cmb_Cat_Empresas_Equipo.ItemData(Cmb_Cat_Empresas_Equipo.ListIndex), "00000") = _
              Trim(Grid_Empresas_Equipos.TextMatrix(Cont_Fila, 0)) Then
                MsgBox "El equipo ya se agregó", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
        Next
        Grid_Empresas_Equipos.Cols = 2
        If Grid_Empresas_Equipos.Rows = 0 Then
            Grid_Empresas_Equipos.AddItem "Equipo_ID" & Chr(9) & "Equipo"
            Grid_Empresas_Equipos.ColWidth(0) = 0    'Equipo_ID
            Grid_Empresas_Equipos.ColWidth(1) = 3500 'Equipo
            Grid_Empresas_Equipos.ColAlignment(1) = flexAlignLeftCenter
        End If
        Grid_Empresas_Equipos.AddItem Format(Cmb_Cat_Empresas_Equipo.ItemData(Cmb_Cat_Empresas_Equipo.ListIndex), "00000") & Chr(9) & _
            Trim(Cmb_Cat_Empresas_Equipo.Text)
        Cmb_Cat_Empresas_Equipo.ListIndex = -1
        Grid_Empresas_Equipos.FixedRows = 1
    End If
End Sub

Private Sub Btn_Cat_Empresas_Eliminar_Equipo_Almacenes_Click()
    If Grid_Empresas_Equipos_Almacenes.Rows > 0 Then
        If Grid_Empresas_Equipos_Almacenes.Rows = 2 Then
            Grid_Empresas_Equipos_Almacenes.Rows = 0
        Else
            Grid_Empresas_Equipos_Almacenes.RemoveItem Grid_Empresas_Equipos_Almacenes.RowSel
        End If
    End If

End Sub

Private Sub Btn_Cat_Empresas_Eliminar_Equipo_Click()
    If Grid_Empresas_Equipos.Rows > 0 Then
        If Grid_Empresas_Equipos.Rows = 2 Then
            Grid_Empresas_Equipos.Rows = 0
        Else
            Grid_Empresas_Equipos.RemoveItem Grid_Empresas_Equipos.RowSel
        End If
    End If
End Sub

Private Sub Btn_Configuracion_Calendarios_Turnos_Limpiar_Calendario_Click()
Dim Cont_Filas As Integer
Dim Cont_Columnas As Integer
'    Grid_Calendarios_Configuracion_Turnos.Cols = 0
'    Grid_Calendarios_Configuracion_Turnos.Rows = 0
    Grid_Calendarios_Configuracion_Turnos.Redraw = False
    For Cont_Filas = Grid_Calendarios_Configuracion_Turnos.FixedRows To Grid_Calendarios_Configuracion_Turnos.Rows - 1
        For Cont_Columnas = Grid_Calendarios_Configuracion_Turnos.FixedCols To Grid_Calendarios_Configuracion_Turnos.Cols - 1
            Grid_Calendarios_Configuracion_Turnos.Row = Cont_Filas
            Grid_Calendarios_Configuracion_Turnos.Col = Cont_Columnas
            If Grid_Calendarios_Configuracion_Turnos.CellBackColor <> 0 Then
                Grid_Calendarios_Configuracion_Turnos.CellBackColor = 0
            End If
            Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas) = ""
        Next Cont_Columnas
    Next Cont_Filas
    Grid_Calendarios_Configuracion_Turnos.Redraw = True
End Sub

Private Sub Btn_Configuracion_Calendarios_Turnos_Limpiar_Datos_Click()
    Txt_Calendarios_Configuracion_Turno.Text = ""
    Dtp_Calendario_Fecha_Inicio.Value = DateValue(Now)
    Dtp_Calendario_Fecha_Termino.Value = DateValue(Now)
    Dtp_Calendario_Hora_Inicio.Value = TimeSerial(0, 0, 0)
    Dtp_Calendario_Hora_Termino.Value = TimeSerial(0, 0, 0)
    Dtp_Calendario_Inicio_Comida.Value = TimeSerial(0, 0, 0)
    Dtp_Calendario_Termino_Comida.Value = TimeSerial(0, 0, 0)
    Txt_Calendario_Horas_Turno.Text = ""
    Txt_Calendario_Horas_Comida.Text = ""
    Txt_Calendario_Filtro_Empleados.Text = ""
    Call Lst_Calendarios_Configuracion_Empleados.Clear
End Sub

Private Sub Btn_Configuracion_Equipo_Almacenes_Click()
Unload Frm_Apl_Configuracion_Checador
    Load Frm_Apl_Configuracion_Checador
    If Trim(Txt_Cat_Equipos_Almacenes_ID.Text) <> "" Then
        Frm_Apl_Configuracion_Checador.Equipo_ID = Trim(Txt_Cat_Equipos_Almacenes_ID.Text)
    End If
    Frm_Apl_Configuracion_Checador.Inicializa
End Sub

Private Sub Btn_Configuracion_Equipo_Click()
    Unload Frm_Apl_Configuracion_Checador
    Load Frm_Apl_Configuracion_Checador
    If Trim(Txt_Cat_Equipos_ID.Text) <> "" Then
        Frm_Apl_Configuracion_Checador.Equipo_ID = Trim(Txt_Cat_Equipos_ID.Text)
    End If
    Frm_Apl_Configuracion_Checador.Inicializa
End Sub

Private Sub Btn_Consultar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
    If Catalogo = "Cat_Empleados" Then
        Nombre = InputBox("Proporcione el No. Nómina, Nombre, Apellido, RFC, NSS para buscar Empleados")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
    Else
        If Catalogo = "Cat_Dias_No_Laborales" Then
            Nombre = InputBox("Proporcione la descripcion del dia", Me.Caption)
        Else
            Nombre = InputBox("Proporcione el nombre", Me.Caption)
        End If
    End If
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Select Case Catalogo
        Case "Cat_Empresas"
            Consulta_Cat_Empresas Nombre
        Case "Cat_Turnos":
            Consulta_Cat_Turnos Nombre
        Case "Cat_Calendarios_Turnos":
            Consulta_Cat_Calendarios_Turnos Nombre
        Case "Cat_Dias_No_Laborales":
            Consulta_Cat_Dias_No_Laborales Nombre
        Case "Cat_Tipos_Faltas"
            Consulta_Cat_Tipos_Faltas Nombre
        Case "Cat_Departamentos": 'Catalogo de Departamentos
            Consulta_Departamentos Nombre
        Case "Cat_Equipos_Identificacion":
            Consulta_Cat_Equipos_Identificadores Nombre
        Case "Cat_Puestos"
            Consulta_Cat_Puestos Nombre
        Case "Cat_Nivel_Estudio"
            Consulta_Cat_Nivel_Estudio Nombre
        Case "Cat_Motivos_Baja"
            Consulta_Cat_Motivos_Baja Nombre
        Case "Cat_Equipos_Almacenes_Identificacion": 'Catalogo de Equipos Almacenes
            Consulta_Cat_Equipos_Almacenes_Identificadores Nombre
    End Select
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo HANDLER
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Conexion_Base.BeginTrans
            Select Case Catalogo
                Case "Cat_Empresas":
                    If Trim(Txt_Cat_Empresas_Empresa_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Empresas", "Empresa_ID", Trim(Txt_Cat_Empresas_Empresa_ID.Text)) = True Then
                            If Grid_Cat_Empresas.Rows = 2 Then
                                Grid_Cat_Empresas.Rows = 0
                            Else
                                Grid_Cat_Empresas.RemoveItem Grid_Cat_Empresas.RowSel
                            End If
                            Call Conectar_Ayudante.Limpiar_Textos(Me)
                            MsgBox "Empresa eliminada", vbInformation + vbOKOnly, Me.Caption
                        Else
                            MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione una empresa para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
                
                Case "Cat_Turnos": 'Catalogo de Turnos
                    If Trim(Txt_Cat_Turnos_Turno_ID.Text) <> "" Then
                        Dim Rs_Modificacion_Cat_Turnos As rdoResultset
                        Mi_SQL = "SELECT * FROM Cat_Turnos WHERE Turno_ID = '" & Trim(Txt_Cat_Turnos_Turno_ID.Text) & "'"
                        Set Rs_Modificacion_Cat_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        'Modifica los datos de la tabla Cat_Turnos
                        With Rs_Modificacion_Cat_Turnos
                            .Edit
                                .rdoColumns("Estatus") = "INACTIVO"
                            .Update
                        .Close
                        End With
                        Set Rs_Modificacion_Cat_Turnos = Nothing
                        MsgBox "Turno Inactivado", vbInformation + vbOKOnly, Me.Caption
                        Consulta_Cat_Turnos ""
'                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Turnos", "Turno_ID", Trim(Txt_Cat_Turnos_Turno_ID.Text)) = True Then
'                            'Quita los datos del usuario contenidos en el Grid
'                            If Grid_Cat_Turnos.Rows = 2 Then
'                                Grid_Cat_Turnos.Rows = 0
'                            Else
'                                Grid_Cat_Turnos.RemoveItem Grid_Cat_Turnos.RowSel
'                            End If 'Grid_productos
'                            MsgBox "Turno Eliminado", vbInformation + vbOKOnly, Me.Caption
'                        Else
'                            MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
'                        End If
                    Else
                        MsgBox "Seleccione un turno para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
                
                Case "Cat_Calendarios_Turnos": 'Catalogo de Calendarios de Turnos
                    If Trim(Txt_Calendario_Turno_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Calendarios_Turnos_Detalles", "Calendario_Turno_ID", Trim(Txt_Calendario_Turno_ID.Text)) = True Then
                            Grid_Calendarios_Configuracion_Turnos.Rows = 0
                            Grid_Calendarios_Configuracion_Turnos.Cols = 0
                            If Conectar_Ayudante.Elimina_Catalogo("Cat_Calendarios_Turnos", "Calendario_Turno_ID", Trim(Txt_Calendario_Turno_ID.Text)) = True Then
                                'Quita los datos del usuario contenidos en el Grid
                                If Grid_Calendarios_Turnos.Rows = 2 Then
                                    Grid_Calendarios_Turnos.Rows = 0
                                Else
                                    Grid_Calendarios_Turnos.RemoveItem Grid_Calendarios_Turnos.RowSel
                                End If 'Grid_productos
                                MsgBox "Calendario Eliminado", vbInformation + vbOKOnly, Me.Caption
                            Else
'                                MsgBox "No se pudo eliminar el Calendario", vbExclamation + vbOKOnly, Me.Caption
                                
                                Dim Rs_Modificacion_Cat_Calendarios_Turnos_Detalles As rdoResultset
                                Mi_SQL = "SELECT * FROM Cat_Calendarios_Turnos_Detalles WHERE Calendario_Turno_ID = '" & Trim(Txt_Calendario_Turno_ID.Text) & "'"
                                Set Rs_Modificacion_Cat_Calendarios_Turnos_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                                'Modifica los datos de la tabla Cat_Turnos
                                With Rs_Modificacion_Cat_Calendarios_Turnos_Detalles
                                    .Edit
                                        .rdoColumns("Estatus") = "INACTIVO"
                                    .Update
                                .Close
                                End With
                                Set Rs_Modificacion_Cat_Calendarios_Turnos_Detalles = Nothing
                                    
                                Dim Rs_Modificacion_Cat_Calendarios_Turnos As rdoResultset
                                Mi_SQL = "SELECT * FROM Cat_Calendarios_Turnos WHERE Calendario_Turno_ID = '" & Trim(Txt_Calendario_Turno_ID.Text) & "'"
                                Set Rs_Modificacion_Cat_Calendarios_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                                'Modifica los datos de la tabla Cat_Turnos
                                With Rs_Modificacion_Cat_Calendarios_Turnos
                                    .Edit
                                        .rdoColumns("Estatus") = "INACTIVO"
                                    .Update
                                .Close
                                End With
                                Set Rs_Modificacion_Cat_Calendarios_Turnos = Nothing
                                MsgBox "Turno Inactivado", vbInformation + vbOKOnly, Me.Caption
                            End If
                        Else
                            MsgBox "No se pudo eliminar los Turnos del Calendario", vbExclamation + vbOKOnly, Me.Caption
'                            Dim Rs_Modificacion_Cat_Calendarios_Turnos_Detalles As rdoResultset
'                            Mi_SQL = "SELECT * FROM Cat_Calendarios_Turnos_Detalles WHERE Calendario_Turno_ID = '" & Trim(Txt_Calendario_Turno_ID.Text) & "'"
'                            Set Rs_Modificacion_Cat_Calendarios_Turnos_Detalles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
'                            'Modifica los datos de la tabla Cat_Turnos
'                            With Rs_Modificacion_Cat_Calendarios_Turnos_Detalles
'                                .Edit
'                                    .rdoColumns("Estatus") = "INACTIVO"
'                                .Update
'                            .Close
'                            End With
'                            Set Rs_Modificacion_Cat_Calendarios_Turnos_Detalles = Nothing
                        End If
                    Else
                        MsgBox "Seleccione un Calendario para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
                
                Case "Cat_Dias_No_Laborales": 'Catalogo de dias
                    If Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Dias_No_Laborales", "Dia_No_Laboral_ID", Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text)) = True Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Cat_Dias_No_Laborales.Rows = 2 Then
                                Grid_Cat_Dias_No_Laborales.Rows = 0
                            Else
                                Grid_Cat_Dias_No_Laborales.RemoveItem Grid_Cat_Dias_No_Laborales.RowSel
                            End If 'Grid_productos
                            MsgBox "Dia Eliminado", vbInformation + vbOKOnly, Me.Caption
                        Else
                            MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione un dia para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
                
                Case "Cat_Tipos_Faltas": 'Catalogo de dias
                    If Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Tipos_Faltas", "Tipo_Falta_ID", Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text)) = True Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Cat_Tipos_Faltas.Rows = 2 Then
                                Grid_Cat_Tipos_Faltas.Rows = 0
                            Else
                                Grid_Cat_Tipos_Faltas.RemoveItem Grid_Cat_Tipos_Faltas.RowSel
                            End If 'Grid_productos
                            MsgBox "Tipo de Falta Eliminado", vbInformation + vbOKOnly, Me.Caption
                        Else
                            MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione un tipo de falta para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                    End If
                    
                Case "Cat_Departamentos": 'Catalogo de Departamentos
                    If Trim(Txt_Departamento_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Departamentos", "Departamento_ID", Trim(Txt_Departamento_ID.Text)) Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Departamentos.Rows = 2 Then
                                Grid_Departamentos.Rows = 0
                            Else
                                Grid_Departamentos.RemoveItem Grid_Departamentos.RowSel
                            End If
                            MsgBox "Departamento Eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
                    End If
                
                Case "Cat_Equipos_Identificacion":
                    If Trim(Txt_Cat_Equipos_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Equipos_Identificadores", "Equipo_ID", Trim(Txt_Cat_Equipos_ID.Text)) Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Cat_Equipos.Rows = 2 Then
                                Grid_Cat_Equipos.Rows = 0
                            Else
                                Grid_Cat_Equipos.RemoveItem Grid_Cat_Equipos.RowSel
                            End If
                            MsgBox "Equipo Eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
                    End If
                
                Case "Cat_Puestos"
                    If Trim(Txt_Cat_Puestos_Puesto_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Puestos", "Puesto_ID", Trim(Txt_Cat_Puestos_Puesto_ID.Text)) = True Then
                            If Grid_Cat_Puestos.Rows = 2 Then
                                Grid_Cat_Puestos.Rows = 0
                            Else
                                Grid_Cat_Puestos.RemoveItem Grid_Cat_Puestos.RowSel
                            End If
                            Call Conectar_Ayudante.Limpiar_Textos(Me)
                            MsgBox "El registro ha sido eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione un Puesto para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                
                Case "Cat_Nivel_Estudio"
                    If Trim(Txt_Cat_Nivel_Estudio_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Nivel_Estudio", "Nivel_Estudio_ID", Trim(Txt_Cat_Nivel_Estudio_ID.Text)) = True Then
                            If Grid_Cat_Nivel_Estudio.Rows = 2 Then
                                Grid_Cat_Nivel_Estudio.Rows = 0
                            Else
                                Grid_Cat_Nivel_Estudio.RemoveItem Grid_Cat_Nivel_Estudio.RowSel
                            End If
                            Call Conectar_Ayudante.Limpiar_Textos(Me)
                            MsgBox "El registro ha sido eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione un Puesto para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                
                Case "Cat_Motivos_Baja"
                    If Trim(Txt_Cat_Motivos_Baja_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Motivos_Baja", "Motivo_Baja_ID", Trim(Txt_Cat_Motivos_Baja_ID.Text)) = True Then
                            If Grid_Cat_Motivos_Baja.Rows = 2 Then
                                Grid_Cat_Motivos_Baja.Rows = 0
                            Else
                                Grid_Cat_Motivos_Baja.RemoveItem Grid_Cat_Motivos_Baja.RowSel
                            End If
                            Call Conectar_Ayudante.Limpiar_Textos(Me)
                            MsgBox "El registro ha sido eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
                    Else
                        MsgBox "Seleccione un Motivo de Baja para poder eliminar", vbInformation + vbOKOnly, Me.Caption
                        Exit Sub
                    End If
                'Catalogo de Equipos Almacen
                Case "Cat_Equipos_Almacenes_Identificacion":
                    If Trim(Txt_Cat_Equipos_Almacenes_ID.Text) <> "" Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Equipos_Almacenes_Identificadores", "Equipo_ID", Trim(Txt_Cat_Equipos_Almacenes_ID.Text)) Then
                            'Quita los datos del usuario contenidos en el Grid
                            If Grid_Cat_Equipos_Almacenes.Rows = 2 Then
                                Grid_Cat_Equipos_Almacenes.Rows = 0
                            Else
                                Grid_Cat_Equipos_Almacenes.RemoveItem Grid_Cat_Equipos_Almacenes.RowSel
                            End If
                            MsgBox "Equipo Eliminado", vbInformation + vbOKOnly, Me.Caption
                        End If
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

Private Sub Btn_Eliminar_Dia_Click()
     If Grid_Detalles_Turnos.Rows > 2 Then
         Grid_Detalles_Turnos.RemoveItem Grid_Detalles_Turnos.RowSel
     Else
         Grid_Detalles_Turnos.Rows = 0
     End If
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
        Select Case Catalogo
            Case "Cat_Empresas"
                'Revisa que exista un registro a modificar y prepara la interfaz
                If Trim(Txt_Cat_Empresas_Empresa_ID.Text) <> "" Then
                    Fra_Cat_Empresas_Datos_Generales.Enabled = True
                    Fra_Cat_Empresas.Enabled = False
                    Txt_Cat_Empresas_Nombre.SetFocus
                    On Error Resume Next
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione una empresa para poder modificar", vbOKOnly + vbInformation, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Turnos": 'Catalogo de Turnos
                'Verifica la seleccion de un registro
                If Trim(Txt_Cat_Turnos_Turno_ID.Text) <> "" Then
                    Fra_Cat_Turnos_Generales.Enabled = True
                    Fra_Turnos_Detalles.Enabled = True
                    Fra_Cat_Turnos.Enabled = False
                    Txt_Cat_Turnos_Nombre.SetFocus
                    'SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un turno para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Calendarios_Turnos": 'Catalogo de Calendarios de Turnos
                'Verifica la seleccion de un registro
                If Trim(Txt_Calendario_Turno_ID.Text) <> "" Then
                    Fra_Calendarios_Turnos_Generales.Enabled = True
                    Fra_Calendarios_Turnos.Enabled = False
                    Fra_Calendarios_Configuración_Turnos.Enabled = True
                    Txt_Calendario_Nombre.SetFocus
                    Tab_Calendarios_Turnos.Tab = 1
                Else
                    MsgBox "Seleccione un Calendario para modificar", vbExclamation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Dias_No_Laborales": 'Catalogo de dias no laborales
                'Verifica la seleccion de un registro
                If Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text) <> "" Then
                    Fra_Cat_Dias_No_Laborales_Generales.Enabled = True
                    Fra_Cat_Dias_No_Laborales.Enabled = False
                    Txt_Cat_Dias_No_Laborales_Comentarios.SetFocus
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un dia para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Tipos_Faltas": 'Catalogo de tipos faltas
                'Verifica la seleccion de un registro
                If Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text) <> "" Then
                    Fra_Cat_Tipos_Faltas_Generales.Enabled = True
                    Fra_Cat_Tipos_Faltas.Enabled = False
                    Txt_Cat_Tipos_Faltas_Descripcion.SetFocus
                    'SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un tipo de falta para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
                
            Case "Cat_Departamentos": 'Catalogo de Departamentos
                'Verifica la seleccion de un registro
                If Trim(Txt_Departamento_ID.Text) <> "" Then
                    Fra_Departamento_Generales.Enabled = True
                    Fra_Departamentos.Enabled = False
                    Txt_Departamento_Clave.SetFocus
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un departamento para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Equipos_Identificacion":
                'Verifica la seleccion de un registro
                If Trim(Txt_Cat_Equipos_ID.Text) <> "" Then
                    Fra_Cat_Equipos_Generales.Enabled = True
                    Fra_Cat_Equipos.Enabled = False
                    Txt_Cat_Equipos_No_Equipo.SetFocus
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un equipo para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
            
            Case "Cat_Puestos"
                If Trim(Txt_Cat_Puestos_Puesto_ID.Text) <> "" Then
                    Fra_Cat_Puestos_Generales.Enabled = True
                    Fra_Cat_Puestos.Visible = True
                    Fra_Cat_Puestos.Enabled = False
                    Txt_Cat_Puestos_Abreviatura.SetFocus
                Else
                    MsgBox "Seleccione un Puesto para poder modificar", vbInformation
                    Exit Sub
                End If
            
            Case "Cat_Nivel_Estudio"
                If Trim(Txt_Cat_Nivel_Estudio_ID.Text) <> "" Then
                    Fra_Cat_Nivel_Estudio_Generales.Enabled = True
                    Fra_Cat_Nivel_Estudio.Visible = True
                    Fra_Cat_Nivel_Estudio.Enabled = False
                    Txt_Cat_Nivel_Estudio_Nombre.SetFocus
                Else
                    MsgBox "Seleccione un Nivel Estudio para poder modificar", vbInformation
                    Exit Sub
                End If
            
            Case "Cat_Motivos_Baja"
                If Trim(Txt_Cat_Motivos_Baja_ID.Text) <> "" Then
                    Fra_Cat_Motivos_Baja_Generales.Enabled = True
                    Fra_Cat_Motivos_Baja.Visible = True
                    Fra_Cat_Motivos_Baja.Enabled = False
                    Txt_Cat_Motivos_Baja_Nombre.SetFocus
                Else
                    MsgBox "Seleccione un Motivo de Baja para poder modificar", vbInformation
                    Exit Sub
                End If
                
            'Equipos Almacenes
            Case "Cat_Equipos_Almacenes_Identificacion":
                'Verifica la seleccion de un registro
                If Trim(Txt_Cat_Equipos_Almacenes_ID.Text) <> "" Then
                    Fra_Cat_Equipos_Almacenes_Generales.Enabled = True
                    Fra_Cat_Equipos_Almacenes.Enabled = False
                    Txt_Cat_Equipos_Almacenes_No_Equipo.SetFocus
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Seleccione un equipo para modificar", vbInformation + vbOKOnly, Me.Caption
                    Exit Sub
                End If
        End Select
        Btn_Modificar.Caption = "Actualizar"
        Btn_Eliminar.Enabled = False
        Btn_Nuevo.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Salir.Caption = "Regresar"
    Else
        Select Case Catalogo
            Case "Cat_Empresas":   'Modifica de Empresas
                If Trim(Txt_Cat_Empresas_Nombre.Text) <> "" Then
'                    If Trim(Txt_Cat_Empresas_Noi_Coi_ID.Text) <> "" Then
'                        If Trim(Txt_Cat_Empresas_Ruta_Noi.Text) <> "" Then
'                            If Trim(Txt_Cat_Empresas_Ruta_Coi.Text) <> "" Then
'                                If Cmb_Cat_Empresas_Tipo_Nomina.ListIndex > -1 Then
                                    Modifica_Cat_Empresas
'                                Else
'                                    MsgBox "Seleccione el tipo de Nomina", vbOKOnly + vbInformation, Me.Caption
'                                    Cmb_Cat_Empresas_Tipo_Nomina.SetFocus
'                                End If
'                            Else
'                                MsgBox "Seleccione la ruta del Sistema COI", vbOKOnly + vbInformation, Me.Caption
'                            End If
'                        Else
'                            MsgBox "Seleccione la ruta del Sistema NOI", vbOKOnly + vbInformation, Me.Caption
'                        End If
'                    Else
'                        MsgBox "Ingrese el Identificador de Empresa en NOI-COI", vbOKOnly + vbInformation, Me.Caption
'                    End If
                Else
                    MsgBox "Ingrese el Nombre de la empresa", vbOKOnly + vbInformation, Me.Caption
                End If
                
            Case "Cat_Turnos": 'Catalogo de Turnos
                'Valida la informacion obligatoria
                If Trim(Txt_Cat_Turnos_Nombre.Text) <> "" Then
                    Modifica_Cat_Turnos
                Else
                    MsgBox "Ingrese el nombre del turno", vbInformation + vbOKOnly, Me.Caption
                End If
                
            Case "Cat_Calendarios_Turnos": 'Catalogo de Calendarios Turnos
                'Valida la informacion obligatoria
                If Trim(Txt_Calendario_Nombre.Text) <> "" Then
                    Modifica_Calendario_Turnos
                Else
                    MsgBox "Ingrese el nombre del Calendario", vbExclamation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Dias_No_Laborales": 'Catalogo de dias no laborales
                'Valida la informacion obligatoria
                If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Dias_No_Laborales", "Fecha", Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy"), "Dia_No_Laboral_ID", Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text)) = False Then
                    Modifica_Cat_Dias_No_Laborales
                Else
                    MsgBox "Ya se ha registro un evento con la misma fecha", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Tipos_Faltas": 'Alta de tipo de falta
                If Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text) <> "" Then
                    If Trim(Txt_Cat_Tipos_Faltas_Simbologia) <> "" Then
                        Modifica_Cat_Tipos_Faltas
                    Else
                        MsgBox "Ingrese el Símbolo para la falta", vbOKOnly + vbInformation, Me.Caption
                        Txt_Cat_Tipos_Faltas_Simbologia.SetFocus
                    End If
                Else
                    MsgBox "Ingrese la descripcion de la falta", vbOKOnly + vbInformation, Me.Caption
                    Txt_Cat_Tipos_Faltas_Descripcion.SetFocus
                End If
                
            Case "Cat_Departamentos": 'Catalogo de Departamentos
                'Valida la informacion obligatoria
                If Trim(Txt_Departamento_Nombre.Text) <> "" Then
                    If Trim(Txt_Departamento_Clave.Text) <> "" Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Departamentos", "Clave", Trim(Txt_Departamento_Clave.Text), "Departamento_ID", Trim(Txt_Departamento_ID.Text)) Then
                            MsgBox "La clave que intenta asignar ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Modifica_Departamento
                Else
                    MsgBox "Faltan datos para actualizar", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Equipos_Identificacion":
                'Valida la informacion obligatoria
                If Trim(Txt_Cat_Equipos_Descripcion.Text) <> "" And Trim(Txt_Cat_Equipos_Direccion_IP.Text) <> "" And _
                    Val(Txt_Cat_Equipos_Puerto_IP.Text) > 0 And Trim(Txt_Cat_Equipos_Descripcion.Text) <> "" Then
                    If Val(Txt_Cat_Equipos_No_Equipo.Text) > 0 Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Equipos_Identificadores", "No_Equipo", Trim(Txt_Cat_Equipos_No_Equipo.Text), "Equipo_ID", Trim(Txt_Cat_Equipos_ID.Text)) Then
                            MsgBox "el no. de equipo ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Modifica_Cat_Equipos_Identificadores
                Else
                    MsgBox "Faltan datos para dar de alta", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Puestos"
                If Trim(Txt_Cat_Puestos_Abreviatura.Text) <> "" And Trim(Txt_Cat_Puestos_Nombre.Text) <> "" Then
                    Modifica_Cat_Puestos 'Da de alta los datos del banco en la base de datos
                    Fra_Cat_Puestos.Visible = True
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
            
            Case "Cat_Nivel_Estudio"
                If Trim(Txt_Cat_Nivel_Estudio_ID.Text) <> "" And Trim(Txt_Cat_Nivel_Estudio_Nombre.Text) <> "" Then
                    Modifica_Cat_Nivel_Estudio 'Da de alta los datos del banco en la base de datos
                    Fra_Cat_Nivel_Estudio.Visible = True
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
            
            Case "Cat_Motivos_Baja"
                If Trim(Txt_Cat_Motivos_Baja_ID.Text) <> "" And Trim(Txt_Cat_Motivos_Baja_Nombre.Text) <> "" Then
                    Modifica_Cat_Motivos_Baja 'Da de alta los datos del banco en la base de datos
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
                
            'Equipos Almacenes
            Case "Cat_Equipos_Almacenes_Identificacion":
                'Valida la informacion obligatoria
                If Trim(Txt_Cat_Equipos_Almacenes_Descripcion.Text) <> "" And Trim(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text) <> "" And _
                    Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text) > 0 And Trim(Txt_Cat_Equipos_Almacenes_Descripcion.Text) <> "" Then
                    If Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text) > 0 Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Equipos_Almacenes_Identificadores", "No_Equipo", Trim(Txt_Cat_Equipos_Almacenes_No_Equipo.Text), "Equipo_ID", Trim(Txt_Cat_Equipos_Almacenes_ID.Text)) Then
                            MsgBox "el no. de equipo ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Modifica_Cat_Equipos_Almacenes_Identificadores
                Else
                    MsgBox "Faltan datos para dar de alta", vbInformation + vbOKOnly, Me.Caption
                End If
            
        End Select
    End If
End Sub

Private Sub Btn_Nuevo_Click()
Dim Catacter As String 'Indica el caractere que se desea comparar
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Consultar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        'Muestra el picture del catalogo seleccionado
        Select Case Catalogo
            Case "Cat_Empresas": 'Catalogo de Empresas, Prepara la interfaz
                Txt_Cat_Empresas_Empresa_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empresas", "Empresa_ID"), "00000")
                Fra_Cat_Empresas_Datos_Generales.Enabled = True
                Fra_Cat_Empresas.Enabled = False
                Cmb_Cat_Empresas_Tipo_Nomina.ListIndex = 0
                Grid_Empresas_Equipos.Rows = 0
                Grid_Empresas_Equipos_Almacenes.Rows = 0
                Txt_Cat_Empresas_Nombre.SetFocus
            
            Case "Cat_Turnos": 'Catalogo de Turnos
                Txt_Cat_Turnos_Turno_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Turnos", "Turno_ID"), "00000")
                Fra_Cat_Turnos_Generales.Enabled = True
                Fra_Turnos_Detalles.Enabled = True
                Fra_Cat_Turnos.Enabled = False
                Dtp_Cat_Turnos_Hora_Inicio.Value = "00:00"
                Dtp_Cat_Turnos_Hora_Termino.Value = "00:00"
                Dtp_Cat_Turnos_Comida_Inicio.Value = "00:00"
                Dtp_Cat_Turnos_Comida_Termino.Value = "00:00"
                Txt_Cat_Turnos_Nombre.SetFocus
                Grid_Detalles_Turnos.Rows = 0
            
            Case "Cat_Calendarios_Turnos": 'Catalogo de Calendarios Turnos
                Txt_Calendario_Turno_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Calendarios_Turnos", "Calendario_Turno_ID"), "00000")
                Txt_Calendario_Estatus.Text = "ACTIVO"
                Fra_Calendarios_Turnos_Generales.Enabled = True
                Fra_Calendarios_Turnos.Enabled = False
                Fra_Calendarios_Configuración_Turnos.Enabled = True
                Txt_Calendario_Nombre.SetFocus
                Dtp_Calendario_Hora_Inicio.Value = "00:00"
                Dtp_Calendario_Hora_Termino.Value = "00:00"
                Dtp_Calendario_Inicio_Comida.Value = "00:00"
                Dtp_Calendario_Termino_Comida.Value = "00:00"
                Grid_Calendarios_Configuracion_Turnos.Rows = 0
                Call Lst_Calendarios_Configuracion_Empleados.Clear
                Tab_Calendarios_Turnos.Tab = 1
            
            Case "Cat_Dias_No_Laborales": 'Catalogo de Turnos
                Txt_Cat_Dias_No_Laborales_Dia_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Dias_No_Laborales", "Dia_No_Laboral_ID"), "00000")
                Fra_Cat_Dias_No_Laborales_Generales.Enabled = True
                Fra_Cat_Dias_No_Laborales.Enabled = False
                Dtp_Cat_Dias_No_Laborales_Fecha.Value = Now
                Dtp_Cat_Dias_No_Laborales_Fecha.SetFocus
            
            Case "Cat_Tipos_Faltas": 'Catalogo de Turnos
                Txt_Cat_Tipos_Faltas_Falta_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Faltas", "Tipo_Falta_ID"), "00000")
                Cmb_Clasificacion_Incidencias.ListIndex = 0
                Fra_Cat_Tipos_Faltas_Generales.Enabled = True
                Fra_Cat_Tipos_Faltas.Enabled = False
                Txt_Cat_Tipos_Faltas_Descripcion.SetFocus
            
            Case "Cat_Departamentos": 'Catalogo de Departamentos
                Txt_Departamento_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Departamentos", "Departamento_ID"), "00000")
                Fra_Departamento_Generales.Enabled = True
                Fra_Departamentos.Enabled = False
                Txt_Departamento_Clave.SetFocus
            
            Case "Cat_Equipos_Identificacion":
                Txt_Cat_Equipos_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Equipos_Identificadores", "Equipo_ID"), "00000")
                Txt_Cat_Equipos_Puerto_IP.Text = "4370"
                Fra_Cat_Equipos_Generales.Enabled = True
                Fra_Cat_Equipos.Enabled = False
                Txt_Cat_Equipos_No_Equipo.SetFocus
                        
            Case "Cat_Puestos"
                Txt_Cat_Puestos_Puesto_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Puestos", "Puesto_ID"), "00000")
                Fra_Cat_Puestos_Generales.Enabled = True
                Fra_Cat_Puestos.Visible = True
                Fra_Cat_Puestos.Enabled = False
                Txt_Cat_Puestos_Abreviatura.SetFocus
                
            Case "Cat_Nivel_Estudio"
                Txt_Cat_Nivel_Estudio_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Nivel_Estudio", "Nivel_Estudio_ID"), "00000")
                Fra_Cat_Nivel_Estudio_Generales.Enabled = True
                Fra_Cat_Nivel_Estudio.Visible = True
                Fra_Cat_Nivel_Estudio.Enabled = False
                Txt_Cat_Nivel_Estudio_Nombre.SetFocus
            
            Case "Cat_Motivos_Baja"
                Txt_Cat_Motivos_Baja_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Motivos_Baja", "Motivo_Baja_ID"), "00000")
                Fra_Cat_Motivos_Baja_Generales.Enabled = True
                Fra_Cat_Motivos_Baja.Visible = True
                Fra_Cat_Motivos_Baja.Enabled = False
                Txt_Cat_Motivos_Baja_Nombre.SetFocus
                
                'Equipos Almacenes
            Case "Cat_Equipos_Almacenes_Identificacion":
                Txt_Cat_Equipos_Almacenes_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Equipos_Almacenes_Identificadores", "Equipo_ID"), "00000")
                Txt_Cat_Equipos_Almacenes_Puerto_IP.Text = "4370"
                Txt_Cat_Equipos_Almacenes_Direccion_IP.Text = ""
                Fra_Cat_Equipos_Almacenes_Generales.Enabled = True
                Fra_Cat_Equipos_Almacenes.Enabled = False
                Txt_Cat_Equipos_Almacenes_No_Equipo.SetFocus
        End Select
    Else
        Select Case Catalogo
            Case "Cat_Empresas":   'Alta de Empresas
                If Trim(Txt_Cat_Empresas_Nombre.Text) <> "" Then
'                    If Trim(Txt_Cat_Empresas_Noi_Coi_ID.Text) <> "" Then
'                        If Trim(Txt_Cat_Empresas_Ruta_Noi.Text) <> "" Then
'                            If Trim(Txt_Cat_Empresas_Ruta_Coi.Text) <> "" Then
                                If Cmb_Cat_Empresas_Tipo_Nomina.ListIndex > -1 Then
                                    Alta_Cat_Empresas
                                Else
                                    MsgBox "Seleccione el tipo de Nomina", vbOKOnly + vbInformation, Me.Caption
                                    Cmb_Cat_Empresas_Tipo_Nomina.SetFocus
                                End If
'                            Else
'                                'MsgBox "Seleccione la ruta del Sistema COI", vbOKOnly + vbInformation, Me.Caption
'                                'Tab_Cat_Empresas.Tab = 1
'                                'Btn_Cat_Empresas_Ruta_COI.SetFocus
'                            End If
'                        Else
'                            'MsgBox "Seleccione la ruta del Sistema NOI", vbOKOnly + vbInformation, Me.Caption
'                            'Tab_Cat_Empresas.Tab = 1
'                            'Btn_Cat_Empresas_Ruta_NOI.SetFocus
'                        End If
'                    Else
'                        'MsgBox "Ingrese el Identificador de Empresa en NOI-COI", vbOKOnly + vbInformation, Me.Caption
'                        'Tab_Cat_Empresas.Tab = 1
'                        'Txt_Cat_Empresas_Noi_Coi_ID.SetFocus
'                    End If
                Else
                    MsgBox "Ingrese el Nombre de la empresa", vbOKOnly + vbInformation, Me.Caption
                    Txt_Cat_Empresas_Nombre.SetFocus
                End If
            
            Case "Cat_Turnos": 'Alta de Turnos
                If Trim(Txt_Cat_Turnos_Nombre.Text) <> "" Then
                    Alta_Cat_Turnos
                Else
                    MsgBox "Ingrese el nombre del turno", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Calendarios_Turnos": 'Alta de Catálogos de Turnos
                If Trim(Txt_Calendario_Nombre.Text) <> "" Then
                    If Grid_Calendarios_Configuracion_Turnos.Rows > 0 Then
                        Alta_Calendario_Turnos
                    Else
                        MsgBox "Agregue un calendario", vbExclamation + vbOKOnly, Me.Caption
                    End If
                Else
                    MsgBox "Ingrese el nombre para el calendario", vbExclamation + vbOKOnly, Me.Caption
                    Txt_Calendario_Nombre.SetFocus
                End If
                
            Case "Cat_Dias_No_Laborales": 'Alta de Dia no Laboral
                If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Dias_No_Laborales", "Fecha", Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy")) = False Then
                    Alta_Cat_Dias_No_Laborales
                Else
                    MsgBox "Ya se ha registro un evento con la misma fecha", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Tipos_Faltas": 'Alta de tipo de falta
                    If Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text) <> "" Then
                        If Trim(Txt_Cat_Tipos_Faltas_Simbologia) <> "" Then
                            Alta_Cat_Tipos_Faltas
                        Else
                            MsgBox "Ingrese el Símbolo para la falta", vbOKOnly + vbInformation, Me.Caption
                            Txt_Cat_Tipos_Faltas_Simbologia.SetFocus
                        End If
                    Else
                        MsgBox "Ingrese la descripcion de la falta", vbOKOnly + vbInformation, Me.Caption
                        Txt_Cat_Tipos_Faltas_Descripcion.SetFocus
                    End If
            
            Case "Cat_Departamentos": 'Catalogo de Departamentos
                'Valida la informacion obligatoria
                If Trim(Txt_Departamento_Nombre.Text) <> "" Then
                    If Trim(Txt_Departamento_Clave.Text) <> "" Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Departamentos", "Clave", Trim(Txt_Departamento_Clave.Text), "Departamento_ID", Trim(Txt_Departamento_ID.Text)) Then
                            MsgBox "La clave que intenta asignar ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Alta_Departamento
                Else
                    MsgBox "Faltan datos para dar de alta", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Equipos_Identificacion":
                'Valida la informacion obligatoria
                If Trim(Txt_Cat_Equipos_Descripcion.Text) <> "" And Trim(Txt_Cat_Equipos_Direccion_IP.Text) <> "" And _
                    Val(Txt_Cat_Equipos_Puerto_IP.Text) > 0 And Trim(Txt_Cat_Equipos_Descripcion.Text) <> "" Then
                    If Val(Txt_Cat_Equipos_No_Equipo.Text) > 0 Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Equipos_Identificadores", "No_Equipo", Trim(Txt_Cat_Equipos_No_Equipo.Text), "Equipo_ID", Trim(Txt_Cat_Equipos_ID.Text)) Then
                            MsgBox "el no. de equipo ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Alta_Cat_Equipos_Identificadores
                Else
                    MsgBox "Faltan datos para dar de alta", vbInformation + vbOKOnly, Me.Caption
                End If
            
            Case "Cat_Puestos"
                If Trim(Txt_Cat_Puestos_Abreviatura.Text) <> "" And Trim(Txt_Cat_Puestos_Nombre.Text) <> "" Then
                    Alta_Cat_Puestos 'Da de alta los datos del banco en la base de datos
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
                
             Case "Cat_Nivel_Estudio"
                If Trim(Txt_Cat_Nivel_Estudio_ID.Text) <> "" And Trim(Txt_Cat_Nivel_Estudio_Nombre.Text) <> "" Then
                    Alta_Cat_Nivel_Estudio 'Da de alta los datos del banco en la base de datos
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
            
            Case "Cat_Motivos_Baja"
                If Trim(Txt_Cat_Motivos_Baja_ID.Text) <> "" And Trim(Txt_Cat_Motivos_Baja_Nombre.Text) <> "" Then
                    Alta_Cat_Motivos_Baja 'Da de alta los datos del motivo de baja
                Else
                    MsgBox "Faltan datos por capturar", vbInformation
                    Exit Sub
                End If
                
                'Equipos Almacenes
             Case "Cat_Equipos_Almacenes_Identificacion":
                'Valida la informacion obligatoria
                If Trim(Txt_Cat_Equipos_Almacenes_Descripcion.Text) <> "" And Trim(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text) <> "" And _
                    Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text) > 0 And Trim(Txt_Cat_Equipos_Almacenes_Descripcion.Text) <> "" Then
                    If Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text) > 0 Then
                        If Conectar_Ayudante.Valida_Campo_Duplicado("Cat_Equipos_Almacenes_Identificadores", "No_Equipo", Trim(Txt_Cat_Equipos_Almacenes_No_Equipo.Text), "Equipo_ID", Trim(Txt_Cat_Equipos_Almacenes_ID.Text)) Then
                            MsgBox "el no. de equipo ya ha sido utilizado," + vbCrLf + "Favor de Verificar.", vbInformation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                    End If
                    Alta_Cat_Equipos_Almacenes_Identificadores
                Else
                    MsgBox "Faltan datos para dar de alta", vbInformation + vbOKOnly, Me.Caption
                End If
        End Select
    End If
End Sub



Private Sub Btn_Ruta_Logo_Click()
If Dir(App.Path & "\Logos_Empresas\", vbDirectory) = "" Then
    MkDir (App.Path & "\Logos_Empresas\")
End If
CommonDialog1.InitDir = App.Path & "\Logos_Empresas\"
CommonDialog1.Filter = "Archivos Jpg|*.jpg|Archivos Bmp|*.bmp|Archivos Gif|*.gif"
CommonDialog1.ShowOpen
Dim Ruta As String
Dim Nombre_Imagen As String
If CommonDialog1.FileName <> "" Then
Dim Datos() As String
Txt_Logo.Text = CommonDialog1.FileTitle
Logo_Temp = LoadPicture(CommonDialog1.FileName)
'Datos() = Split(CommonDialog1.FileName, "\")
'        Txt_Logo.Text = Datos(UBound(Datos))
'Else
'   MsgBox "No se seleccionó ningún archivo"
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
        Btn_Modificar.Caption = "Modificar"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        
        Select Case Catalogo
            Case "Cat_Empresas"
                Fra_Cat_Empresas_Datos_Generales.Enabled = False
                Fra_Cat_Empresas.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Empresas", Me)
            
            Case "Cat_Turnos": 'Catalogo de Turnos
                Fra_Cat_Turnos_Generales.Enabled = False
                Fra_Turnos_Detalles.Enabled = False
                Fra_Cat_Turnos.Enabled = True
                Dtp_Cat_Turnos_Hora_Inicio.Value = "12:00"
                Dtp_Cat_Turnos_Hora_Termino.Value = "12:00"
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Turnos", Me)
            
            Case "Cat_Calendarios_Turnos": 'Catalogo de Calendarios de Turnos
                Fra_Calendarios_Turnos_Generales.Enabled = False
                Fra_Calendarios_Configuración_Turnos.Enabled = False
                Fra_Calendarios_Turnos.Enabled = True
                Tab_Calendarios_Turnos.Tab = 0
'                Dim Cont_I As Integer
'                For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
'                    Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = False
'                Next Cont_I
                Call Lst_Calendarios_Configuracion_Empleados.Clear
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Btn_Adm_RH_Panel_Cat_Calendarios_Turnos", Me)
            
            Case "Cat_Dias_No_Laborales": 'Catalogo de Dias no laborales
                Fra_Cat_Dias_No_Laborales_Generales.Enabled = False
                Fra_Cat_Dias_No_Laborales.Enabled = True
                Dtp_Cat_Dias_No_Laborales_Fecha.Value = Now
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Dias_No_Laborales", Me)
            
            Case "Cat_Tipos_Faltas": 'Catalogo de Dias no laborales
                Fra_Cat_Tipos_Faltas_Generales.Enabled = False
                Fra_Cat_Tipos_Faltas.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Tipos_Faltas", Me)
            
            Case "Cat_Departamentos": 'Catalogo de Departamentos
                Fra_Departamento_Generales.Enabled = False
                Fra_Departamentos.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Departamentos", Me)
            
            Case "Cat_Equipos_Identificacion":
                Fra_Cat_Equipos_Generales.Enabled = False
                Fra_Cat_Equipos.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Equipos_Identificacion", Me)
            
            Case "Cat_Puestos"
                Fra_Cat_Puestos_Generales.Enabled = False
                Fra_Cat_Puestos.Visible = True
                Fra_Cat_Puestos.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Puestos", Me)
                
            Case "Cat_Nivel_Estudio"
                Fra_Cat_Nivel_Estudio_Generales.Enabled = False
                Fra_Cat_Nivel_Estudio.Visible = True
                Fra_Cat_Nivel_Estudio.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Nivel_Estudio", Me)
            
            Case "Cat_Motivos_Baja"
                Fra_Cat_Motivos_Baja_Generales.Enabled = False
                Fra_Cat_Motivos_Baja.Enabled = True
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Motivos_Baja", Me)
        End Select
    End If
End Sub

Private Sub Dtp_Calendario_Fecha_Inicio_CloseUp()
'    If DatePart("w", Dtp_Calendario_Fecha_Inicio.Value) = vbSunday Then
''        Dtp_Calendario_Fecha_Inicio.Value = DateAdd("d", -6, Dtp_Calendario_Fecha_Inicio.Value)
'        Dtp_Calendario_Fecha_Inicio.Value = DateAdd("d", -7, Dtp_Calendario_Fecha_Inicio.Value)
'    Else
''        Dtp_Calendario_Fecha_Inicio.Value = DateAdd("d", -(DatePart("w", Dtp_Calendario_Fecha_Inicio.Value) - vbMonday), Dtp_Calendario_Fecha_Inicio.Value)
'        Dtp_Calendario_Fecha_Inicio.Value = DateAdd("d", -(DatePart("w", Dtp_Calendario_Fecha_Inicio.Value) - vbSunday), Dtp_Calendario_Fecha_Inicio.Value)
'    End If
    Call Crear_Calendario
End Sub

Private Sub Dtp_Calendario_Fecha_Inicio_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Dtp_Calendario_Fecha_Inicio_CloseUp
    End If
End Sub

Private Sub Dtp_Calendario_Fecha_Termino_CloseUp()
'    If DatePart("w", Dtp_Calendario_Fecha_Termino.Value) <> vbSunday Then
''        Dtp_Calendario_Fecha_Termino.Value = DateAdd("d", ((vbSaturday + 1) - DatePart("w", Dtp_Calendario_Fecha_Termino.Value)), Dtp_Calendario_Fecha_Termino.Value)
'        Dtp_Calendario_Fecha_Termino.Value = DateAdd("d", ((vbSaturday) - DatePart("w", Dtp_Calendario_Fecha_Termino.Value)), Dtp_Calendario_Fecha_Termino.Value)
'    End If
    Call Crear_Calendario
End Sub

Private Sub Dtp_Calendario_Fecha_Termino_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Dtp_Calendario_Fecha_Termino_CloseUp
    End If
End Sub

Private Sub Dtp_Calendario_Hora_Inicio_Change()
    Call Calcular_Horas_Calendarios_Turnos
End Sub

Private Sub Dtp_Calendario_Hora_Termino_Change()
    Call Calcular_Horas_Calendarios_Turnos
End Sub

Private Sub Dtp_Calendario_Inicio_Comida_Change()
    Call Calcular_Horas_Calendarios_Turnos
End Sub

Private Sub Dtp_Calendario_Termino_Comida_Change()
    Call Calcular_Horas_Calendarios_Turnos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 _
    And Shift = 2 Then
        Copia_Configuracion_Calendario = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 _
    And Shift = 0 Then
        Copia_Configuracion_Calendario = False
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7455
    Me.Width = 8525
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Grid_Calendarios_Configuracion_Turnos_Click()
Dim Lista_Empleados() As String
Dim Numeros_Empleados As String
Dim Cont_I As Integer
Dim Cont_L As Integer

'    If Not Ejecutando_MouseUp _
'    And Not Ejecutando_LostFocus Then
        If Grid_Calendarios_Configuracion_Turnos.Rows > Grid_Calendarios_Configuracion_Turnos.FixedRows Then
            Temp_Enter_Cell = True
            If Copia_Configuracion_Calendario Then
                If Txt_Calendarios_Configuracion_Turno.Enabled Then
                    If Trim(Txt_Calendarios_Configuracion_Turno.Text) <> "" Then
                        Call Asignar_Turno_Calendario
                    Else
                        MsgBox "Debe escribir un nombre de Turno", vbExclamation, "Calendario de Turnos"
                        Txt_Calendarios_Configuracion_Turno.SetFocus
                    End If
                Else
                    Call Btn_Configuracion_Calendarios_Turnos_Limpiar_Datos_Click
                End If
            Else
                If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col)) = "" Then
                    If Txt_Calendarios_Configuracion_Turno.Enabled Then
                        If Trim(Txt_Calendarios_Configuracion_Turno.Text) <> "" Then
                            Call Asignar_Turno_Calendario
                        Else
                            MsgBox "Debe escribir un nombre de Turno", vbExclamation, "Calendario de Turnos"
                            Txt_Calendarios_Configuracion_Turno.SetFocus
                        End If
                    Else
                        Call Btn_Configuracion_Calendarios_Turnos_Limpiar_Datos_Click
                    End If
                Else
                    Call Obtener_Horario_Turno_Grid
                    Call Calcular_Horas_Calendarios_Turnos
                End If
            End If
    '   Else
    '       MsgBox "Debe seleccionar un rango de fechas para armar la plantilla del Calendario", vbExclamation, "Calendario de Turnos"
        End If
        
        If Grid_Calendarios_Configuracion_Turnos.Row >= Grid_Calendarios_Configuracion_Turnos.FixedRows _
        And Grid_Calendarios_Configuracion_Turnos.Col >= Grid_Calendarios_Configuracion_Turnos.FixedCols Then
            Temp_Col = Grid_Calendarios_Configuracion_Turnos.Col
            Temp_Row = Grid_Calendarios_Configuracion_Turnos.Row
            If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col)) <> "" Then
                Numeros_Empleados = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1)
                If Trim(Numeros_Empleados) <> "" Then
                    Call Conectar_Ayudante.Llena_List_Item("No_Tarjeta, CAST(No_Tarjeta AS VARCHAR)+' - '+Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno", "Cat_Empleados WHERE No_Tarjeta IN (" & Numeros_Empleados & ") AND Estatus = 'A'", Lst_Calendarios_Configuracion_Empleados, 0, "Cat_Empleados.No_Tarjeta")
                End If
                Lista_Empleados = Split(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1), ",")
                Lst_Calendarios_Configuracion_Empleados.ListIndex = -1
                For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
                    Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = False
                Next Cont_I
                For Cont_L = 0 To UBound(Lista_Empleados)
                    For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
                        If Lst_Calendarios_Configuracion_Empleados.ItemData(Cont_I) = Lista_Empleados(Cont_L) Then
                            Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = True
                        End If
                    Next Cont_I
                Next Cont_L
            End If
        End If
'    End If
'    Ejecutando_MouseUp = False
'    Ejecutando_LostFocus = False
End Sub

Private Sub Grid_Calendarios_Configuracion_Turnos_DblClick()
Dim Cont_I As Integer
    If Txt_Calendarios_Configuracion_Turno.Enabled Then
        Ejecutando_Grid_Calendarios_Configuracion_Turnos_DblClick = True
        If Grid_Calendarios_Configuracion_Turnos.Rows > 0 _
        And Grid_Calendarios_Configuracion_Turnos.Cols > 0 _
        And Grid_Calendarios_Configuracion_Turnos.MouseRow >= Grid_Calendarios_Configuracion_Turnos.FixedRows _
        And Grid_Calendarios_Configuracion_Turnos.MouseCol >= Grid_Calendarios_Configuracion_Turnos.FixedCols Then
            If Grid_Calendarios_Configuracion_Turnos.Col <= (Grid_Calendarios_Configuracion_Turnos.Cols - 5) Then
                Grid_Calendarios_Configuracion_Turnos.CellBackColor = 0
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 2) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 3) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 4) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 5) = ""
                For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
                    Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = False
                Next Cont_I
            End If
        End If
    End If
End Sub

Private Sub Grid_Calendarios_Configuracion_Turnos_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Tool_Tip_Text As String

    If Grid_Calendarios_Configuracion_Turnos.Rows > 0 _
    And Grid_Calendarios_Configuracion_Turnos.Cols > 0 _
    And Grid_Calendarios_Configuracion_Turnos.MouseRow >= Grid_Calendarios_Configuracion_Turnos.FixedRows _
    And Grid_Calendarios_Configuracion_Turnos.MouseCol >= Grid_Calendarios_Configuracion_Turnos.FixedCols Then
        If Temp_MouseRow <> Grid_Calendarios_Configuracion_Turnos.MouseRow _
        Or Temp_MouseCol <> Grid_Calendarios_Configuracion_Turnos.MouseCol Then
            Temp_MouseRow = Grid_Calendarios_Configuracion_Turnos.MouseRow
            Temp_MouseCol = Grid_Calendarios_Configuracion_Turnos.MouseCol
            Grid_Calendarios_Configuracion_Turnos.ToolTipText = ""
        End If
        If Grid_Calendarios_Configuracion_Turnos.MouseCol < (Grid_Calendarios_Configuracion_Turnos.Cols - 5) Then
            Tool_Tip_Text = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.MouseRow, Grid_Calendarios_Configuracion_Turnos.MouseCol)
            If Trim(Tool_Tip_Text) <> "" Then
                Tool_Tip_Text = Tool_Tip_Text & ":" & Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.MouseRow, Grid_Calendarios_Configuracion_Turnos.MouseCol + 2)
                Tool_Tip_Text = Tool_Tip_Text & "-" & Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.MouseRow, Grid_Calendarios_Configuracion_Turnos.MouseCol + 3)
        '        Tool_Tip_Text = Tool_Tip_Text & "|" & Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.MouseRow, Grid_Calendarios_Configuracion_Turnos.MouseCol+4)
        '        Tool_Tip_Text = Tool_Tip_Text & ":" & Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.MouseRow, Grid_Calendarios_Configuracion_Turnos.MouseCol+5)
            Else
                Tool_Tip_Text = ""
            End If
        End If
        Grid_Calendarios_Configuracion_Turnos.ToolTipText = Trim(Tool_Tip_Text)
    End If
End Sub

'Private Sub Grid_Calendarios_Configuracion_Turnos_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'Dim Lista_Empleados() As String
'Dim Cont_L As Integer
'Dim Cont_I As Integer
'    If Button = 2 Then
'        Ejecutando_MouseUp = True
''        If Lst_Calendarios_Configuracion_Empleados.Visible Then
'            Grid_Calendarios_Configuracion_Turnos.TextMatrix(Temp_Row, Temp_Col + 1) = Obtener_Lista_Empleados_Seleccionados
''        End If
'        Grid_Calendarios_Configuracion_Turnos.Col = Grid_Calendarios_Configuracion_Turnos.MouseCol
'        Grid_Calendarios_Configuracion_Turnos.Row = Grid_Calendarios_Configuracion_Turnos.MouseRow
'        Temp_Col = Grid_Calendarios_Configuracion_Turnos.Col
'        Temp_Row = Grid_Calendarios_Configuracion_Turnos.Row
'        If Temp_Row >= Grid_Calendarios_Configuracion_Turnos.FixedRows _
'        And Temp_Col >= Grid_Calendarios_Configuracion_Turnos.FixedCols Then
'            If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col)) <> "" Then
''                Lst_Calendarios_Configuracion_Empleados.Left = Tab_Calendarios_Turnos.Left + Fra_Calendarios_Configuración_Turnos.Left + Grid_Calendarios_Configuracion_Turnos.Left + X
''                Lst_Calendarios_Configuracion_Empleados.Top = Tab_Calendarios_Turnos.Top + Fra_Calendarios_Configuración_Turnos.Top + Grid_Calendarios_Configuracion_Turnos.Top + y
'                Lista_Empleados = Split(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1), ",")
'                Lst_Calendarios_Configuracion_Empleados.ListIndex = -1
'                For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
'                    Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = False
'                Next Cont_I
'                For Cont_L = 0 To UBound(Lista_Empleados)
'                    For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
'                        If Lst_Calendarios_Configuracion_Empleados.ItemData(Cont_I) = Lista_Empleados(Cont_L) Then
'                            Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) = True
'                        End If
'                    Next Cont_I
'                Next Cont_L
''                Lst_Calendarios_Configuracion_Empleados.Visible = True
''                Lst_Calendarios_Configuracion_Empleados.SetFocus
'            End If
'        End If
''        Ejecutando_MouseUp = False
'    End If
'End Sub

Private Sub Grid_Calendarios_Turnos_Click()
Dim Rs_Consulta_Calendario_Turnos As rdoResultset
    If Grid_Calendarios_Turnos.Rows > 1 Then
        Call Btn_Configuracion_Calendarios_Turnos_Limpiar_Datos_Click
        'Consulta los datos del turno
        Mi_SQL = "SELECT  Calendario_Turno_ID,Estatus,Nombre,Fecha_Inicio,Fecha_Termino,Comentarios"
        Mi_SQL = Mi_SQL & " FROM Cat_Calendarios_Turnos"
        Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID='" & Trim(Grid_Calendarios_Turnos.TextMatrix(Grid_Calendarios_Turnos.RowSel, 0)) & "'"
        Set Rs_Consulta_Calendario_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Calendario_Turnos.EOF Then
            Txt_Calendario_Turno_ID.Text = Rs_Consulta_Calendario_Turnos.rdoColumns("Calendario_Turno_ID")
            Txt_Calendario_Estatus.Text = Rs_Consulta_Calendario_Turnos.rdoColumns("Estatus")
            Txt_Calendario_Nombre.Text = Rs_Consulta_Calendario_Turnos.rdoColumns("Nombre")
            Txt_Calendario_Comentarios.Text = Rs_Consulta_Calendario_Turnos.rdoColumns("Comentarios")
            Dtp_Calendario_Fecha_Inicio.Value = Rs_Consulta_Calendario_Turnos.rdoColumns("Fecha_Inicio")
            Dtp_Calendario_Fecha_Termino.Value = Rs_Consulta_Calendario_Turnos.rdoColumns("Fecha_Termino")
        End If
        Rs_Consulta_Calendario_Turnos.Close
        Grid_Calendarios_Configuracion_Turnos.Rows = 0
        Grid_Calendarios_Configuracion_Turnos.Cols = 0
        Call Crear_Calendario
        Tab_Calendarios_Turnos.Tab = 1
        Call Consulta_Detalles_Calendarios_Turnos
        Call Calcular_Horas_Calendarios_Turnos
        Call Lst_Calendarios_Configuracion_Empleados.Clear
    End If
End Sub

Private Sub Grid_Cat_Dias_No_Laborales_Click()
    With Grid_Cat_Dias_No_Laborales
        If .Rows > 1 Then
            Txt_Cat_Dias_No_Laborales_Dia_ID.Text = Trim(.TextMatrix(.RowSel, 0))
            Dtp_Cat_Dias_No_Laborales_Fecha.Value = CDate(.TextMatrix(.RowSel, 1))
            Txt_Cat_Dias_No_Laborales_Comentarios.Text = Trim(.TextMatrix(.RowSel, 2))
        End If
    End With
End Sub

Private Sub Grid_Cat_Dias_No_Laborales_EnterCell()
    Grid_Cat_Dias_No_Laborales_Click
End Sub

Private Sub Grid_Cat_Empresas_Click()
Dim Rs_Consulta_Cat_Empresas As rdoResultset    'Informacion de la empresa
Dim Rs_Consulta_Cat_Empresas_Equipos_Identificacion As rdoResultset    'Informacion de la empresa
Dim Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes As rdoResultset    'Informacion de los equipos de Almacen

With Grid_Cat_Empresas
    If .Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Empresas"
        Mi_SQL = Mi_SQL & " WHERE Empresa_ID ='" & Trim(.TextMatrix(.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Empresas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Cat_Empresas
            If Not .EOF Then
                Txt_Cat_Empresas_Empresa_ID.Text = .rdoColumns("Empresa_ID")
                Txt_Cat_Empresas_Acronimo.Text = .rdoColumns("Acronimo")
                Txt_Cat_Empresas_Nombre.Text = .rdoColumns("Nombre")
                Txt_Cat_Empresas_RFC.Text = .rdoColumns("RFC")
                Txt_Cat_Empresas_Direccion.Text = .rdoColumns("Direccion")
                Txt_Cat_Empresas_Colonia.Text = .rdoColumns("Colonia")
                Txt_Cat_Empresas_Ciudad.Text = .rdoColumns("Ciudad")
                Txt_Cat_Empresas_Estado.Text = .rdoColumns("Estado")
                Txt_Cat_Empresas_CP.Text = .rdoColumns("Codigo_Postal")
                Txt_Cat_Empresas_Telefono.Text = .rdoColumns("Telefono")
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Cat_Empresas_Comentarios.Text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("Tipo_Nomina")) Then
                    Cmb_Cat_Empresas_Tipo_Nomina.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Tipo_Nomina")), Cmb_Cat_Empresas_Tipo_Nomina)
                Else
                    Cmb_Cat_Empresas_Tipo_Nomina.ListIndex = -1
                End If
                If Not IsNull(.rdoColumns("Logo")) Then
                    Txt_Logo = .rdoColumns("Logo")
                End If
                Grid_Empresas_Equipos.Rows = 0
                'Llena los checadores de la empresa
                Mi_SQL = "SELECT CEEI.Equipo_ID, (cast(CEI.No_Equipo as varchar)+' '+CEI.Descripcion) as Equipo"
                Mi_SQL = Mi_SQL & " FROM Cat_Empresas_Equipos_Identificacion CEEI, Cat_Equipos_Identificadores CEI"
                Mi_SQL = Mi_SQL & " WHERE CEEI.Equipo_ID = CEI.Equipo_ID"
                Mi_SQL = Mi_SQL & " AND CEEI.Empresa_ID = '" & .rdoColumns("Empresa_ID") & "'"
                Set Rs_Consulta_Cat_Empresas_Equipos_Identificacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empresas_Equipos_Identificacion.EOF Then
                        Grid_Empresas_Equipos.Cols = 2
                        If Grid_Empresas_Equipos.Rows = 0 Then
                            Grid_Empresas_Equipos.AddItem "Equipo_ID" & Chr(9) & "Equipo"
                            Grid_Empresas_Equipos.ColWidth(0) = 0    'Equipo_ID
                            Grid_Empresas_Equipos.ColWidth(1) = 3500 'Equipo
                            Grid_Empresas_Equipos.ColAlignment(1) = flexAlignLeftCenter
                        End If
                        While Not Rs_Consulta_Cat_Empresas_Equipos_Identificacion.EOF
                                Grid_Empresas_Equipos.AddItem Rs_Consulta_Cat_Empresas_Equipos_Identificacion.rdoColumns("Equipo_ID") & Chr(9) & _
                                    Rs_Consulta_Cat_Empresas_Equipos_Identificacion.rdoColumns("Equipo")
                            Rs_Consulta_Cat_Empresas_Equipos_Identificacion.MoveNext
                        Wend
                        Grid_Empresas_Equipos.FixedRows = 1
                    End If
                Set Rs_Consulta_Cat_Empresas_Equipos_Identificacion = Nothing
'                .Close
'            End If
                 Grid_Empresas_Equipos_Almacenes.Rows = 0
                'Llena los checadores almacenes de la empresa
                Mi_SQL = "SELECT CEEI.Equipo_ID, (cast(CEI.No_Equipo as varchar)+' '+CEI.Descripcion) as Equipo"
                Mi_SQL = Mi_SQL & " FROM Cat_Empresas_Equipos_Identificacion_Almacenes CEEI, Cat_Equipos_Almacenes_Identificadores CEI"
                Mi_SQL = Mi_SQL & " WHERE CEEI.Equipo_ID = CEI.Equipo_ID"
                Mi_SQL = Mi_SQL & " AND CEEI.Empresa_ID = '" & .rdoColumns("Empresa_ID") & "'"
                Set Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes.EOF Then
                        Grid_Empresas_Equipos_Almacenes.Cols = 2
                        If Grid_Empresas_Equipos_Almacenes.Rows = 0 Then
                            Grid_Empresas_Equipos_Almacenes.AddItem "Equipo_ID" & Chr(9) & "Equipo"
                            Grid_Empresas_Equipos_Almacenes.ColWidth(0) = 0    'Equipo_ID
                            Grid_Empresas_Equipos_Almacenes.ColWidth(1) = 3500 'Equipo
                            Grid_Empresas_Equipos_Almacenes.ColAlignment(1) = flexAlignLeftCenter
                        End If
                        While Not Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes.EOF
                                Grid_Empresas_Equipos_Almacenes.AddItem Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes.rdoColumns("Equipo_ID") & Chr(9) & _
                                    Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes.rdoColumns("Equipo")
                            Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes.MoveNext
                        Wend
                        Grid_Empresas_Equipos_Almacenes.FixedRows = 1
                    End If
                Set Rs_Consulta_Cat_Empresas_Equipos_Identificacion_Almacenes = Nothing
                .Close
            End If
            Set Rs_Consulta_Cat_Empresas = Nothing
        End With
    End If
End With
End Sub

Private Sub Grid_Cat_Empresas_EnterCell()
    Grid_Cat_Empresas_Click
End Sub

Private Sub Grid_Cat_Equipos_Almacenes_Click()
With Grid_Cat_Equipos_Almacenes
    If .Rows > 1 Then
        Txt_Cat_Equipos_Almacenes_ID.Text = Trim(.TextMatrix(.RowSel, 0))
        Txt_Cat_Equipos_Almacenes_No_Equipo.Text = Trim(.TextMatrix(.RowSel, 1))
        Txt_Cat_Equipos_Almacenes_Direccion_IP.Text = Trim(.TextMatrix(.RowSel, 2))
        Txt_Cat_Equipos_Almacenes_Puerto_IP.Text = Trim(.TextMatrix(.RowSel, 3))
        Txt_Cat_Equipos_Almacenes_Descripcion.Text = Trim(.TextMatrix(.RowSel, 4))
    End If
End With
End Sub
Private Sub Grid_Cat_Equipos_Click()
With Grid_Cat_Equipos
    If .Rows > 1 Then
        Txt_Cat_Equipos_ID.Text = Trim(.TextMatrix(.RowSel, 0))
        Txt_Cat_Equipos_No_Equipo.Text = Trim(.TextMatrix(.RowSel, 1))
        Txt_Cat_Equipos_Direccion_IP.Text = Trim(.TextMatrix(.RowSel, 2))
        Txt_Cat_Equipos_Puerto_IP.Text = Trim(.TextMatrix(.RowSel, 3))
        Txt_Cat_Equipos_Descripcion.Text = Trim(.TextMatrix(.RowSel, 4))
    End If
End With
End Sub

Private Sub Grid_Cat_Equipos_EnterCell()
    Grid_Cat_Equipos_Click
End Sub
Private Sub Grid_Cat_Equipos_Almacenes_EnterCell()
    Grid_Cat_Equipos_Almacenes_Click
End Sub
Private Sub Grid_Cat_Motivos_Baja_Click()
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    With Grid_Cat_Motivos_Baja
        If .Rows > 1 Then
            Txt_Cat_Motivos_Baja_ID.Text = Trim(.TextMatrix(.RowSel, 0))
            Txt_Cat_Motivos_Baja_Nombre.Text = Trim(.TextMatrix(.RowSel, 1))
            Txt_Cat_Motivos_Baja_Descripcion.Text = Trim(.TextMatrix(.RowSel, 2))
            Txt_Clave_SAP_Motivos_Baja.Text = Trim(.TextMatrix(.RowSel, 3))
        End If
    End With
End Sub

Private Sub Grid_Cat_Motivos_Baja_EnterCell()
    Grid_Cat_Motivos_Baja_Click
End Sub

Private Sub Grid_Cat_Nivel_Estudio_Click()
Call Conectar_Ayudante.Limpiar_Textos(Me)
With Grid_Cat_Nivel_Estudio
    If .Rows > 1 Then
        Txt_Cat_Nivel_Estudio_ID.Text = Trim(.TextMatrix(.RowSel, 0))
        Txt_Cat_Nivel_Estudio_Nombre.Text = Trim(.TextMatrix(.RowSel, 1))
        Txt_Cat_Nivel_Estudio_Descripcion.Text = Trim(.TextMatrix(.RowSel, 2))
    End If
End With
End Sub

Private Sub Grid_Cat_Nivel_Estudio_EnterCell()
    Grid_Cat_Nivel_Estudio_Click
End Sub

Private Sub Grid_Cat_Puestos_Click()
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    With Grid_Cat_Puestos
        If .Rows > 1 Then
            Txt_Cat_Puestos_Puesto_ID.Text = Trim(.TextMatrix(.RowSel, 0))
            Txt_Cat_Puestos_Nombre.Text = Trim(.TextMatrix(.RowSel, 1))
            Txt_Cat_Puestos_Abreviatura.Text = Trim(.TextMatrix(.RowSel, 2))
            Txt_Cat_Puestos_Comentarios.Text = Trim(.TextMatrix(.RowSel, 3))
            Txt_Clave_SAP_Puestos.Text = Trim(.TextMatrix(.RowSel, 4))
        End If
    End With
End Sub

Private Sub Grid_Cat_Puestos_EnterCell()
    Grid_Cat_Puestos_Click
End Sub

Private Sub Grid_Cat_Tipos_Faltas_Click()
Dim Rs_Consulta_Cat_Tipos_Faltas As rdoResultset    'Informacion de la empresa
    If Grid_Cat_Tipos_Faltas.Rows > 1 Then
        'Consulta el catálogo
        Mi_SQL = "SELECT * FROM Cat_Tipos_Faltas"
        Mi_SQL = Mi_SQL & " WHERE Tipo_Falta_ID='" & Trim(Grid_Cat_Tipos_Faltas.TextMatrix(Grid_Cat_Tipos_Faltas.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Tipos_Faltas.EOF Then
            With Rs_Consulta_Cat_Tipos_Faltas
                Txt_Cat_Tipos_Faltas_Falta_ID.Text = .rdoColumns("Tipo_Falta_ID")
                If Not IsNull(.rdoColumns("Clave_SAP")) Then
                    Txt_Clave_SAP_Tipos_Faltas.Text = .rdoColumns("Clave_SAP")
                Else
                    Txt_Clave_SAP_Tipos_Faltas.Text = ""
                End If
                Txt_Cat_Tipos_Faltas_Descripcion.Text = .rdoColumns("Descripcion")
                If Not IsNull(.rdoColumns("Simbologia")) Then Txt_Cat_Tipos_Faltas_Simbologia.Text = .rdoColumns("Simbologia")
                Txt_Cat_Tipos_Faltas_Comentarios.Text = .rdoColumns("Comentarios")
                If Not IsNull(.rdoColumns("Clave_SAP")) Then
                    Txt_Clave_SAP_Tipos_Faltas.Text = .rdoColumns("Clave_SAP")
                Else
                    Txt_Clave_SAP_Tipos_Faltas.Text = ""
                End If
                If Not IsNull(.rdoColumns("Clasificacion")) Then
                    Cmb_Clasificacion_Incidencias.Text = .rdoColumns("Clasificacion")
                Else
                    Cmb_Clasificacion_Incidencias.ListIndex = 0
                End If
            End With
        End If
        Rs_Consulta_Cat_Tipos_Faltas.Close
    End If
End Sub

Private Sub Grid_Cat_Tipos_Faltas_EnterCell()
    Grid_Cat_Tipos_Faltas_Click
End Sub

Private Sub Grid_Cat_Turnos_Click()
Dim Rs_Consulta_Cat_Turnos As rdoResultset
    If Grid_Cat_Turnos.Rows > 1 Then
        'Consulta los datos del turno
        Mi_SQL = "SELECT  Turno_ID,Nombre,Hora_Inicio,Hora_Termino,Comentarios"
        Mi_SQL = Mi_SQL & " ,ISNULL(Comida_Inicio,'00:00') AS Comida_Inicio,ISNULL(Comida_Termino,'00:00') AS Comida_Termino"
        Mi_SQL = Mi_SQL & " ,ISNULL(Horas_Turno,0) AS Horas_Turno,ISNULL(Horas_Comida,0) AS Horas_Comida"
        Mi_SQL = Mi_SQL & " FROM Cat_Turnos"
        Mi_SQL = Mi_SQL & " WHERE Turno_ID='" & Trim(Grid_Cat_Turnos.TextMatrix(Grid_Cat_Turnos.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Turnos.EOF Then
            Txt_Cat_Turnos_Turno_ID.Text = Rs_Consulta_Cat_Turnos.rdoColumns("Turno_ID")
            Txt_Cat_Turnos_Nombre.Text = Rs_Consulta_Cat_Turnos.rdoColumns("Nombre")
            Txt_Cat_Turnos_Comentarios.Text = Rs_Consulta_Cat_Turnos.rdoColumns("Comentarios")
        End If
        Rs_Consulta_Cat_Turnos.Close
        Call Consulta_Detalles_Turno
    End If
End Sub

Private Sub Grid_Cat_Turnos_EnterCell()
    Grid_Cat_Turnos_Click
End Sub

Private Sub Grid_Departamentos_Click()
    If Grid_Departamentos.Rows > 1 Then
        Txt_Departamento_ID.Text = Trim(Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 0))
        Txt_Departamento_Nombre.Text = Trim(Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 1))
        Txt_Departamento_Clave.Text = Trim(Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 2))
        Txt_Departamento_Comentarios.Text = Trim(Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 3))
        Txt_Clave_SAP_Departamentos.Text = Trim(Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 4))
    End If
End Sub

Private Sub Grid_Departamentos_EnterCell()
    Grid_Departamentos_Click
End Sub

Private Sub Lst_Calendarios_Configuracion_Empleados_ItemCheck(Item As Integer)
    If Not Ejecutando_Grid_Calendarios_Configuracion_Turnos_DblClick Then
        If Grid_Calendarios_Configuracion_Turnos.Cols > 0 _
        And Grid_Calendarios_Configuracion_Turnos.Rows > 0 Then
            If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col)) = "" Then
                Lst_Calendarios_Configuracion_Empleados.Selected(Item) = Not Lst_Calendarios_Configuracion_Empleados.Selected(Item)
            End If
        Else
            On Error Resume Next
            Lst_Calendarios_Configuracion_Empleados.Selected(Item) = Not Lst_Calendarios_Configuracion_Empleados.Selected(Item)
        End If
    End If
End Sub

Private Sub Lst_Calendarios_Configuracion_Empleados_LostFocus()
'    Ejecutando_LostFocus = True
'    Lst_Calendarios_Configuracion_Empleados.Visible = False
    If Grid_Calendarios_Configuracion_Turnos.Cols > 0 _
    And Grid_Calendarios_Configuracion_Turnos.Rows > 0 Then
        If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Temp_Row, Temp_Col + 1)) <> "" _
        And Not Temp_Enter_Cell Then
            Grid_Calendarios_Configuracion_Turnos.TextMatrix(Temp_Row, Temp_Col + 1) = Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Temp_Row, Temp_Col + 1)) & "," & Obtener_Lista_Empleados_Seleccionados
        Else
            Grid_Calendarios_Configuracion_Turnos.TextMatrix(Temp_Row, Temp_Col + 1) = Obtener_Lista_Empleados_Seleccionados
            Temp_Enter_Cell = False
        End If
    End If
End Sub

Private Sub Txt_Calendario_Filtro_Empleados_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_List_Item("No_Tarjeta, CAST(No_Tarjeta AS VARCHAR)+' - '+Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno", "Cat_Empleados WHERE (Nombre LIKE '%" & Trim(Txt_Calendario_Filtro_Empleados.Text) & "%' OR Apellido_Paterno LIKE '%" & Trim(Txt_Calendario_Filtro_Empleados.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Txt_Calendario_Filtro_Empleados.Text) & "%') AND Estatus = 'A'", Lst_Calendarios_Configuracion_Empleados, 0, "Cat_Empleados.No_Tarjeta")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Txt_Calendario_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Calendarios_Configuracion_Turno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Dias_No_Laborales_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Acronimo_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Ciudad_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Colonia_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_CP_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Direccion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Estado_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Empresas_Noi_Coi_ID_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Cat_Empresas_Noi_Coi_ID, False)
End Sub

Private Sub Txt_Cat_Empresas_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_RFC_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Empresas_Telefono_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Equipos_Descripcion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Equipos_Almacenes_Descripcion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Motivos_Baja_Descripcion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Motivos_Baja_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub
Private Sub Txt_Cat_Nivel_Estudio_Descripcion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Nivel_Estudio_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Puestos_Abreviatura_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Puestos_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Puestos_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Tipos_Faltas_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Tipos_Faltas_Descripcion_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Tipos_Faltas_Simbologia_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Turnos_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Cat_Turnos_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Txt_Clave_SAP_Departamentos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Clave_SAP_Motivos_Baja_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Clave_SAP_Puestos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Clave_SAP_Tipos_Faltas_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Departamento_Clave_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Departamento_Comentarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Departamento_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Valida_Login_Password_Usuario
    'DESCRIPCIÓN: Valida que no se repita algun Login ya existente cuando se agrega
    '             o se modifica algun usuario
    'PARÁMETROS: 1. Campo: Indica el campo que se va a comparar
    '            2. Valor: Indica el valor con el que se va a comparar el campo
    '            3. Cat_Usuario_ID: Si tiene algun valor, es porque se va a comprarar cuando se haga alguna modificacion
    'CREO: Susana Ledesma Ramírez
    'FECHA_CREO: 26/Abril/2006
    'MODIFICO:
    'FECHA_MODIFICO: 25/Octubre/2007
    'CAUSA_MODIFICACIÓN: Porque se necesitaba validar tambien el password
'*******************************************************************************

Public Function Valida_Login_Password_Usuario(Campo As String, Valor As String, Optional Cat_Usuario_ID As String) As Boolean
Dim Rs_Consulta_Cat_Usuarios As rdoResultset    'Maneja el registro de la Tabla de Cat_Usuarios

Set Conectar_Ayudante = New Ayudante
'Establece la consulta en Cat_Usuarios para saber el Login o el Password ya existen
Mi_SQL = "SELECT Login FROM Apl_Cat_Usuarios WHERE " & Campo & " = '" & Valor & "'"
'Si es alguna modificaciones, entonces se busca en todos los usuarios, excepto en el actual
If Cat_Usuario_ID <> "" Then Mi_SQL = Mi_SQL & " AND Usuario_ID<>'" & Cat_Usuario_ID & "'"
Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Si encuantra algun dato; entonces ese login o password ya existe
If Not Rs_Consulta_Cat_Usuarios.EOF Then
    Valida_Login_Password_Usuario = True 'Si el Login o password ya existe
End If
Rs_Consulta_Cat_Usuarios.Close
End Function

'************************************************Inicio Empresas***************************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Empresas
    'DESCRIPCIÓN:           Consulta las Empresas y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la Empresa
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           15 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Empresas(Nombre As String)
Dim Rs_Consulta_Cat_Empresas As rdoResultset       'Informacion de los registros
    
    Grid_Cat_Empresas.Rows = 0
    
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Empresa_ID, Nombre, Acronimo, Logo"
    Mi_SQL = Mi_SQL & " FROM Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Empresas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    With Rs_Consulta_Cat_Empresas
        If Not .EOF Then
            
            Grid_Cat_Empresas.AddItem "Empresa ID" & Chr(9) & "Nombre" & Chr(9) & "Acronimo" & Chr(9) & "Logo"
            While Not .EOF
                Grid_Cat_Empresas.AddItem .rdoColumns("Empresa_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Acronimo") & Chr(9) & .rdoColumns("Logo")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del grid_usuarios
            Grid_Cat_Empresas.FixedRows = 1
            Grid_Cat_Empresas.ColWidth(0) = 0      'Empresa_ID
            Grid_Cat_Empresas.ColWidth(1) = 5000    'Nombre
            Grid_Cat_Empresas.ColWidth(2) = 1200   'Acronimo
            Grid_Cat_Empresas.ColWidth(3) = 1500
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Empresas = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Empresas
    'DESCRIPCIÓN:           Da de alta un registro en Cat_Empresas
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           15 Mayo 2009
    'MODIFICO          : Flores Ramirez Yazmin
    'FECHA_MODIFICO    :07 Diciembre 2016
    'CAUSA_MODIFICACIÓN:Agregar equipos de identificacion de almacenes
'*******************************************************************************
Private Sub Alta_Cat_Empresas()
Guardar_Imagen_Logo
Dim Rs_Alta_Cat_Empresas As rdoResultset            'Informacion del registro
Dim Rs_Alta_Cat_Empresas_Checadores As rdoResultset 'Informacion de los checadores
Dim Rs_Alta_Cat_Empresas_Checadores_Almacenes As rdoResultset 'Informacion de los checadores Almacenes
Dim Cont_Fila As Integer
Dim Cont_Fila_Almacenes As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Set Rs_Alta_Cat_Empresas = Conectar_Ayudante.Recordset_Agregar("Cat_Empresas")
    'Agrega el reigstro del Empresa
    With Rs_Alta_Cat_Empresas
        .AddNew
            Txt_Cat_Empresas_Empresa_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empresas", "Empresa_ID"), "00000")
            .rdoColumns("Empresa_ID") = Trim(Txt_Cat_Empresas_Empresa_ID.Text)
            .rdoColumns("Acronimo") = Trim(Txt_Cat_Empresas_Acronimo.Text)
            .rdoColumns("Nombre") = Trim(Txt_Cat_Empresas_Nombre.Text)
            .rdoColumns("RFC") = Trim(Txt_Cat_Empresas_RFC.Text)
            .rdoColumns("Direccion") = Trim(Txt_Cat_Empresas_Direccion.Text)
            .rdoColumns("Colonia") = Trim(Txt_Cat_Empresas_Colonia.Text)
            .rdoColumns("Ciudad") = Trim(Txt_Cat_Empresas_Ciudad.Text)
            .rdoColumns("Estado") = Trim(Txt_Cat_Empresas_Estado.Text)
            .rdoColumns("Codigo_Postal") = Trim(Txt_Cat_Empresas_CP.Text)
            .rdoColumns("Telefono") = Trim(Txt_Cat_Empresas_Telefono.Text)
            .rdoColumns("Logo") = Txt_Logo.Text
            .rdoColumns("Tipo_Nomina") = Trim(Cmb_Cat_Empresas_Tipo_Nomina.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Empresas_Comentarios.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Cat_Empresas = Nothing
    'Agrega los checadores
    Set Rs_Alta_Cat_Empresas_Checadores = Conectar_Ayudante.Recordset_Agregar("Cat_Empresas_Equipos_Identificacion")
    With Rs_Alta_Cat_Empresas_Checadores
        For Cont_Fila = 1 To Grid_Empresas_Equipos.Rows - 1
            .AddNew
                .rdoColumns("Empresa_ID") = Trim(Txt_Cat_Empresas_Empresa_ID.Text)
                .rdoColumns("Equipo_ID") = Trim(Grid_Empresas_Equipos.TextMatrix(Cont_Fila, 0))
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empresas_Checadores = Nothing
    
    'Agrega los checadores Almacenes
    Set Rs_Alta_Cat_Empresas_Checadores_Almacenes = Conectar_Ayudante.Recordset_Agregar("Cat_Empresas_Equipos_Identificacion_Almacenes")
    With Rs_Alta_Cat_Empresas_Checadores_Almacenes
        For Cont_Fila_Almacenes = 1 To Grid_Empresas_Equipos_Almacenes.Rows - 1
            .AddNew
                .rdoColumns("Empresa_ID") = Trim(Txt_Cat_Empresas_Empresa_ID.Text)
                .rdoColumns("Equipo_ID") = Trim(Grid_Empresas_Equipos_Almacenes.TextMatrix(Cont_Fila_Almacenes, 0))
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empresas_Checadores_Almacenes = Nothing
    
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    
    Fra_Cat_Empresas_Datos_Generales.Enabled = False
    Fra_Cat_Empresas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    With Grid_Cat_Empresas
        If .Rows = 0 Then
            .AddItem "Empresa ID" & Chr(9) & "Nombre" & Chr(9) & "Acronimo"
        End If
        'Llena el grid con los datos del nuevo usuario
        .AddItem Trim(Txt_Cat_Empresas_Empresa_ID.Text) & Chr(9) & Trim(Txt_Cat_Empresas_Nombre.Text) & Chr(9) & Trim(Txt_Cat_Empresas_Acronimo.Text)
        
        'Configura el tamaño de las columnas del grid_usuarios
        .FixedRows = 1
        .ColWidth(0) = 0      'Empresa_ID
        .ColWidth(1) = 6000   'Nombre
        .ColWidth(2) = 1800   'acronimo

    End With
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Me)
    MsgBox "Empresa dada de alta", vbOKOnly + vbInformation, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Empresas
    'DESCRIPCIÓN:           Modifica el registro de la Empresa
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO        :    15 Mayo 2009
    'MODIFICO          : Flores Ramirez Yazmin
    'FECHA_MODIFICO    : 07 Diciembre 2016
    'CAUSA_MODIFICACIÓN: Agregar Registro de Equipos de indentificacion Almacenes
'*******************************************************************************
Private Sub Modifica_Cat_Empresas()
Dim Rs_Modificacion_Cat_Empresas As rdoResultset 'Informacion del registro
Dim Rs_Alta_Cat_Empresas_Checadores As rdoResultset
Dim Rs_Alta_Cat_Empresas_Checadores_Almacenes As rdoResultset
Guardar_Imagen_Logo
Dim Cont_Fila As Integer
Dim Cont_Fila_Almacenes As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Empresa_ID ='" & Trim(Txt_Cat_Empresas_Empresa_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Empresas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    
   
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Empresas
        .Edit
            .rdoColumns("Acronimo") = Trim(Txt_Cat_Empresas_Acronimo.Text)
            .rdoColumns("Nombre") = Trim(Txt_Cat_Empresas_Nombre.Text)
            .rdoColumns("RFC") = Trim(Txt_Cat_Empresas_RFC.Text)
            .rdoColumns("Direccion") = Trim(Txt_Cat_Empresas_Direccion.Text)
            .rdoColumns("Colonia") = Trim(Txt_Cat_Empresas_Colonia.Text)
            .rdoColumns("Ciudad") = Trim(Txt_Cat_Empresas_Ciudad.Text)
            .rdoColumns("Estado") = Trim(Txt_Cat_Empresas_Estado.Text)
            .rdoColumns("Codigo_Postal") = Trim(Txt_Cat_Empresas_CP.Text)
            .rdoColumns("Telefono") = Trim(Txt_Cat_Empresas_Telefono.Text)
            .rdoColumns("Tipo_Nomina") = Trim(Cmb_Cat_Empresas_Tipo_Nomina.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Empresas_Comentarios.Text))
            .rdoColumns("Logo") = Txt_Logo.Text
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Empresas = Nothing
    
    'Agrega los checadores
    Mi_SQL = "DELETE Cat_Empresas_Equipos_Identificacion WHERE Empresa_ID = '" & Trim(Txt_Cat_Empresas_Empresa_ID.Text) & "'"
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empresas_Checadores = Conectar_Ayudante.Recordset_Agregar("Cat_Empresas_Equipos_Identificacion")
    With Rs_Alta_Cat_Empresas_Checadores
        For Cont_Fila = 1 To Grid_Empresas_Equipos.Rows - 1
            .AddNew
                .rdoColumns("Empresa_ID") = Trim(Txt_Cat_Empresas_Empresa_ID.Text)
                .rdoColumns("Equipo_ID") = Trim(Grid_Empresas_Equipos.TextMatrix(Cont_Fila, 0))
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empresas_Checadores = Nothing
    
    'Agrega los checadores Almacenes
    Mi_SQL = "DELETE Cat_Empresas_Equipos_Identificacion_Almacenes WHERE Empresa_ID = '" & Trim(Txt_Cat_Empresas_Empresa_ID.Text) & "'"
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Empresas_Checadores_Almacenes = Conectar_Ayudante.Recordset_Agregar("Cat_Empresas_Equipos_Identificacion_Almacenes")
    With Rs_Alta_Cat_Empresas_Checadores_Almacenes
        For Cont_Fila_Almacenes = 1 To Grid_Empresas_Equipos_Almacenes.Rows - 1
            .AddNew
                .rdoColumns("Empresa_ID") = Trim(Txt_Cat_Empresas_Empresa_ID.Text)
                .rdoColumns("Equipo_ID") = Trim(Grid_Empresas_Equipos_Almacenes.TextMatrix(Cont_Fila_Almacenes, 0))
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Empresas_Checadores_Almacenes = Nothing
    
    With Grid_Cat_Empresas
        .TextMatrix(.RowSel, 1) = Trim(Txt_Cat_Empresas_Nombre.Text)
        .TextMatrix(.RowSel, 2) = Trim(Txt_Cat_Empresas_Acronimo.Text)
    End With
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Cat_Empresas_Datos_Generales.Enabled = False
    Fra_Cat_Empresas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Me)
    MsgBox "La Empresa ha sido modificada", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'************************************************Termino Empresas***************************************

'************************************************Inicio Dias No Laborales***************************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Dias_No_Laborales
    'DESCRIPCIÓN:           Consulta los dias no laborales registrados
    'PARÁMETROS :           Nombre: nombre del turno a buscar
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           05 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Consulta_Cat_Dias_No_Laborales(Nombre As String)
Dim Rs_Consulta_Cat_Dias_No_Laborales As rdoResultset     'Informacion de los Maquinas

Grid_Cat_Dias_No_Laborales.Rows = 0
'Consulta todos los roles que se encuentran dados de alta
Mi_SQL = "SELECT Dia_No_Laboral_ID, Fecha, Comentarios FROM Cat_Dias_No_Laborales"
Mi_SQL = Mi_SQL & " WHERE Comentarios LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY Fecha"
Set Rs_Consulta_Cat_Dias_No_Laborales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Dias_No_Laborales
    If Not .EOF Then
        Grid_Cat_Dias_No_Laborales.AddItem "Dia_No_Laboral_ID" & Chr(9) & "Fecha" & Chr(9) & "Comentarios"
        While Not .EOF
            Grid_Cat_Dias_No_Laborales.AddItem .rdoColumns("Dia_No_Laboral_ID") & Chr(9) & .rdoColumns("Fecha") & Chr(9) & _
                                .rdoColumns("Comentarios")
            .MoveNext
        Wend
        'Asigna los tamaños de las columnas del grid_roles
        Grid_Cat_Dias_No_Laborales.FixedRows = 1
        Grid_Cat_Dias_No_Laborales.ColWidth(0) = 0     'Dia_No_Laboral_ID
        Grid_Cat_Dias_No_Laborales.ColWidth(1) = 2000  'Fecha
        Grid_Cat_Dias_No_Laborales.ColWidth(2) = 5000  'Comentarios
    End If
    .Close
End With

Set Rs_Consulta_Cat_Dias_No_Laborales = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Dias_No_Laborales
    'DESCRIPCIÓN:           Realiza el Alta de un Dia no laboral en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           15 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Alta_Cat_Dias_No_Laborales()
Dim Rs_Alta_Cat_Dias_No_Laborales As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Set Rs_Alta_Cat_Dias_No_Laborales = Conectar_Ayudante.Recordset_Agregar("Cat_Dias_No_Laborales")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Dias_No_Laborales
        .AddNew
            Txt_Cat_Dias_No_Laborales_Dia_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Dias_No_Laborales", "Dia_No_Laboral_ID"), "00000")
            .rdoColumns("Dia_No_Laboral_ID") = Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text)
            .rdoColumns("Fecha") = Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy")
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Dias_No_Laborales_Comentarios.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Cat_Dias_No_Laborales = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    'Pone un encabezado en el grid
    With Grid_Cat_Dias_No_Laborales
        If .Rows = 0 Then
            .AddItem "Dia_No_Laboral_ID" & Chr(9) & "Fecha" & Chr(9) & "Comentarios"
        End If
        'Llena el grid con los datos del nuevo Maquina
        .AddItem Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text) & Chr(9) & Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy") & Chr(9) & _
                 Trim(Txt_Cat_Dias_No_Laborales_Comentarios.Text)
        'Asigna los tamaños de las columnas del grid_roles
        .FixedRows = 1
        .ColWidth(0) = 0     'Dia_No_Laboral_ID
        .ColWidth(1) = 2000  'Fecha
        .ColWidth(2) = 5000  'Comentarios
    End With
    Fra_Cat_Dias_No_Laborales_Generales.Enabled = False
    Fra_Cat_Dias_No_Laborales.Enabled = True
    Dtp_Cat_Dias_No_Laborales_Fecha.Value = Now
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales", Me)
    MsgBox "Dia dado de alta", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Dias_No_Laborales
    'DESCRIPCIÓN:           Realiza la modificacion del dia no laboral en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           15 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Modifica_Cat_Dias_No_Laborales()
Dim Rs_Modifica_Cat_Dias_No_Laborales As rdoResultset 'Informacion del Maquinas
Dim Mi_SQL As String
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Dias_No_Laborales"
    Mi_SQL = Mi_SQL & " WHERE Dia_No_Laboral_ID = '" & Trim(Txt_Cat_Dias_No_Laborales_Dia_ID.Text) & "'"
    
    'Modifica Maquina
    Set Rs_Modifica_Cat_Dias_No_Laborales = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Modifica_Cat_Dias_No_Laborales
        .Edit
            .rdoColumns("Fecha") = Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy")
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Dias_No_Laborales_Comentarios.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    
    Set Rs_Modifica_Cat_Dias_No_Laborales = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    'Modifica la informacion en el grid
    With Grid_Cat_Dias_No_Laborales
        .TextMatrix(.RowSel, 1) = Format(Dtp_Cat_Dias_No_Laborales_Fecha.Value, "MM/dd/yyyy")
        .TextMatrix(.RowSel, 2) = Trim(Txt_Cat_Dias_No_Laborales_Comentarios.Text)
    End With
    Fra_Cat_Dias_No_Laborales_Generales.Enabled = False
    Fra_Cat_Dias_No_Laborales.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Dtp_Cat_Dias_No_Laborales_Fecha.Value = Now
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales", Me)
    MsgBox "Dia no laboral Modificado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub
'************************************************Termino Dias No Laborales***************************************

'************************************************Inicio Faltas***************************************
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Tipos_Faltas
'DESCRIPCION: Consulta los tipos de incidencias que existen en el catálogo
'PARAMETROS : Nombre: nombre de la incidencia a buscar
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Consulta_Cat_Tipos_Faltas(Nombre As String)
Dim Rs_Consulta_Cat_Tipos_Faltas As rdoResultset     'Informacion de los Maquinas

    Grid_Cat_Tipos_Faltas.Rows = 0
    Mi_SQL = "SELECT * FROM Cat_Tipos_Faltas"
    Mi_SQL = Mi_SQL & " WHERE Descripcion LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Descripcion"
    Set Rs_Consulta_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Tipos_Faltas
        If Not .EOF Then
            Grid_Cat_Tipos_Faltas.AddItem "ID" & Chr(9) & "Descripcion" & Chr(9) & "Simbolo" & Chr(9) & "Codigo" & Chr(9) & "Comentarios" & Chr(9) & "ClaveSAP"
            While Not .EOF
                Grid_Cat_Tipos_Faltas.AddItem .rdoColumns("Tipo_Falta_ID") & Chr(9) & .rdoColumns("Descripcion") _
                    & Chr(9) & .rdoColumns("Simbologia") & Chr(9) & .rdoColumns("Codigo_NOI") & Chr(9) & .rdoColumns("Comentarios") & Chr(9) & .rdoColumns("Clave_SAP")
                .MoveNext
            Wend
            'Asigna los tamaños de las columnas del grid_roles
            With Grid_Cat_Tipos_Faltas
                .FixedRows = 1
                .ColWidth(0) = 1000 'Tipo_Falta_ID
                .ColWidth(1) = 4000 'Descripcion
                .ColAlignment(1) = flexAlignLeftCenter
                .ColWidth(2) = 1200 'Simbolo
                .ColAlignment(2) = flexAlignLeftCenter
                .ColWidth(3) = 0    'NOI
                .ColWidth(4) = 0    'Comentarios
                .ColWidth(5) = 0    'Clave_SAP
                '.ColAlignment(5) = flexAlignLeftCenter
            End With
        End If
        .Close
    End With
    Set Rs_Consulta_Cat_Tipos_Faltas = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Tipos_Faltas
    'DESCRIPCIÓN:           Realiza el Alta de un Dia no laboral en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           15 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Alta_Cat_Tipos_Faltas()
Dim Rs_Alta_Cat_Tipos_Faltas As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Set Rs_Alta_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Agregar("Cat_Tipos_Faltas")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Tipos_Faltas
        .AddNew
            Txt_Cat_Tipos_Faltas_Falta_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Faltas", "Tipo_Falta_ID"), "00000")
            .rdoColumns("Tipo_Falta_ID") = Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text)
            .rdoColumns("Descripcion") = Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text)
            .rdoColumns("Simbologia") = Trim(Txt_Cat_Tipos_Faltas_Simbologia.Text)
            .rdoColumns("Comentarios") = Trim(Txt_Cat_Tipos_Faltas_Comentarios.Text)
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Tipos_Faltas.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Cat_Tipos_Faltas = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Tipos_Faltas_Generales.Enabled = False
    Fra_Cat_Tipos_Faltas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    With Grid_Cat_Tipos_Faltas
        If .Rows = 0 Then
            .AddItem "ID" & Chr(9) & "Descripcion" & Chr(9) & "Simbolo" & Chr(9) & "Codigo" & Chr(9) & "Comentarios"
        End If
        'Llena el grid con los datos del nuevo Maquina
        .AddItem Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text) & Chr(9) & Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text) & Chr(9) & _
                 Trim(Txt_Cat_Tipos_Faltas_Simbologia.Text) & Chr(9) & "" & Chr(9) & _
                 Trim(Txt_Cat_Tipos_Faltas_Comentarios.Text)
        'Asigna los tamaños de las columnas del grid_roles
        .FixedRows = 1
        .ColWidth(0) = 1000     'Tipo_Falta_ID
        .ColWidth(1) = 4000  'Descripcion
        .ColAlignment(2) = flexAlignRightCenter
        .ColWidth(2) = 900  'Simbolo
        .ColAlignment(3) = flexAlignRightCenter
        .ColWidth(3) = 1000  'NOI
        .ColAlignment(4) = flexAlignLeftCenter
        .ColWidth(4) = 0  'Comentarios
        .ColWidth(5) = 0  '
    End With
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias", Me)
    MsgBox "Incidencia Extraordinaria dada de alta", vbInformation + vbOKOnly, Me.Caption
Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Cat_Tipos_Faltas
'DESCRIPCION: Realiza la modificacion de la incidencia en la base de datos
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Cat_Tipos_Faltas()
Dim Rs_Modifica_Cat_Tipos_Faltas As rdoResultset 'Informacion del Maquinas
Dim Mi_SQL As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Tipos_Faltas"
    Mi_SQL = Mi_SQL & " WHERE Tipo_Falta_ID = '" & Trim(Txt_Cat_Tipos_Faltas_Falta_ID.Text) & "'"
    'Modifica Maquina
    Set Rs_Modifica_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Modifica_Cat_Tipos_Faltas
        .Edit
            .rdoColumns("Clasificacion") = Cmb_Clasificacion_Incidencias.Text
            .rdoColumns("Descripcion") = Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text)
            .rdoColumns("Simbologia") = Trim(Txt_Cat_Tipos_Faltas_Simbologia.Text)
            .rdoColumns("Comentarios") = Trim(Txt_Cat_Tipos_Faltas_Comentarios.Text)
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Tipos_Faltas.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Tipos_Faltas = Nothing
    'Modifica la informacion en el grid
    With Grid_Cat_Tipos_Faltas
        .TextMatrix(.RowSel, 1) = Trim(Txt_Cat_Tipos_Faltas_Descripcion.Text)
        .TextMatrix(.RowSel, 2) = Trim(Txt_Cat_Tipos_Faltas_Simbologia.Text)
        .TextMatrix(.RowSel, 3) = ""
        .TextMatrix(.RowSel, 4) = Trim(Txt_Cat_Tipos_Faltas_Comentarios.Text)
        .TextMatrix(.RowSel, 5) = Trim(Txt_Clave_SAP_Tipos_Faltas.Text)
    End With
    Fra_Cat_Tipos_Faltas_Generales.Enabled = False
    Fra_Cat_Tipos_Faltas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias", Me)
    MsgBox "Incidencia Extraordinaria Modificada", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub
'************************************************Termino Faltas***************************************

'´***********************************************Inicio Departamentos*********************************
'*******************************************************************************
'NOMBRE_FUNCION:  Consulta_Departamentos
'DESCRIPCION: Consulta los departamentos de la base de datos
'PARAMETROS : Nombre- Nombre del departamento a buscar
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************

Private Sub Consulta_Departamentos(Nombre As String)
Dim Rs_Consulta_Cat_Departamentos As rdoResultset     'Informacion de los Maquinas

    Grid_Departamentos.Rows = 0
    Grid_Departamentos.Cols = 5
    'Consulta todos los roles que se encuentran dados de alta
    Mi_SQL = "SELECT * FROM Cat_Departamentos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Departamentos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Departamentos
        If Not .EOF Then
            Grid_Departamentos.AddItem "Departamento ID" & Chr(9) & "Nombre" & Chr(9) & "Clave" & Chr(9) & "Comentarios" & Chr(9) & "ClaveSAP"
            While Not .EOF
                Grid_Departamentos.AddItem .rdoColumns("Departamento_ID") & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Clave") & Chr(9) & .rdoColumns("Comentarios") & Chr(9) & .rdoColumns("Clave_SAP")
                .MoveNext
            Wend
            'Asigna los tamaños de las columnas del grid_roles
            Grid_Departamentos.FixedRows = 1
            Grid_Departamentos.ColWidth(0) = 1550   'Departamento ID
            Grid_Departamentos.ColWidth(1) = 4500   'Nombre
            Grid_Departamentos.ColWidth(2) = 1500   'Clave
            Grid_Departamentos.ColWidth(3) = 0      'Comentarios
            Grid_Departamentos.ColWidth(4) = 0      'ClaveSAP
        End If
        .Close
    End With
    Set Rs_Consulta_Cat_Departamentos = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Departamento
    'DESCRIPCIÓN:           Realiza el Alta de un departamento en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           04 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Alta_Departamento()
Dim Rs_Alta_Cat_Departamentos As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Set Rs_Alta_Cat_Departamentos = Conectar_Ayudante.Recordset_Agregar("Cat_Departamentos")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Departamentos
        .AddNew
            Txt_Departamento_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Departamentos", "Departamento_ID"), "00000")
            .rdoColumns("Departamento_ID") = Trim(Txt_Departamento_ID.Text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Departamento_Nombre.Text))
            .rdoColumns("Clave") = Trim(UCase(Txt_Departamento_Clave.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Departamento_Comentarios.Text))
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Departamentos.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    
    Set Rs_Alta_Cat_Departamentos = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Departamento_Generales.Enabled = False
    Fra_Departamentos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Departamentos.Rows = 0 Then
        Grid_Departamentos.Cols = 4
        Grid_Departamentos.AddItem "Departamento ID" & Chr(9) & "Nombre" & Chr(9) & "Clave" & Chr(9) & "Comentarios"
    End If
    'Llena el grid con los datos del nuevo Departamento
    Grid_Departamentos.AddItem Trim(Txt_Departamento_ID.Text) & Chr(9) & UCase(Txt_Departamento_Nombre.Text) & Chr(9) & _
                               Trim(Txt_Departamento_Clave.Text) & Chr(9) & Trim(Txt_Departamento_Comentarios.Text)
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Departamentos.FixedRows = 1
    Grid_Departamentos.ColWidth(0) = 1550    'Departamento ID
    Grid_Departamentos.ColWidth(1) = 4500    'Nombre
    Grid_Departamentos.ColWidth(2) = 1500    'Clave
    Grid_Departamentos.ColWidth(3) = 0     'Comentarios
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos", Me)
    MsgBox "Departamento dado de alta", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
    Debug.Print Err.Description
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Departamento
'DESCRIPCION: Realiza la modificacion de un departamento en la base de datos
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Departamento()
Dim Rs_Modifica_Cat_Departamentos As rdoResultset 'Informacion del Maquinas
Dim Mi_SQL As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Departamentos"
    Mi_SQL = Mi_SQL & " WHERE Departamento_ID = '" & Trim(Txt_Departamento_ID.Text) & "'"
    Set Rs_Modifica_Cat_Departamentos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Modifica_Cat_Departamentos
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Departamento_Nombre.Text))
            .rdoColumns("Clave") = Trim(UCase(Txt_Departamento_Clave.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Departamento_Comentarios.Text))
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Departamentos.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Departamentos = Nothing
    Conexion_Base.CommitTrans
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir o modificar los valoes
    Fra_Departamento_Generales.Enabled = False
    Fra_Departamentos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    'Modifica la informacion en el grid
    Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 1) = Trim(Txt_Departamento_Nombre.Text)
    Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 2) = Trim(Txt_Departamento_Clave.Text)
    Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 3) = Trim(Txt_Departamento_Comentarios.Text)
    Grid_Departamentos.TextMatrix(Grid_Departamentos.RowSel, 4) = Trim(Txt_Clave_SAP_Departamentos.Text)
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos", Me)
    MsgBox "Departamento Modificado", vbInformation + vbOKOnly, Me.Caption
Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*****************************************Termino Departamentos**********************************

Public Sub Inicializa()
    Select Case Catalogo
        Case "Cat_Empresas"
            Consulta_Cat_Empresas ""
            'Llena los checadores
            Call Conectar_Ayudante.Llena_Combo_Item("Equipo_ID, cast(No_Equipo as varchar)+' - '+Descripcion", "Cat_Equipos_Identificadores", Cmb_Cat_Empresas_Equipo, 0, "No_Equipo")
            Call Conectar_Ayudante.Llena_Combo_Item("Equipo_ID, cast(No_Equipo as varchar)+' - '+Descripcion", "Cat_Equipos_Almacenes_Identificadores", Cmb_Cat_Empresas_Equipo_Almacenes, 0, "No_Equipo")
        
        Case "Cat_Turnos":
            Consulta_Cat_Turnos ""
            Dtp_Cat_Turnos_Hora_Inicio = "12:00"
            Dtp_Cat_Turnos_Hora_Termino = "12:00"
            
        Case "Cat_Calendarios_Turnos":
            Me.Width = 11605
            Btn_Modificar.Left = Btn_Modificar.Left * 1.4
            Btn_Eliminar.Left = Btn_Eliminar.Left * 1.4
            Btn_Consultar.Left = Btn_Consultar.Left * 1.4
            Btn_Salir.Left = Btn_Salir.Left * 1.4
'            Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, (Turno_ID COLLATE DATABASE_DEFAULT + ' ' + Nombre COLLATE DATABASE_DEFAULT) AS Nombre", "Cat_Turnos", Cmb_Calendario_Turnos, 0, "Turno_ID")
'            Cmb_Calendario_Turnos.ListIndex = -1
            Call Conectar_Ayudante.Llena_List_Item("Cat_Empleados.No_Tarjeta, CAST(Cat_Empleados.No_Tarjeta AS VARCHAR)+' - '+Cat_Empleados.Nombre + ' ' + Cat_Empleados.Apellido_Paterno + ' ' + Cat_Empleados.Apellido_Materno", "Cat_Empleados, Cat_Areas_Detalles, Cat_Usuarios WHERE Cat_Empleados.Estatus = 'A' AND Cat_Usuarios.Usuario_ID = '" & Usuario_ID & "' AND Cat_Usuarios.Rol_ID = 4 AND Cat_Empleados.Empleado_ID = Cat_Areas_Detalles.Empleado_ID AND Cat_Areas_Detalles.Area_ID = Cat_Usuarios.Area_ID", Lst_Calendarios_Configuracion_Empleados, 0, "Cat_Empleados.No_Tarjeta")
            Consulta_Cat_Calendarios_Turnos ""
            Dtp_Calendario_Fecha_Inicio = Now
            Dtp_Calendario_Fecha_Termino = Now
            Dtp_Cat_Turnos_Hora_Inicio = "00:00"
            Dtp_Cat_Turnos_Hora_Termino = "00:00"
            
        Case "Cat_Dias_No_Laborales":
            Consulta_Cat_Dias_No_Laborales ""
            Dtp_Cat_Dias_No_Laborales_Fecha.Value = Now
        
        Case "Cat_Tipos_Faltas":
            Consulta_Cat_Tipos_Faltas ""
            
        Case "Cat_Departamentos": 'Catalogo de Departamentos
            Consulta_Departamentos ""
        
        Case "Cat_Puestos"
            Consulta_Cat_Puestos ""
            
        Case "Cat_Equipos_Identificacion":
            Consulta_Cat_Equipos_Identificadores ""
        
        Case "Cat_Nivel_Estudio"
            Consulta_Cat_Nivel_Estudio ""
        
        Case "Cat_Motivos_Baja"
            Consulta_Cat_Motivos_Baja ""
            
        Case "Cat_Equipos_Almacenes_Identificacion":
            Consulta_Cat_Equipos_Almacenes_Identificadores ""
    End Select
End Sub

'*********************************Inicio Puestos***************************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Cat_Puestos
    'DESCRIPCIÓN:          Alta de los puestos
    'PARÁMETROS :
    'CREO       :          Yañez Rodriguez Diego Neftali
    'FECHA_CREO :          20 Octubre 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Cat_Puestos()
Dim Rs_Alta_Cat_Puestos As rdoResultset     'Informacion de la base de datos
Dim Rs_Alta_Cat_Puestos_Dominio As rdoResultset
Dim Cont_Filas As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Da de alta los valores introducidos por el usuario en la tabla Cat_Unidades_Medidas
    Set Rs_Alta_Cat_Puestos = Conectar_Ayudante.Recordset_Agregar("Cat_Puestos")
    With Rs_Alta_Cat_Puestos
        'Da de alta el registro de la unidad de medida en la tabla Cat_Unidades_Medidas
        .AddNew
            Txt_Cat_Puestos_Puesto_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Puestos", "Puesto_ID"), "00000")
            .rdoColumns("Puesto_ID") = Trim(Txt_Cat_Puestos_Puesto_ID.Text)
            .rdoColumns("Nombre") = Trim((Txt_Cat_Puestos_Nombre.Text))
            .rdoColumns("Abreviatura") = Trim((Txt_Cat_Puestos_Abreviatura.Text))
            .rdoColumns("Descripcion") = Trim((Txt_Cat_Puestos_Comentarios.Text))
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Puestos.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Puestos.Close
    Set Rs_Alta_Cat_Puestos = Nothing
    
    Conexion_Base.CommitTrans
    With Grid_Cat_Puestos
        If .Rows = 0 Then
            .AddItem "Puesto_ID" & Chr(9) & "Nombre" & Chr(9) & "Abreviatura" & Chr(9) & "Comentarios"
        End If
        .AddItem Trim(Txt_Cat_Puestos_Puesto_ID.Text) & Chr(9) & _
        Trim(Txt_Cat_Puestos_Nombre.Text) & Chr(9) & _
        Trim(Txt_Cat_Puestos_Abreviatura.Text) & Chr(9) & _
        Trim(Txt_Cat_Puestos_Comentarios.Text)
        
        'Configura el tamaño de las celdas del grid
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 1000     'ID
        .ColWidth(1) = 5500     'Nombre
        .ColWidth(2) = 1000     'Abreviatura
        .ColWidth(3) = 0        'Comentarios
    End With
    'Habilita y deshabilita los controles de la forma
    Fra_Cat_Puestos.Enabled = True
    Fra_Cat_Puestos_Generales.Enabled = False
    Fra_Cat_Puestos.Visible = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Puestos", Me)
    MsgBox "Puesto Registrado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Cat_Puestos
'DESCRIPCION: Modifica los puestos
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Cat_Puestos()
Dim Rs_Modificacion_Cat_Puestos As rdoResultset 'Modifica los valores de la unida de medida que selecciono el usuario
Dim Rs_Alta_Cat_Puestos_Dominio As rdoResultset

Dim Cont_Filas As Integer

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Informacion de la unidad de medida
    Mi_SQL = "SELECT * FROM Cat_Puestos"
    Mi_SQL = Mi_SQL & " WHERE Puesto_ID='" & Trim(Txt_Cat_Puestos_Puesto_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Puestos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Puestos.EOF Then
        'Actualiza los valores del regisro seleccionado
        With Rs_Modificacion_Cat_Puestos
            .Edit
                .rdoColumns("Nombre") = Trim((Txt_Cat_Puestos_Nombre.Text))
                .rdoColumns("Abreviatura") = Trim((Txt_Cat_Puestos_Abreviatura.Text))
                .rdoColumns("Descripcion") = Trim((Txt_Cat_Puestos_Comentarios.Text))
                .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Puestos.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Puestos.Close
    Set Rs_Modificacion_Cat_Puestos = Nothing
    Conexion_Base.CommitTrans
    Grid_Cat_Puestos.TextMatrix(Grid_Cat_Puestos.RowSel, 1) = Trim(UCase(Txt_Cat_Puestos_Nombre.Text))
    Grid_Cat_Puestos.TextMatrix(Grid_Cat_Puestos.RowSel, 2) = Trim((Txt_Cat_Puestos_Abreviatura.Text))
    Grid_Cat_Puestos.TextMatrix(Grid_Cat_Puestos.RowSel, 3) = Trim((Txt_Cat_Puestos_Comentarios.Text))
    Grid_Cat_Puestos.TextMatrix(Grid_Cat_Puestos.RowSel, 4) = Trim(Txt_Clave_SAP_Puestos.Text)
    Fra_Cat_Puestos_Generales.Enabled = False
    Fra_Cat_Puestos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Puestos", Me)
    MsgBox "Registro Actualizado correctamente", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Puestos
'DESCRIPCIÓN: Consulta la informacion de los puestos
'PARAMETROS : Nombre: Puesto a buscar
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Cat_Puestos(Nombre As String)
Dim Rs_Consulta_Cat_Puestos As rdoResultset    'Informacion de las unidades de medida

On Error GoTo HANDLER
    Grid_Cat_Puestos.Rows = 0
    'Consulta todos los registros dados de alta en la tabla Cat_Unidades_Medidas
    Mi_SQL = "SELECT * FROM Cat_Puestos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Set Rs_Consulta_Cat_Puestos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si se encontraron valores entonces los agrega al Grid_Unidades_Medidas
    With Rs_Consulta_Cat_Puestos
        If Not .EOF Then
            Grid_Cat_Puestos.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Abreviatura" & Chr(9) & "Comentarios" & Chr(9) & "ClaveSAP"
            'Llena el Grid_Unidades_Medidas con los valores traidos de la consulta
            While Not .EOF
                Grid_Cat_Puestos.AddItem .rdoColumns("Puesto_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Abreviatura") _
                    & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Clave_SAP")
                .MoveNext
            Wend
            'Configura el Grid_Unidades_Medidas
            Grid_Cat_Puestos.FixedCols = 1
            Grid_Cat_Puestos.FixedRows = 1
            Grid_Cat_Puestos.ColWidth(0) = 1000     'ID
            Grid_Cat_Puestos.ColWidth(1) = 5500     'Puesto
            Grid_Cat_Puestos.ColWidth(2) = 1000     'Abreviatura
            Grid_Cat_Puestos.ColWidth(3) = 0        'Comentarios
            Grid_Cat_Puestos.ColWidth(4) = 0        'Clave SAP
            .Close
        End If
    End With
    Set Rs_Consulta_Cat_Puestos = Nothing
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*********************************termino Puestos ***************************************

'´***********************************************Inicio Equipos Identificacion*********************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Equipos_Identificadores
    'DESCRIPCIÓN:           Consulta los equipos de la base de datos
    'PARÁMETROS :           Nombre: numero  del equipo a buscar
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           05 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************

Private Sub Consulta_Cat_Equipos_Identificadores(Nombre As String)
Dim Rs_Consulta_Cat_Equipos_Identificadores As rdoResultset     'Informacion de los Maquinas

Grid_Cat_Equipos.Rows = 0
Grid_Cat_Equipos.Cols = 5
'Consulta todos los roles que se encuentran dados de alta
Mi_SQL = "SELECT Equipo_ID, No_Equipo, Direccion_IP, Puerto_IP, Descripcion "
Mi_SQL = Mi_SQL & " FROM Cat_Equipos_Identificadores"
Mi_SQL = Mi_SQL & " WHERE No_Equipo LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY No_Equipo"
Set Rs_Consulta_Cat_Equipos_Identificadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Equipos_Identificadores
    If Not .EOF Then
        Grid_Cat_Equipos.AddItem "Equipo ID" & Chr(9) & "No Equipo" & Chr(9) & "Direccion IP" & Chr(9) & "Puerto_IP" & Chr(9) & "Descripcion"
        While Not .EOF
            Grid_Cat_Equipos.AddItem .rdoColumns("Equipo_ID") & Chr(9) & .rdoColumns("No_Equipo") & Chr(9) & _
                                      .rdoColumns("Direccion_IP") & Chr(9) & .rdoColumns("Puerto_IP") & Chr(9) & _
                                      .rdoColumns("Descripcion")
            .MoveNext
        Wend
        'Asigna los tamaños de las columnas del grid_roles
        Grid_Cat_Equipos.FixedRows = 1
        Grid_Cat_Equipos.ColWidth(0) = 0    'Equipo ID
        Grid_Cat_Equipos.ColWidth(1) = 1500    'Numero equipo
        Grid_Cat_Equipos.ColWidth(2) = 2000    'Direccion IP
        Grid_Cat_Equipos.ColWidth(3) = 0     'Puerto IP
        Grid_Cat_Equipos.ColWidth(4) = 3200     'Descirpcion
    End If
    .Close
End With

Set Rs_Consulta_Cat_Equipos_Identificadores = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Equipos_Identificadores
    'DESCRIPCIÓN:           Realiza el Alta de un equipo en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           04 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Alta_Cat_Equipos_Identificadores()
Dim Rs_Alta_Cat_Equipos_Identificadores As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Set Rs_Alta_Cat_Equipos_Identificadores = Conectar_Ayudante.Recordset_Agregar("Cat_Equipos_Identificadores")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Equipos_Identificadores
        .AddNew
            Txt_Cat_Equipos_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Equipos_Identificadores", "Equipo_ID"), "00000")
            .rdoColumns("Equipo_ID") = Trim(Txt_Cat_Equipos_ID.Text)
            .rdoColumns("No_Equipo") = Val(Txt_Cat_Equipos_No_Equipo.Text)
            .rdoColumns("Direccion_IP") = Trim(UCase(Txt_Cat_Equipos_Direccion_IP.Text))
            .rdoColumns("Puerto_IP") = Val(Txt_Cat_Equipos_Puerto_IP.Text)
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Cat_Equipos_Descripcion.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    
    Set Rs_Alta_Cat_Equipos_Identificadores = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Equipos_Generales.Enabled = False
    Fra_Cat_Equipos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Equipos.Rows = 0 Then
        Grid_Cat_Equipos.Cols = 5
        Grid_Cat_Equipos.AddItem "Equipo ID" & Chr(9) & "No Equipo" & Chr(9) & "Direccion IP" & Chr(9) & "Puerto_IP" & Chr(9) & "Descripcion"
    End If
    'Llena el grid con los datos del nuevo Departamento
    Grid_Cat_Equipos.AddItem Trim(Txt_Cat_Equipos_ID.Text) & Chr(9) & Val(Txt_Cat_Equipos_No_Equipo.Text) & Chr(9) & _
                                      Trim(UCase(Txt_Cat_Equipos_Direccion_IP.Text)) & Chr(9) & Val(Txt_Cat_Equipos_Puerto_IP.Text) & Chr(9) & _
                                      Trim(UCase(Txt_Cat_Equipos_Descripcion.Text))
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Cat_Equipos.FixedRows = 1
    Grid_Cat_Equipos.ColWidth(0) = 0    'Equipo ID
    Grid_Cat_Equipos.ColWidth(1) = 1500    'Numero equipo
    Grid_Cat_Equipos.ColWidth(2) = 2000    'Direccion IP
    Grid_Cat_Equipos.ColWidth(3) = 0     'Puerto IP
    Grid_Cat_Equipos.ColWidth(4) = 3200     'Descirpcion
    Conexion_Base.CommitTrans
    Txt_Cat_Equipos_Direccion_IP.Text = ""
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion", Me)
    MsgBox "Equipo dado de alta", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
    Debug.Print Err.Description
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Equipos_Identificadores
    'DESCRIPCIÓN:           Realiza la modificacion de un equipo en la base de datos
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           04 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Modifica_Cat_Equipos_Identificadores()
Dim Rs_Modifica_Cat_Equipos_Identificadores As rdoResultset 'Informacion del Maquinas
Dim Mi_SQL As String
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Equipos_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Equipo_ID = '" & Trim(Txt_Cat_Equipos_ID.Text) & "'"
    
    'Modifica Maquina
    Set Rs_Modifica_Cat_Equipos_Identificadores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Modifica_Cat_Equipos_Identificadores
        .Edit
            .rdoColumns("No_Equipo") = Val(Txt_Cat_Equipos_No_Equipo.Text)
            .rdoColumns("Direccion_IP") = Trim(UCase(Txt_Cat_Equipos_Direccion_IP.Text))
            .rdoColumns("Puerto_IP") = Val(Txt_Cat_Equipos_Puerto_IP.Text)
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Cat_Equipos_Descripcion.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Equipos_Identificadores = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Equipos_Generales.Enabled = False
    Fra_Cat_Equipos.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    'Modifica la informacion en el grid
    Grid_Cat_Equipos.TextMatrix(Grid_Cat_Equipos.RowSel, 1) = Val(Txt_Cat_Equipos_No_Equipo.Text)
    Grid_Cat_Equipos.TextMatrix(Grid_Cat_Equipos.RowSel, 2) = Trim(UCase(Txt_Cat_Equipos_Direccion_IP.Text))
    Grid_Cat_Equipos.TextMatrix(Grid_Cat_Equipos.RowSel, 3) = Val(Txt_Cat_Equipos_Puerto_IP.Text)
    Grid_Cat_Equipos.TextMatrix(Grid_Cat_Equipos.RowSel, 4) = Trim(UCase(Txt_Cat_Equipos_Descripcion.Text))
    Conexion_Base.CommitTrans
    Txt_Cat_Equipos_Direccion_IP.Text = ""
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion", Me)
    MsgBox "Equipo Modificado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub
'*****************************************TErmino Equipos Identificacion**********************************


'*********************************Inicio Nivel Estudio***************************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Cat_Nivel_Estudio
    'DESCRIPCIÓN:          Alta de los niveles de estudio
    'PARÁMETROS :
    'CREO       :          Yañez Rodriguez Diego Neftali
    'FECHA_CREO :          17 Enero 2011
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Cat_Nivel_Estudio()
Dim Rs_Alta_Cat_Nivel_Estudio As rdoResultset     'Informacion de la base de datos

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Da de alta los valores introducidos por el usuario en la tabla Cat_Unidades_Medidas
    Set Rs_Alta_Cat_Nivel_Estudio = Conectar_Ayudante.Recordset_Agregar("Cat_Nivel_Estudio")
    With Rs_Alta_Cat_Nivel_Estudio
        'Da de alta el registro de la unidad de medida en la tabla Cat_Unidades_Medidas
        .AddNew
            Txt_Cat_Nivel_Estudio_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Nivel_Estudio", "Nivel_Estudio_ID"), "00000")
            .rdoColumns("Nivel_Estudio_ID") = Trim(Txt_Cat_Nivel_Estudio_ID.Text)
            .rdoColumns("Nombre") = Trim((Txt_Cat_Nivel_Estudio_Nombre.Text))
            .rdoColumns("Descripcion") = Trim((Txt_Cat_Nivel_Estudio_Descripcion.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Nivel_Estudio.Close
    Set Rs_Alta_Cat_Nivel_Estudio = Nothing
    
    Conexion_Base.CommitTrans
    With Grid_Cat_Nivel_Estudio
        If .Rows = 0 Then
            .AddItem "Nivel_Estudio_ID" & Chr(9) & "Nombre" & Chr(9) & "Descripcion"
        End If
        .AddItem Trim(Txt_Cat_Nivel_Estudio_ID.Text) & Chr(9) & _
        Trim(Txt_Cat_Nivel_Estudio_Nombre.Text) & Chr(9) & _
        Trim(Txt_Cat_Nivel_Estudio_Descripcion.Text)
        
        'Configura el tamaño de las celdas del grid
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 1000     'ID
        .ColWidth(1) = 5500     'Nombre
        .ColWidth(2) = 1000     'Abreviatura
    End With
    'Habilita y deshabilita los controles de la forma
    Fra_Cat_Nivel_Estudio.Enabled = True
    Fra_Cat_Nivel_Estudio_Generales.Enabled = False
    Fra_Cat_Nivel_Estudio.Visible = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio", Me)
    MsgBox "Nivel Estudio Registrado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Nivel_Estdio
    'DESCRIPCIÓN:           Modifica los niveles de estudio
    'PARÁMETROS :
    'CREO       :          Yañez Rodriguez Diego Neftali
    'FECHA_CREO :          18 Enero 2011
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Cat_Nivel_Estudio()
Dim Rs_Modificacion_Cat_Nivel_Estudio As rdoResultset 'Modifica los valores de la unida de medida que selecciono el usuario

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Informacion de la unidad de medida
    Mi_SQL = "SELECT * FROM Cat_Nivel_Estudio"
    Mi_SQL = Mi_SQL & " WHERE Nivel_Estudio_ID='" & Trim(Txt_Cat_Nivel_Estudio_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Nivel_Estudio = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Nivel_Estudio.EOF Then
        'Actualiza los valores del regisro seleccionado
        With Rs_Modificacion_Cat_Nivel_Estudio
            .Edit
                .rdoColumns("Nombre") = Trim((Txt_Cat_Nivel_Estudio_Nombre.Text))
                .rdoColumns("Descripcion") = Trim((Txt_Cat_Nivel_Estudio_Descripcion.Text))
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Nivel_Estudio.Close
    Set Rs_Modificacion_Cat_Nivel_Estudio = Nothing
    
    Conexion_Base.CommitTrans
    Grid_Cat_Nivel_Estudio.TextMatrix(Grid_Cat_Nivel_Estudio.RowSel, 1) = Trim(UCase(Txt_Cat_Nivel_Estudio_Nombre.Text))
    Grid_Cat_Nivel_Estudio.TextMatrix(Grid_Cat_Nivel_Estudio.RowSel, 2) = Trim((Txt_Cat_Nivel_Estudio_Descripcion.Text))
    Fra_Cat_Nivel_Estudio_Generales.Enabled = False
    Fra_Cat_Nivel_Estudio.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio", Me)
    MsgBox "Registro Actualizado correctamente", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
    
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Nive_Estudio
    'DESCRIPCIÓN:           Consulta la informacion de los nivel estudio
    'PARÁMETROS :           Nombre: Nivel a buscar
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           18 Enero 2011
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Nivel_Estudio(Nombre As String)
Dim Rs_Consulta_Cat_Nivel_Estudio As rdoResultset    'Informacion de las unidades de medida

On Error GoTo HANDLER
Grid_Cat_Nivel_Estudio.Rows = 0
'Consulta todos los registros dados de alta en la tabla Cat_Unidades_Medidas
Mi_SQL = "SELECT Nivel_Estudio_ID, Nombre, Descripcion"
Mi_SQL = Mi_SQL & " FROM Cat_Nivel_Estudio"
Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
Set Rs_Consulta_Cat_Nivel_Estudio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'Si se encontraron valores entonces los agrega al Grid_Unidades_Medidas
With Rs_Consulta_Cat_Nivel_Estudio
    If Not .EOF Then
        Grid_Cat_Nivel_Estudio.AddItem "Nivel_Estudio_ID" & Chr(9) & "Nombre" & Chr(9) & "Descripcion"
            'Llena el Grid_Unidades_Medidas con los valores traidos de la consulta
            While Not .EOF
                Grid_Cat_Nivel_Estudio.AddItem .rdoColumns("Nivel_Estudio_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Descripcion")
                .MoveNext
            Wend
        'Configura el Grid_Unidades_Medidas
        Grid_Cat_Nivel_Estudio.FixedCols = 1
        Grid_Cat_Nivel_Estudio.FixedRows = 1
        Grid_Cat_Nivel_Estudio.ColWidth(0) = 1000     'ID
        Grid_Cat_Nivel_Estudio.ColWidth(1) = 5500     'Nombre
        Grid_Cat_Nivel_Estudio.ColWidth(2) = 1000     'Descripcion
        .Close
    End If
End With
Set Rs_Consulta_Cat_Nivel_Estudio = Nothing
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*********************************termino Nivel Estudio ***************************************


'*********************************Inicio Motivos Baja***************************************
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Cat_Motivos_Baja
'DESCRIPCION: Alta de los motivos de baja
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Cat_Motivos_Baja()
Dim Rs_Alta_Cat_Motivos_Baja As rdoResultset     'Informacion de la base de datos

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Da de alta los valores introducidos por el usuario en la tabla Cat_Unidades_Medidas
    Set Rs_Alta_Cat_Motivos_Baja = Conectar_Ayudante.Recordset_Agregar("Cat_Motivos_Baja")
    With Rs_Alta_Cat_Motivos_Baja
        'Da de alta el registro de la unidad de medida en la tabla Cat_Unidades_Medidas
        .AddNew
            Txt_Cat_Motivos_Baja_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Motivos_Baja", "Motivo_Baja_ID"), "00000")
            .rdoColumns("Motivo_Baja_ID") = Trim(Txt_Cat_Motivos_Baja_ID.Text)
            .rdoColumns("Nombre") = Trim((Txt_Cat_Motivos_Baja_Nombre.Text))
            .rdoColumns("Descripcion") = Trim((Txt_Cat_Motivos_Baja_Descripcion.Text))
            .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Motivos_Baja.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Motivos_Baja.Close
    Set Rs_Alta_Cat_Motivos_Baja = Nothing
    
    Conexion_Base.CommitTrans
    With Grid_Cat_Motivos_Baja
        If .Rows = 0 Then
            .AddItem "Motivo_Baja_ID" & Chr(9) & "Nombre" & Chr(9) & "Descripcion"
        End If
        .AddItem Trim(Txt_Cat_Motivos_Baja_ID.Text) & Chr(9) & _
            Trim(Txt_Cat_Motivos_Baja_Nombre.Text) & Chr(9) & _
            Trim(Txt_Cat_Motivos_Baja_Descripcion.Text)
        'Configura el tamaño de las celdas del grid
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 0     'ID
        .ColWidth(1) = 4000     'Nombre
        .ColWidth(2) = 3500     'Abreviatura
    End With
    'Habilita y deshabilita los controles de la forma
    Fra_Cat_Motivos_Baja.Enabled = True
    Fra_Cat_Motivos_Baja_Generales.Enabled = False
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Consultar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja", Me)
    MsgBox "Motivo de Baja Registrado", vbInformation + vbOKOnly, Me.Caption
Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Cat_Motivos_Baja
'DESCRIPCION: Modifica los motivos de baja
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Cat_Motivos_Baja()
Dim Rs_Modificacion_Cat_Motivos_Baja As rdoResultset 'Modifica los valores de la unida de medida que selecciono el usuario

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Informacion de la unidad de medida
    Mi_SQL = "SELECT * FROM Cat_Motivos_Baja"
    Mi_SQL = Mi_SQL & " WHERE Motivo_Baja_ID='" & Trim(Txt_Cat_Motivos_Baja_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Motivos_Baja = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Motivos_Baja.EOF Then
        'Actualiza los valores del regisro seleccionado
        With Rs_Modificacion_Cat_Motivos_Baja
            .Edit
                .rdoColumns("Nombre") = Trim((Txt_Cat_Motivos_Baja_Nombre.Text))
                .rdoColumns("Descripcion") = Trim((Txt_Cat_Motivos_Baja_Descripcion.Text))
                .rdoColumns("Clave_SAP") = Trim(Txt_Clave_SAP_Motivos_Baja.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Motivos_Baja.Close
    Set Rs_Modificacion_Cat_Motivos_Baja = Nothing
    Conexion_Base.CommitTrans
    Grid_Cat_Motivos_Baja.TextMatrix(Grid_Cat_Motivos_Baja.RowSel, 1) = Trim(UCase(Txt_Cat_Motivos_Baja_Nombre.Text))
    Grid_Cat_Motivos_Baja.TextMatrix(Grid_Cat_Motivos_Baja.RowSel, 2) = Trim((Txt_Cat_Motivos_Baja_Descripcion.Text))
    Grid_Cat_Motivos_Baja.TextMatrix(Grid_Cat_Motivos_Baja.RowSel, 3) = Trim(Txt_Clave_SAP_Motivos_Baja.Text)
    Fra_Cat_Motivos_Baja_Generales.Enabled = False
    Fra_Cat_Motivos_Baja.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja", Me)
    MsgBox "Registro Actualizado correctamente", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Motivos_Baja
'DESCRIPCION: Consulta la informacion de los motivos de baja
'PARAMETROS : Nombre: Motivo a buscar
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Consulta_Cat_Motivos_Baja(Nombre As String)
Dim Rs_Consulta_Cat_Motivos_Baja As rdoResultset    'Informacion de las unidades de medida

On Error GoTo HANDLER
    Grid_Cat_Motivos_Baja.Rows = 0
    'Consulta todos los registros dados de alta en la tabla Cat_Unidades_Medidas
    Mi_SQL = "SELECT * FROM Cat_Motivos_Baja"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Set Rs_Consulta_Cat_Motivos_Baja = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si se encontraron valores entonces los agrega al Grid_Unidades_Medidas
    With Rs_Consulta_Cat_Motivos_Baja
        If Not .EOF Then
            Grid_Cat_Motivos_Baja.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Descripcion" & Chr(9) & "ClaveSAP"
            'Llena el Grid_Unidades_Medidas con los valores traidos de la consulta
            While Not .EOF
                Grid_Cat_Motivos_Baja.AddItem .rdoColumns("Motivo_Baja_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Descripcion") & Chr(9) & .rdoColumns("Clave_SAP")
                .MoveNext
            Wend
            'Configura el Grid_Unidades_Medidas
            Grid_Cat_Motivos_Baja.FixedRows = 1
            Grid_Cat_Motivos_Baja.ColWidth(0) = 1000     'ID
            Grid_Cat_Motivos_Baja.ColWidth(1) = 3000     'Nombre
            Grid_Cat_Motivos_Baja.ColWidth(2) = 3500     'Descripcion
            Grid_Cat_Motivos_Baja.ColWidth(3) = 0        'ClaveSAP
            .Close
        End If
    End With
    Set Rs_Consulta_Cat_Motivos_Baja = Nothing
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*********************************termino Motivos Baja ***************************************

'************************************************Inicio Turnos***************************************
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Turnos
'DESCRIPCION: Consulta los turnos registrados
'PARAMETROS : Nombre: Nombre del turno a buscar
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Consulta_Cat_Turnos(Nombre As String)
Dim Rs_Consulta_Cat_Turnos As rdoResultset     'Informacion de los Maquinas

    Grid_Cat_Turnos.Rows = 0
    'Consulta todos los roles que se encuentran dados de alta
    Mi_SQL = "SELECT Turno_ID,Nombre,Hora_Inicio,Hora_Termino,Comentarios"
    Mi_SQL = Mi_SQL & " FROM Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " AND ISNULL(Estatus, '') <> 'INACTIVO'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Turnos.EOF Then
        With Rs_Consulta_Cat_Turnos
            Grid_Cat_Turnos.AddItem "Turno ID" & Chr(9) & "Nombre" & Chr(9) & "Inicio" & Chr(9) & "Termino" & Chr(9) & "Comentarios"
            While Not .EOF
                Grid_Cat_Turnos.AddItem .rdoColumns("Turno_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & _
                                    Format(.rdoColumns("Hora_Inicio"), "HH:mm:ss") & Chr(9) & Format(.rdoColumns("Hora_termino"), "HH:mm:ss") & Chr(9) & _
                                    .rdoColumns("Comentarios")
                .MoveNext
            Wend
            'Asigna los tamaños de las columnas del grid_roles
            Grid_Cat_Turnos.FixedRows = 1
            Grid_Cat_Turnos.ColWidth(0) = 1000  'Turno_ID
            Grid_Cat_Turnos.ColWidth(1) = 4000  'Nombre_Turno
            Grid_Cat_Turnos.ColAlignment(1) = flexAlignLeftCenter
            Grid_Cat_Turnos.ColWidth(2) = 1000  'Hora Inicio
            Grid_Cat_Turnos.ColWidth(3) = 1000  'Hora Termino
            Grid_Cat_Turnos.ColWidth(4) = 0     'Comentarios
        End With
    End If
    Rs_Consulta_Cat_Turnos.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cat_Calendarios_Turnos
'DESCRIPCION: Consulta los Calendarios registrados
'PARAMETROS : Nombre: Nombre del Calendario a buscar
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 20/Abril/2015
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Consulta_Cat_Calendarios_Turnos(Nombre As String)
Dim Rs_Consulta_Cat_Calendarios_Turnos As rdoResultset     'Informacion de los Maquinas

    Grid_Calendarios_Configuracion_Turnos.Rows = 0
    Grid_Calendarios_Configuracion_Turnos.Cols = 0
    Grid_Calendarios_Turnos.Rows = 0
    'Consulta todos los roles que se encuentran dados de alta
    Mi_SQL = "SELECT Calendario_Turno_ID,Nombre,Fecha_Inicio,Fecha_Termino,Comentarios"
    Mi_SQL = Mi_SQL & " FROM Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Calendarios_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Calendarios_Turnos.EOF Then
        With Rs_Consulta_Cat_Calendarios_Turnos
            Grid_Calendarios_Turnos.AddItem "Calendario ID" & Chr(9) & "Nombre" & Chr(9) & "Inicio" & Chr(9) & "Termino" & Chr(9) & "Comentarios"
            While Not .EOF
                Grid_Calendarios_Turnos.AddItem .rdoColumns("Calendario_Turno_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & _
                                    Format(.rdoColumns("Fecha_Inicio"), "dd/MMM/yyyy") & Chr(9) & Format(.rdoColumns("Fecha_Termino"), "dd/MMM/yyyy") & Chr(9) & _
                                    .rdoColumns("Comentarios")
                .MoveNext
            Wend
            'Asigna los tamaños de las columnas del grid_roles
            Grid_Calendarios_Turnos.FixedRows = 1
            Grid_Calendarios_Turnos.ColWidth(0) = 1050  'Calendario_Turno_ID
            Grid_Calendarios_Turnos.ColWidth(1) = 7250  'Nombre_Turno
            Grid_Calendarios_Turnos.ColAlignment(1) = flexAlignLeftCenter
            Grid_Calendarios_Turnos.ColWidth(2) = 1200  'Fecha Inicio
            Grid_Calendarios_Turnos.ColWidth(3) = 1200  'Fecha Termino
            Grid_Calendarios_Turnos.ColWidth(4) = 0     'Comentarios
        End With
    End If
    Rs_Consulta_Cat_Calendarios_Turnos.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Detalles_Turno
'DESCRIPCION: Consulta los detalles del turno
'PARAMETROS :
'CREO       : Julio César Cruz Paredes
'FECHA_CREO : 08-Abril-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Detalles_Turno()
Dim Rs_Consulta_Cat_Turnos_Detalles As rdoResultset '

On Error GoTo HANDLER
    Grid_Detalles_Turnos.Rows = 0
    Mi_SQL = " SELECT * FROM Cat_Turnos_Detalles "
    Mi_SQL = Mi_SQL & " WHERE Turno_ID ='" & Txt_Cat_Turnos_Turno_ID.Text & "'"
    Set Rs_Consulta_Cat_Turnos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Turnos_Detalles
        If Not .EOF Then
            Dtp_Cat_Turnos_Hora_Inicio.Value = .rdoColumns("Hora_Inicio")
            Dtp_Cat_Turnos_Hora_Termino.Value = .rdoColumns("Hora_Termino")
            Dtp_Cat_Turnos_Comida_Inicio.Value = .rdoColumns("Comida_Inicio")
            Dtp_Cat_Turnos_Comida_Termino.Value = .rdoColumns("Comida_Termino")
            Txt_Horas_Turno.Text = .rdoColumns("Horas_Turno")
            Txt_Horas_Comida.Text = .rdoColumns("Horas_Comida")
            If Grid_Detalles_Turnos.Rows < 1 Then
                ''Se agrega encabezado
                Grid_Detalles_Turnos.Cols = 9
                Grid_Detalles_Turnos.AddItem "Turno ID" _
                   & Chr(9) & "Dia" _
                   & Chr(9) & "H.Inicio" _
                   & Chr(9) & "H.Termino" _
                   & Chr(9) & "C.Inicio" _
                   & Chr(9) & "C.Termino" _
                   & Chr(9) & "H.Turnos" _
                   & Chr(9) & "H.Comida" _
                   & Chr(9) & "Descanso"
            End If
            While Not .EOF
                'Se agrega el registro
                Grid_Detalles_Turnos.AddItem .rdoColumns("Turno_ID") _
                    & Chr(9) & .rdoColumns("Dia_Semana") _
                    & Chr(9) & Format(.rdoColumns("Hora_Inicio"), "HH:mm") _
                    & Chr(9) & Format(.rdoColumns("Hora_Termino"), "HH:mm") _
                    & Chr(9) & Format(.rdoColumns("Comida_Inicio"), "HH:mm") _
                    & Chr(9) & Format(.rdoColumns("Comida_Termino"), "HH:mm") _
                    & Chr(9) & .rdoColumns("Horas_Turno") _
                    & Chr(9) & .rdoColumns("Horas_Comida") _
                    & Chr(9) & Rs_Consulta_Cat_Turnos_Detalles!Dia_Descanso
                .MoveNext
            Wend
            'Se formatea el grid
            With Grid_Detalles_Turnos
                .FixedRows = 1
                .FixedCols = 2
                .ColWidth(0) = 0  'Dia
                .ColWidth(1) = 800  'H.Inicio
                .ColAlignment(1) = flexAlignCenterTop
                .ColWidth(2) = 1000   'H.Termino
                .ColAlignment(2) = flexAlignCenterTop
                .ColWidth(3) = 1000  'C.Inicio
                .ColAlignment(3) = flexAlignCenterTop
                .ColWidth(4) = 1000     'C.Termino
                .ColAlignment(4) = flexAlignCenterTop
                .ColWidth(5) = 1000     'H.Turnos
                .ColAlignment(5) = flexAlignCenterTop
                .ColWidth(6) = 1000     'H.Comida
                .ColAlignment(6) = flexAlignCenterTop
                .ColWidth(7) = 1000     'Descanso
                .ColAlignment(7) = flexAlignCenterTop
                .ColWidth(8) = 870     'Descanso
                .ColAlignment(8) = flexAlignCenterTop
            End With
        End If
    End With
    Exit Sub
HANDLER:
    MsgBox Err.Description, vbExclamation
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Detalles_Calendarios_Turnos
'DESCRIPCION: Consulta los detalles del calendario del turno
'PARAMETROS :
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 20/Abril/2015
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Detalles_Calendarios_Turnos()
Dim Rs_Consulta_Cat_Calendarios_Turnos_Detalles As rdoResultset '
Dim Cont_Filas As Integer
Dim Cont_Columnas As Integer
Dim Turno As String

On Error GoTo HANDLER
    Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID ='" & Txt_Calendario_Turno_ID.Text & "'"
    Mi_SQL = Mi_SQL & " AND Estatus <> 'ELIMINADO'"
    Set Rs_Consulta_Cat_Calendarios_Turnos_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Calendarios_Turnos_Detalles
        If Not .EOF Then
            Grid_Calendarios_Configuracion_Turnos.Redraw = False
            While Not .EOF
                'Se agrega el registro
                For Cont_Filas = Grid_Calendarios_Configuracion_Turnos.FixedRows To Grid_Calendarios_Configuracion_Turnos.Rows - 1
                    If Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4)) = .rdoColumns("Semana") Then
                        Grid_Calendarios_Configuracion_Turnos.Row = Cont_Filas
                        Exit For
                    End If
                Next Cont_Filas
                For Cont_Columnas = 2 To Grid_Calendarios_Configuracion_Turnos.Cols - 5 Step 6
                    If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas)) = .rdoColumns("Dia_Semana") Then
                        Grid_Calendarios_Configuracion_Turnos.Col = Cont_Columnas
                        Exit For
                    End If
                Next Cont_Columnas
                
'                Grid_Calendarios_Configuracion_Turnos.CellBackColor = Porcentaje_Rango(255, 16777215, Val(.rdoColumns("Turno_ID")) / (Cmb_Calendario_Turnos.ListCount * (1.3)))
                Grid_Calendarios_Configuracion_Turnos.CellBackColor = Obtener_Codigo_Color(Convertir_Cadena_A_Numero(Trim(.rdoColumns("Nombre_Turno"))), 255, 16777215)
                Turno = Trim(.rdoColumns("Nombre_Turno"))
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, 0) = .rdoColumns("Calendario_Turno_ID")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col) = Turno
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1) = .rdoColumns("Lista_Empleados")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 2) = Format(.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 3) = Format(.rdoColumns("Hora_Termino"), "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 4) = Format(.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 5) = Format(.rdoColumns("Comida_Termino"), "HH:mm:ss")
                .MoveNext
            Wend
            Grid_Calendarios_Configuracion_Turnos.Redraw = True
            'Se formatea el grid
            With Grid_Calendarios_Configuracion_Turnos
                .FixedRows = 1
                .FixedCols = 2
                .ColWidth(0) = 0  'Calendario_Turno_ID
                .ColWidth(1) = 700  'Semana
                .ColAlignment(1) = flexAlignCenterTop
                .ColWidth(2) = 1000     'Lunes
                .ColAlignment(2) = flexAlignCenterTop
                .ColWidth(3) = 0  'Turno_ID
                .ColWidth(4) = 0       'Hora_Inicio
'                .ColAlignment(4) = flexAlignCenterTop
                .ColWidth(5) = 0       'Hora_Termino
'                .ColAlignment(5) = flexAlignCenterTop
                .ColWidth(6) = 0       'Comida_Inicio
'                .ColAlignment(6) = flexAlignCenterTop
                .ColWidth(7) = 0       'Comida_Termino
'                .ColAlignment(7) = flexAlignCenterTop
                .ColWidth(8) = 1000     'Martes
                .ColAlignment(8) = flexAlignCenterTop
                .ColWidth(9) = 0  'Turno_ID
                .ColWidth(10) = 0       'Hora_Inicio
'                .ColAlignment(10) = flexAlignCenterTop
                .ColWidth(11) = 0       'Hora_Termino
'                .ColAlignment(11) = flexAlignCenterTop
                .ColWidth(12) = 0       'Comida_Inicio
'                .ColAlignment(12) = flexAlignCenterTop
                .ColWidth(13) = 0       'Comida_Termino
'                .ColAlignment(13) = flexAlignCenterTop
                .ColWidth(14) = 1000     'Miércoles
                .ColAlignment(14) = flexAlignCenterTop
                .ColWidth(15) = 0  'Turno_ID
                .ColWidth(16) = 0       'Hora_Inicio
'                .ColAlignment(16) = flexAlignCenterTop
                .ColWidth(17) = 0       'Hora_Termino
'                .ColAlignment(17) = flexAlignCenterTop
                .ColWidth(18) = 0       'Comida_Inicio
'                .ColAlignment(18) = flexAlignCenterTop
                .ColWidth(19) = 0       'Comida_Termino
'                .ColAlignment(19) = flexAlignCenterTop
                .ColWidth(20) = 1000     'Jueves
                .ColAlignment(20) = flexAlignCenterTop
                .ColWidth(21) = 0  'Turno_ID
                .ColWidth(22) = 0       'Hora_Inicio
'                .ColAlignment(22) = flexAlignCenterTop
                .ColWidth(23) = 0       'Hora_Termino
'                .ColAlignment(23) = flexAlignCenterTop
                .ColWidth(24) = 0       'Comida_Inicio
'                .ColAlignment(24) = flexAlignCenterTop
                .ColWidth(25) = 0       'Comida_Termino
'                .ColAlignment(25) = flexAlignCenterTop
                .ColWidth(26) = 1000     'Viernes
                .ColAlignment(26) = flexAlignCenterTop
                .ColWidth(27) = 0  'Turno_ID
                .ColWidth(28) = 0       'Hora_Inicio
'                .ColAlignment(28) = flexAlignCenterTop
                .ColWidth(29) = 0       'Hora_Termino
'                .ColAlignment(29) = flexAlignCenterTop
                .ColWidth(30) = 0       'Comida_Inicio
'                .ColAlignment(30) = flexAlignCenterTop
                .ColWidth(31) = 0       'Comida_Termino
'                .ColAlignment(31) = flexAlignCenterTop
                .ColWidth(32) = 1000     'Sábado
                .ColAlignment(32) = flexAlignCenterTop
                .ColWidth(33) = 0  'Turno_ID
                .ColWidth(34) = 0       'Hora_Inicio
'                .ColAlignment(34) = flexAlignCenterTop
                .ColWidth(35) = 0       'Hora_Termino
'                .ColAlignment(35) = flexAlignCenterTop
                .ColWidth(36) = 0       'Comida_Inicio
'                .ColAlignment(36) = flexAlignCenterTop
                .ColWidth(37) = 0       'Comida_Termino
'                .ColAlignment(37) = flexAlignCenterTop
                .ColWidth(38) = 1000     'Domingo
                .ColAlignment(38) = flexAlignCenterTop
                .ColWidth(39) = 0  'Turno_ID
                .ColWidth(40) = 0       'Hora_Inicio
'                .ColAlignment(40) = flexAlignCenterTop
                .ColWidth(41) = 0       'Hora_Termino
'                .ColAlignment(41) = flexAlignCenterTop
                .ColWidth(42) = 0       'Comida_Inicio
'                .ColAlignment(42) = flexAlignCenterTop
                .ColWidth(43) = 0       'Comida_Termino
'                .ColAlignment(43) = flexAlignCenterTop
            End With
        End If
    End With
    Exit Sub
HANDLER:
    MsgBox Err.Description, vbExclamation
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Cat_Turnos
'DESCRIPCION: Realiza el Alta de un turno en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 04-Febrero-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Cat_Turnos()
Dim Rs_Alta_Cat_Turnos As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    'Alta del turno
    Set Rs_Alta_Cat_Turnos = Conectar_Ayudante.Recordset_Agregar("Cat_Turnos")
    With Rs_Alta_Cat_Turnos
        .AddNew
            Txt_Cat_Turnos_Turno_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Turnos", "Turno_ID"), "00000")
            .rdoColumns("Turno_ID") = Trim(Txt_Cat_Turnos_Turno_ID.Text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Cat_Turnos_Nombre.Text))
            .rdoColumns("Hora_Inicio") = Format(Dtp_Cat_Turnos_Hora_Inicio.Value, "HH:mm:ss")
            .rdoColumns("Hora_Termino") = Format(Dtp_Cat_Turnos_Hora_Termino.Value, "HH:mm:ss")
            .rdoColumns("Comida_Inicio") = Format(Dtp_Cat_Turnos_Comida_Inicio.Value, "HH:mm:ss")
            .rdoColumns("Comida_Termino") = Format(Dtp_Cat_Turnos_Comida_Termino.Value, "HH:mm:ss")
            'Guarda las horas efectivas del turno
            Txt_Horas_Turno.Text = (Val(DateDiff("n", .rdoColumns("Hora_Inicio"), .rdoColumns("Comida_Inicio"))) + Val(DateDiff("n", .rdoColumns("Comida_Termino"), .rdoColumns("Hora_Termino")))) / 60
            .rdoColumns("Horas_Turno") = Val(Txt_Horas_Turno.Text)
            'Guarda las horas de comida
            Txt_Horas_Comida.Text = Val(DateDiff("n", .rdoColumns("Comida_Inicio"), .rdoColumns("Comida_Termino"))) / 60
            .rdoColumns("Horas_Comida") = Val(Txt_Horas_Comida.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Turnos_Comentarios.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Turnos.Close
    
    ''Se da de alta el registro
    Call Alta_Cat_Turnos_Detalles("NO")
    
    MsgBox "El Turno ha sido dado de alta", vbInformation
    'Pone un encabezado en el grid
    If Grid_Cat_Turnos.Rows = 0 Then
        Grid_Cat_Turnos.AddItem "Turno ID" & Chr(9) & "Nombre" & Chr(9) & "Inicio" & Chr(9) & "Termino" & Chr(9) & "Comentarios"
        Grid_Cat_Turnos.ColWidth(0) = 1000  'Turno_ID
        Grid_Cat_Turnos.ColWidth(1) = 4000  'Nombre_Turno
        Grid_Cat_Turnos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Cat_Turnos.ColWidth(2) = 1000  'Hora Inicio
        Grid_Cat_Turnos.ColWidth(3) = 1000  'Hora Termino
        Grid_Cat_Turnos.ColWidth(4) = 0     'Comentarios
    End If
    Grid_Cat_Turnos.AddItem Trim(Txt_Cat_Turnos_Turno_ID.Text) & Chr(9) & Trim(Txt_Cat_Turnos_Nombre.Text) _
        & Chr(9) & Format(Dtp_Cat_Turnos_Hora_Inicio.Value, "HH:mm:ss") _
        & Chr(9) & Format(Dtp_Cat_Turnos_Hora_Termino.Value, "HH:mm:ss") _
        & Chr(9) & Trim(Txt_Cat_Turnos_Comentarios.Text)
    Grid_Cat_Turnos.FixedRows = 1
    Dtp_Cat_Turnos_Hora_Inicio.Value = "00:00"
    Dtp_Cat_Turnos_Hora_Termino.Value = "00:00"
    Dtp_Cat_Turnos_Comida_Inicio.Value = "00:00"
    Dtp_Cat_Turnos_Comida_Termino.Value = "00:00"
    Btn_Salir_Click
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Calendario_Turnos
'DESCRIPCION: Realiza el Alta del calendario de turnos en la base de datos
'PARAMETROS :
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 20/Abril/2015
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Calendario_Turnos()
Dim Rs_Alta_Calendario_Turnos As rdoResultset 'Informacion del Maquinas
Dim Cont_Filas As Integer

On Error GoTo HANDLER
    If Dtp_Calendario_Fecha_Inicio.Value <= Dtp_Calendario_Fecha_Termino.Value Then
        'Alta del turno
        If Grid_Calendarios_Configuracion_Turnos.Rows >= 2 Then
            Dtp_Calendario_Fecha_Inicio.Value = DateAdd("d", Obtener_Numero_Dia_Semana(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows - 1, 2)) - 1, DateAdd("ww", Replace(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), "SEM", "") - 1, DateSerial(Dtp_Calendario_Fecha_Inicio.Year, 1, 1)))
            Dtp_Calendario_Fecha_Termino.Value = DateAdd("d", Obtener_Numero_Dia_Semana(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows - 1, Grid_Calendarios_Configuracion_Turnos.Cols - 6)) - 1, DateAdd("ww", Replace(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), "SEM", ""), DateSerial(Dtp_Calendario_Fecha_Inicio.Year, 1, 1)))
        End If
        Set Rs_Alta_Calendario_Turnos = Conectar_Ayudante.Recordset_Agregar("Cat_Calendarios_Turnos")
        With Rs_Alta_Calendario_Turnos
            .AddNew
                Txt_Calendario_Turno_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Calendarios_Turnos", "Calendario_Turno_ID"), "00000")
                .rdoColumns("Calendario_Turno_ID") = Trim(Txt_Calendario_Turno_ID.Text)
                .rdoColumns("Estatus") = "ACTIVO"
                .rdoColumns("Nombre") = Trim(UCase(Txt_Calendario_Nombre.Text))
                .rdoColumns("Fecha_Inicio") = Dtp_Calendario_Fecha_Inicio.Value
                .rdoColumns("Fecha_Termino") = Dtp_Calendario_Fecha_Termino.Value
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Calendario_Comentarios.Text))
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        End With
        Rs_Alta_Calendario_Turnos.Close
        
        ''Se da de alta el registro
        Call Alta_Calendario_Turnos_Detalles("NO")
        
        MsgBox "El Turno ha sido dado de alta", vbInformation
        'Pone un encabezado en el grid
        If Grid_Calendarios_Turnos.Rows = 0 Then
            Grid_Calendarios_Turnos.AddItem "Calendario ID" & Chr(9) & "Nombre" & Chr(9) & "Inicio" & Chr(9) & "Termino" & Chr(9) & "Comentarios"
            'Asigna los tamaños de las columnas del grid_roles
            Grid_Calendarios_Turnos.ColWidth(0) = 1000  'Calendario_Turno_ID
            Grid_Calendarios_Turnos.ColWidth(1) = 4000  'Nombre_Turno
            Grid_Calendarios_Turnos.ColAlignment(1) = flexAlignLeftCenter
            Grid_Calendarios_Turnos.ColWidth(2) = 1200  'Fecha_Inicio
            Grid_Calendarios_Turnos.ColWidth(3) = 1200  'Fecha_Termino
            Grid_Calendarios_Turnos.ColWidth(4) = 0     'Comentarios
        End If
        Grid_Calendarios_Turnos.AddItem Trim(Txt_Calendario_Turno_ID.Text) _
            & Chr(9) & Trim(Txt_Calendario_Nombre.Text) _
            & Chr(9) & Format(Dtp_Calendario_Fecha_Inicio.Value, "dd/MMM/yyyy") _
            & Chr(9) & Format(Dtp_Calendario_Fecha_Termino.Value, "dd/MMM/yyyy") _
            & Chr(9) & Trim(Txt_Calendario_Comentarios.Text)
        Grid_Calendarios_Turnos.FixedRows = 1
        Txt_Calendarios_Configuracion_Turno.Text = ""
        Txt_Calendario_Nombre.Text = ""
        Txt_Calendario_Comentarios.Text = ""
        Dtp_Calendario_Fecha_Inicio.Value = DateValue(Now)
        Dtp_Calendario_Fecha_Termino.Value = DateValue(Now)
        Dtp_Calendario_Hora_Inicio.Value = TimeSerial(0, 0, 0)
        Dtp_Calendario_Hora_Termino.Value = TimeSerial(0, 0, 0)
        Dtp_Calendario_Inicio_Comida.Value = TimeSerial(0, 0, 0)
        Dtp_Calendario_Termino_Comida.Value = TimeSerial(0, 0, 0)
        Txt_Calendario_Horas_Turno.Text = ""
        Txt_Calendario_Horas_Comida.Text = ""
        Grid_Calendarios_Configuracion_Turnos.Rows = 0
        Grid_Calendarios_Configuracion_Turnos.Cols = 0
        Btn_Salir_Click
    Else
        MsgBox "Revise las fechas de Inicio y Término. No se puede dar de alta el Calendario.", vbInformation
    End If
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Cat_Turnos_Detalles
'DESCRIPCION: Realiza el Alta un detalle de un turno
'PARAMETROS : Eliminar_Detalles indica si se eliminaran los detalles anteriores
'CREO       : Julio César Cruz Paredes
'FECHA_CREO : 08-Abril-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Cat_Turnos_Detalles(Eliminar_Detalles As String)
Dim Rs_Alta_Cat_Turnos_Detalles As rdoResultset 'Informacion del Maquinas
Dim Cont_Fila As Integer
Dim Rd_Eliminar As rdoResultset
Dim Mi_SQL As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    If Eliminar_Detalles = "SI" Then ' valida si se van a eliminar los registros
        Mi_SQL = " SELECT * FROM Cat_Turnos_Detalles "
        Mi_SQL = Mi_SQL & " WHERE Turno_ID ='" & Txt_Cat_Turnos_Turno_ID.Text & "'"
        Set Rd_Eliminar = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
        While Not Rd_Eliminar.EOF
            Rd_Eliminar.Delete
            Rd_Eliminar.MoveNext
        Wend
        Rd_Eliminar.Close
    End If
    For Cont_Fila = 1 To Grid_Detalles_Turnos.Rows - 1
        'Alta
        Set Rs_Alta_Cat_Turnos_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Turnos_Detalles")
        'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
        With Rs_Alta_Cat_Turnos_Detalles
            .AddNew
                .rdoColumns("Turno_ID") = Trim(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Dia_Semana") = Trim(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 1))
                .rdoColumns("Hora_Inicio") = Format(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 2), "HH:mm:ss")
                .rdoColumns("Hora_Termino") = Format(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 3), "HH:mm:ss")
                .rdoColumns("Comida_Inicio") = Format(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 4), "HH:mm:ss")
                .rdoColumns("Comida_Termino") = Format(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 5), "HH:mm:ss")
                'Guarda las horas efectivas del turno
                .rdoColumns("Horas_Turno") = Val(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 6))
                'Guarda las horas de comida
                .rdoColumns("Horas_Comida") = Val(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 7))
                .rdoColumns("Dia_Descanso") = Trim(Grid_Detalles_Turnos.TextMatrix(Cont_Fila, 8))
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
            .Close
        End With
        Set Rs_Alta_Cat_Turnos_Detalles = Nothing
    Next
    Conexion_Base.CommitTrans
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Calendario_Turnos_Detalles
'DESCRIPCION: Realiza el Alta del los detalles del Calendario
'PARAMETROS : Eliminar_Detalles indica si se eliminaran los detalles anteriores
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 20/Abril/2015
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Alta_Calendario_Turnos_Detalles(Eliminar_Detalles As String)
Dim Rs_Alta_Calendario_Turnos_Detalles As rdoResultset 'Informacion del Maquinas
Dim Cont_Filas As Integer
Dim Cont_Columnas As Integer
Dim Rd_Consultar As rdoResultset
Dim Rd_Eliminar As rdoResultset
Dim Rd_Actualizar As rdoResultset
Dim Rd_Actualizar_Cat_Calendarios_Turnos As rdoResultset
Dim Rs_Alta_Calendarios_Turnos_Roles As rdoResultset
Dim Mi_SQL As String
Dim Dia_Semana As String
Dim Estatus As String
Dim Calendario_Turno_Detalle_ID As String
Dim Lista_Empleados() As String
Dim Cont_Lista_Empleados As Integer
Dim Calendario_Turno_Rol_ID As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    If Eliminar_Detalles = "SI" Then ' valida si se van a eliminar los registros
        Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos_Roles "
        Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID ='" & Txt_Calendario_Turno_ID.Text & "'"
        Set Rd_Eliminar = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
        While Not Rd_Eliminar.EOF
            Rd_Eliminar.Delete
            Rd_Eliminar.MoveNext
        Wend
        Rd_Eliminar.Close
'
'        Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos_Detalles "
'        Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID ='" & Txt_Calendario_Turno_ID.Text & "'"
'        Set Rd_Eliminar = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
'        While Not Rd_Eliminar.EOF
'            Rd_Eliminar.Delete
'            Rd_Eliminar.MoveNext
'        Wend
'        Rd_Eliminar.Close
    End If
    Estatus = "INACTIVO"
    For Cont_Filas = Grid_Calendarios_Configuracion_Turnos.FixedRows To Grid_Calendarios_Configuracion_Turnos.Rows - 1
        For Cont_Columnas = 2 To Grid_Calendarios_Configuracion_Turnos.Cols - 5 Step 6
            'Alta
            If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas)) <> "" Then
                Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos_Detalles "
                Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID = '" & Txt_Calendario_Turno_ID.Text & "'"
                Mi_SQL = Mi_SQL & " AND Semana = '" & Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4)) & "'"
                Mi_SQL = Mi_SQL & " AND Dia_Semana = '" & Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas)) & "'"
                Set Rd_Actualizar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Rd_Actualizar.EOF Then
                    Set Rs_Alta_Calendario_Turnos_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Calendarios_Turnos_Detalles")
                    'Llena la tabla de Cat_Calendarios_Turnos_Detalles con los datos contenidos en el grid
                    With Rs_Alta_Calendario_Turnos_Detalles
                        .AddNew
                            .rdoColumns("Calendario_Turno_ID") = Txt_Calendario_Turno_ID.Text
                            Calendario_Turno_Detalle_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Calendarios_Turnos_Detalles WHERE Calendario_Turno_ID = '" & Txt_Calendario_Turno_ID.Text & "'", "Calendario_Turno_Detalle_ID"), "00000")
                            .rdoColumns("Calendario_Turno_Detalle_ID") = Calendario_Turno_Detalle_ID
                            Dia_Semana = Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas))
                            If Val(Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4))) > DatePart("ww", Now) Then
                                .rdoColumns("Estatus") = "PENDIENTE"
                            Else
                                If Switch(Dia_Semana = "Lunes", 2, Dia_Semana = "Martes", 3, Dia_Semana = "Miércoles", 4, Dia_Semana = "Jueves", 5, Dia_Semana = "Viernes", 6, Dia_Semana = "Sábado", 7, Dia_Semana = "Domingo", 1) >= DatePart("w", Now) Then
                                    .rdoColumns("Estatus") = "PENDIENTE"
                                Else
                                    .rdoColumns("Estatus") = "VENCIDO"
                                End If
                            End If
                            .rdoColumns("Nombre_Turno") = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas)
                            .rdoColumns("Semana") = Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4))
                            .rdoColumns("Dia_Semana") = Dia_Semana
                            .rdoColumns("Hora_Inicio") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 2), "HH:mm:ss")
                            .rdoColumns("Hora_Termino") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 3), "HH:mm:ss")
                            .rdoColumns("Comida_Inicio") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 4), "HH:mm:ss")
                            .rdoColumns("Comida_Termino") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 5), "HH:mm:ss")
                            .rdoColumns("Lista_Empleados") = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 1)
                            .rdoColumns("Usuario_Creo") = Nombre_Usuario
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                        .Close
                    End With
                    Set Rs_Alta_Calendario_Turnos_Detalles = Nothing
                Else
                    With Rd_Actualizar
                        Calendario_Turno_Detalle_ID = .rdoColumns("Calendario_Turno_Detalle_ID")
                        .Edit
                            Dia_Semana = Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas))
                            If Val(Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4))) > DatePart("ww", Now) Then
                                .rdoColumns("Estatus") = "PENDIENTE"
                            Else
                                If Switch(Dia_Semana = "Lunes", 2, Dia_Semana = "Martes", 3, Dia_Semana = "Miércoles", 4, Dia_Semana = "Jueves", 5, Dia_Semana = "Viernes", 6, Dia_Semana = "Sábado", 7, Dia_Semana = "Domingo", 1) >= DatePart("w", Now) Then
                                    .rdoColumns("Estatus") = "PENDIENTE"
                                Else
                                    .rdoColumns("Estatus") = "VENCIDO"
                                End If
                            End If
                            .rdoColumns("Nombre_Turno") = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas)
                            .rdoColumns("Semana") = Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4))
                            .rdoColumns("Dia_Semana") = Dia_Semana
                            .rdoColumns("Hora_Inicio") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 2), "HH:mm:ss")
                            .rdoColumns("Hora_Termino") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 3), "HH:mm:ss")
                            .rdoColumns("Comida_Inicio") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 4), "HH:mm:ss")
                            .rdoColumns("Comida_Termino") = Format(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 5), "HH:mm:ss")
                            .rdoColumns("Lista_Empleados") = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 1)
                            .rdoColumns("Usuario_Creo") = Nombre_Usuario
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                        .Close
                    End With
                End If
                Set Rd_Actualizar = Nothing
                Estatus = "ACTIVO"
                
                If Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 1)) <> "" Then
                    Lista_Empleados = Split(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, Cont_Columnas + 1), ",")
                    Calendario_Turno_Rol_ID = Conectar_Ayudante.Maximo_Catalogo("Cat_Calendarios_Turnos_Roles", "Calendario_Turno_Rol_ID")
                    Set Rs_Alta_Calendarios_Turnos_Roles = Conectar_Ayudante.Recordset_Agregar("Cat_Calendarios_Turnos_Roles")
                    'Llena la tabla de Cat_Calendarios_Turnos_Roles con los datos contenidos en el grid
                    With Rs_Alta_Calendarios_Turnos_Roles
                        For Cont_Lista_Empleados = 0 To UBound(Lista_Empleados)
                            .AddNew
                                .rdoColumns("Calendario_Turno_Rol_ID") = Format(Calendario_Turno_Rol_ID, "00000")
                                .rdoColumns("Calendario_Turno_ID") = Txt_Calendario_Turno_ID.Text
                                .rdoColumns("Calendario_Turno_Detalle_ID") = Calendario_Turno_Detalle_ID
                                .rdoColumns("No_Tarjeta") = Lista_Empleados(Cont_Lista_Empleados)
                            .Update
                            Calendario_Turno_Rol_ID = Val(Calendario_Turno_Rol_ID) + 1
                        Next Cont_Lista_Empleados
                        .Close
                    End With
                    Set Rs_Alta_Calendarios_Turnos_Roles = Nothing
                End If
            Else
                Mi_SQL = " SELECT DISTINCT Cat_Calendarios_Turnos_Detalles.* FROM Cat_Calendarios_Turnos_Detalles"
                Mi_SQL = Mi_SQL & " LEFT OUTER JOIN Adm_Asistencias ON Adm_Asistencias.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
                Mi_SQL = Mi_SQL & "     AND Adm_Asistencias.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID"
                Mi_SQL = Mi_SQL & " LEFT OUTER JOIN Cat_Calendarios_Turnos_Roles ON Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
                Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID"
                Mi_SQL = Mi_SQL & " WHERE Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = '" & Txt_Calendario_Turno_ID.Text & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Calendarios_Turnos_Detalles.Semana = '" & Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4)) & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Calendarios_Turnos_Detalles.Dia_Semana = '" & Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas)) & "'"
                Set Rd_Actualizar = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Rd_Actualizar.EOF Then
                    Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos_Detalles "
                    Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID = '" & Txt_Calendario_Turno_ID.Text & "'"
                    Mi_SQL = Mi_SQL & " AND Semana = '" & Trim(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Cont_Filas, 1), 4)) & "'"
                    Mi_SQL = Mi_SQL & " AND Dia_Semana = '" & Trim(Grid_Calendarios_Configuracion_Turnos.TextMatrix(0, Cont_Columnas)) & "'"
                    Set Rd_Eliminar = Conectar_Ayudante.Recordset_Eliminar(Mi_SQL)
                    While Not Rd_Eliminar.EOF
                        Rd_Eliminar.Delete
                        Rd_Eliminar.MoveNext
                    Wend
                    Rd_Eliminar.Close
                Else
                    With Rd_Actualizar
                        .Edit
                            .rdoColumns("Estatus") = "ELIMINADO"
                        .Update
                        .Close
                    End With
                End If
                Set Rd_Actualizar = Nothing
            End If
        Next Cont_Columnas
    Next Cont_Filas
    Mi_SQL = " SELECT * FROM Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID = '" & Txt_Calendario_Turno_ID.Text & "'"
    Set Rd_Actualizar_Cat_Calendarios_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rd_Actualizar_Cat_Calendarios_Turnos.EOF Then
        Rd_Actualizar_Cat_Calendarios_Turnos.Edit
            Rd_Actualizar_Cat_Calendarios_Turnos.rdoColumns("Estatus") = Estatus
        Rd_Actualizar_Cat_Calendarios_Turnos.Update
        Rd_Actualizar_Cat_Calendarios_Turnos.MoveNext
    End If
    Rd_Actualizar_Cat_Calendarios_Turnos.Close
    Conexion_Base.CommitTrans
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Cat_Turnos
'DESCRIPCION: Realiza la modificacion de turno en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 04-Febrero-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Cat_Turnos()
Dim Mi_SQL As String
Dim Rs_Modifica_Cat_Turnos As rdoResultset

On Error GoTo HANDLER
    'Consulta el turno para modificar
    Mi_SQL = "SELECT * FROM Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Turno_ID='" & Trim(Txt_Cat_Turnos_Turno_ID.Text) & "'"
    Set Rs_Modifica_Cat_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Cat_Turnos.EOF Then
        With Rs_Modifica_Cat_Turnos
            .Edit
                .rdoColumns("Nombre") = Trim(UCase(Txt_Cat_Turnos_Nombre.Text))
                .rdoColumns("Hora_Inicio") = Format(Dtp_Cat_Turnos_Hora_Inicio.Value, "HH:mm:ss")
                .rdoColumns("Hora_Termino") = Format(Dtp_Cat_Turnos_Hora_Termino.Value, "HH:mm:ss")
                .rdoColumns("Comida_Inicio") = Format(Dtp_Cat_Turnos_Comida_Inicio.Value, "HH:mm:ss")
                .rdoColumns("Comida_Termino") = Format(Dtp_Cat_Turnos_Comida_Termino.Value, "HH:mm:ss")
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Cat_Turnos_Comentarios.Text))
                'Guarda las horas efectivas del turno
                Txt_Horas_Turno.Text = (Val(DateDiff("n", .rdoColumns("Hora_Inicio"), .rdoColumns("Comida_Inicio"))) + Val(DateDiff("n", .rdoColumns("Comida_Termino"), .rdoColumns("Hora_Termino")))) / 60
                .rdoColumns("Horas_Turno") = Val(Txt_Horas_Turno.Text)
                'Guarda las horas de comida
                Txt_Horas_Comida.Text = Val(DateDiff("n", .rdoColumns("Comida_Inicio"), .rdoColumns("Comida_Termino"))) / 60
                .rdoColumns("Horas_Comida") = Val(Txt_Horas_Comida.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modifica_Cat_Turnos.Close
    Call Alta_Cat_Turnos_Detalles("SI")
    MsgBox "El Turno ha sido Modificado", vbInformation
    'Modifica la informacion en el grid
    Grid_Cat_Turnos.TextMatrix(Grid_Cat_Turnos.RowSel, 1) = Trim(Txt_Cat_Turnos_Nombre.Text)
    Grid_Cat_Turnos.TextMatrix(Grid_Cat_Turnos.RowSel, 2) = Format(Dtp_Cat_Turnos_Hora_Inicio.Value, "HH:mm:ss")
    Grid_Cat_Turnos.TextMatrix(Grid_Cat_Turnos.RowSel, 3) = Format(Dtp_Cat_Turnos_Hora_Termino.Value, "HH:mm:ss")
    Grid_Cat_Turnos.TextMatrix(Grid_Cat_Turnos.RowSel, 4) = Trim(Txt_Cat_Turnos_Comentarios.Text)
    Btn_Salir_Click
    Dtp_Cat_Turnos_Hora_Inicio.Value = "00:00"
    Dtp_Cat_Turnos_Hora_Termino.Value = "00:00"
    Dtp_Cat_Turnos_Comida_Inicio.Value = "00:00"
    Dtp_Cat_Turnos_Comida_Termino.Value = "00:00"
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Calendario_Turnos
'DESCRIPCION: Realiza la modificacion del calendario en la base de datos
'PARAMETROS :
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 20/Abril/2015
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Modifica_Calendario_Turnos()
Dim Mi_SQL As String
Dim Rs_Modifica_Calendario_Turnos As rdoResultset
Dim Cont_Filas As Integer

On Error GoTo HANDLER
    'Consulta el turno para modificar
    If Dtp_Calendario_Fecha_Inicio.Value <= Dtp_Calendario_Fecha_Termino.Value Then
        Mi_SQL = "SELECT * FROM Cat_Calendarios_Turnos"
        Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID='" & Trim(Txt_Calendario_Turno_ID.Text) & "'"
        Set Rs_Modifica_Calendario_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Modifica_Calendario_Turnos.EOF Then
            With Rs_Modifica_Calendario_Turnos
                .Edit
                    .rdoColumns("Nombre") = Trim(UCase(Txt_Calendario_Nombre.Text))
                    .rdoColumns("Fecha_Inicio") = Dtp_Calendario_Fecha_Inicio.Value
                    .rdoColumns("Fecha_Termino") = Dtp_Calendario_Fecha_Termino.Value
                    .rdoColumns("Comentarios") = Trim(UCase(Txt_Calendario_Comentarios.Text))
                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                    .rdoColumns("Fecha_Modifico") = Now
                .Update
            End With
        End If
        Rs_Modifica_Calendario_Turnos.Close
        Call Alta_Calendario_Turnos_Detalles("SI")
        MsgBox "El Calendario ha sido Modificado", vbInformation
        'Modifica la informacion en el grid
        Grid_Calendarios_Turnos.TextMatrix(Grid_Calendarios_Turnos.RowSel, 1) = Trim(Txt_Calendario_Nombre.Text)
        Grid_Calendarios_Turnos.TextMatrix(Grid_Calendarios_Turnos.RowSel, 2) = DateValue(Dtp_Calendario_Fecha_Inicio.Value)
        Grid_Calendarios_Turnos.TextMatrix(Grid_Calendarios_Turnos.RowSel, 3) = DateValue(Dtp_Calendario_Fecha_Termino.Value)
        Grid_Calendarios_Turnos.TextMatrix(Grid_Calendarios_Turnos.RowSel, 4) = Trim(Txt_Calendario_Comentarios.Text)
        Btn_Salir_Click
    Else
        MsgBox "Revise las fechas de Inicio y Término. No se puede modificar el Calendario.", vbInformation
    End If
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************************************
'NOMBRE_FUNCIÓN:
'DESCRIPCIÓN:
'PARÁMETROS:
'CREO:
'FECHA_CREO:
'MODIFICÓ:
'FECHA_MODIFICÓ:
'CAUSA_MODIFICACIÓN:
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
        MsgBox "El archivo que está intentando abrir no se encontró en el directorio indicado.  ", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    MsgBox Err.Description
End Sub

Private Sub Txt_Horas_Comida_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Horas_Comida, True)
End Sub

Private Sub Txt_Horas_Turno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Horas_Turno, True)
End Sub

'*******************************************************************************************************
'NOMBRE_FUNCIÓN     : Crear_Calendario
'DESCRIPCIÓN        : Arma la plantilla de semanas en grid
'PARÁMETROS:
'CREO               : Antonio Salvador Benavides Guardado
'FECHA_CREO         : 20/Abril/2015
'MODIFICÓ:
'FECHA_MODIFICÓ:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************************************
Private Sub Crear_Calendario()
Dim Cont_Semanas As Integer
Dim Cant_Semanas As Integer
Dim Cont_Filas As Integer
Dim Semana As Integer
Dim Primer_Semana As Integer
Dim Ultima_Semana As Integer
Dim Dia_Semana As Date
    
    If Grid_Calendarios_Configuracion_Turnos.Rows <= 0 Then
        Grid_Calendarios_Configuracion_Turnos.Cols = 44
        Grid_Calendarios_Configuracion_Turnos.AddItem "Calendario ID" _
           & Chr(9) & "Semana" _
           & Chr(9) & "Lunes" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Martes" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Miércoles" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Jueves" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Viernes" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Sábado" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida" _
           & Chr(9) & "Domingo" _
           & Chr(9) & "Turno_ID" & Chr(9) & "Hora_Inicio" & Chr(9) & "Hora_Termino" & Chr(9) & "Inicio_Comida" & Chr(9) & "Termino_Comida"
    End If
    
    If DatePart("w", Dtp_Calendario_Fecha_Inicio.Value) = vbMonday _
    And DatePart("w", Dtp_Calendario_Fecha_Termino.Value) = vbSunday Then
        Cant_Semanas = (DateDiff("d", Dtp_Calendario_Fecha_Inicio, Dtp_Calendario_Fecha_Termino)) \ 7 + 1
    Else
        Cant_Semanas = DateDiff("ww", Dtp_Calendario_Fecha_Inicio, Dtp_Calendario_Fecha_Termino) + 1
    End If
    
    If Cant_Semanas > 0 Then
        Dia_Semana = Dtp_Calendario_Fecha_Inicio.Value
        Semana = DatePart("ww", Dia_Semana)
        If Semana < Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), 4)) Then
            Cant_Semanas = Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), 4)) - Semana
            For Cont_Semanas = 1 To Cant_Semanas
                Semana = DatePart("ww", Dia_Semana)
                Call Grid_Calendarios_Configuracion_Turnos.AddItem("" & Chr(9) & "SEM" & Format(Semana, "00"), Grid_Calendarios_Configuracion_Turnos.FixedRows + (Cont_Semanas - 1))
                Dia_Semana = DateAdd("ww", 1, Dia_Semana)
            Next Cont_Semanas
        Else
            If Semana > Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), 4)) _
            And Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), 4)) > 0 _
            And Grid_Calendarios_Configuracion_Turnos.Rows > 2 Then
                Cant_Semanas = Semana - Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.FixedRows, 1), 4))
                For Cont_Semanas = 1 To Cant_Semanas
                    Call Grid_Calendarios_Configuracion_Turnos.RemoveItem(Grid_Calendarios_Configuracion_Turnos.FixedRows + (Cont_Semanas - 1))
                Next Cont_Semanas
            End If
        End If
        
        Dia_Semana = Dtp_Calendario_Fecha_Termino.Value
        Semana = DatePart("ww", Dia_Semana) - 1
        If Semana > Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), 4)) Then
            If Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), 4)) = 0 Then
                Cant_Semanas = DateDiff("ww", Dtp_Calendario_Fecha_Inicio.Value, Dtp_Calendario_Fecha_Termino.Value) + 1
            Else
                Cant_Semanas = Semana - Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), 4))
            End If
            Dia_Semana = DateAdd("ww", -Cant_Semanas + 1, Dia_Semana)
            For Cont_Semanas = 1 To Cant_Semanas
                Semana = DatePart("ww", Dia_Semana)
                Call Grid_Calendarios_Configuracion_Turnos.AddItem("" & Chr(9) & "SEM" & Format(Semana, "00"))
                Dia_Semana = DateAdd("ww", 1, Dia_Semana)
            Next Cont_Semanas
        Else
            If Semana < Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), 4)) _
            And Grid_Calendarios_Configuracion_Turnos.Rows > 2 Then
                If Grid_Calendarios_Configuracion_Turnos.Rows > 1 Then
                    Cant_Semanas = Val(Mid(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Rows - 1, 1), 4)) - Semana
                Else
                    Cant_Semanas = 1
                End If
                For Cont_Semanas = 1 To Cant_Semanas - 1
                    Call Grid_Calendarios_Configuracion_Turnos.RemoveItem(Grid_Calendarios_Configuracion_Turnos.Rows - 1)
                Next Cont_Semanas
            End If
        End If
    End If
    
    If Grid_Calendarios_Configuracion_Turnos.Cols > 1 _
    And Grid_Calendarios_Configuracion_Turnos.FixedCols = 0 Then
        Grid_Calendarios_Configuracion_Turnos.FixedCols = 2
    End If
    
    If Grid_Calendarios_Configuracion_Turnos.Rows > 1 _
    And Grid_Calendarios_Configuracion_Turnos.FixedRows = 0 Then
        Grid_Calendarios_Configuracion_Turnos.FixedRows = 1
    End If
            With Grid_Calendarios_Configuracion_Turnos
                .ColWidth(0) = 0  'Calendario_Turno_ID
                .ColWidth(1) = 700  'Semana
                .ColAlignment(1) = flexAlignCenterTop
                .ColWidth(2) = 1000     'Lunes
                .ColAlignment(2) = flexAlignCenterTop
                .ColWidth(3) = 0  'Turno_ID
                .ColWidth(4) = 0       'Hora_Inicio
'                .ColAlignment(4) = flexAlignCenterTop
                .ColWidth(5) = 0       'Hora_Termino
'                .ColAlignment(5) = flexAlignCenterTop
                .ColWidth(6) = 0       'Comida_Inicio
'                .ColAlignment(6) = flexAlignCenterTop
                .ColWidth(7) = 0       'Comida_Termino
'                .ColAlignment(7) = flexAlignCenterTop
                .ColWidth(8) = 1000     'Martes
                .ColAlignment(8) = flexAlignCenterTop
                .ColWidth(9) = 0  'Turno_ID
                .ColWidth(10) = 0       'Hora_Inicio
'                .ColAlignment(10) = flexAlignCenterTop
                .ColWidth(11) = 0       'Hora_Termino
'                .ColAlignment(11) = flexAlignCenterTop
                .ColWidth(12) = 0       'Comida_Inicio
'                .ColAlignment(12) = flexAlignCenterTop
                .ColWidth(13) = 0       'Comida_Termino
'                .ColAlignment(13) = flexAlignCenterTop
                .ColWidth(14) = 1000     'Miércoles
                .ColAlignment(14) = flexAlignCenterTop
                .ColWidth(15) = 0  'Turno_ID
                .ColWidth(16) = 0       'Hora_Inicio
'                .ColAlignment(16) = flexAlignCenterTop
                .ColWidth(17) = 0       'Hora_Termino
'                .ColAlignment(17) = flexAlignCenterTop
                .ColWidth(18) = 0       'Comida_Inicio
'                .ColAlignment(18) = flexAlignCenterTop
                .ColWidth(19) = 0       'Comida_Termino
'                .ColAlignment(19) = flexAlignCenterTop
                .ColWidth(20) = 1000     'Jueves
                .ColAlignment(20) = flexAlignCenterTop
                .ColWidth(21) = 0  'Turno_ID
                .ColWidth(22) = 0       'Hora_Inicio
'                .ColAlignment(22) = flexAlignCenterTop
                .ColWidth(23) = 0       'Hora_Termino
'                .ColAlignment(23) = flexAlignCenterTop
                .ColWidth(24) = 0       'Comida_Inicio
'                .ColAlignment(24) = flexAlignCenterTop
                .ColWidth(25) = 0       'Comida_Termino
'                .ColAlignment(25) = flexAlignCenterTop
                .ColWidth(26) = 1000     'Viernes
                .ColAlignment(26) = flexAlignCenterTop
                .ColWidth(27) = 0  'Turno_ID
                .ColWidth(28) = 0       'Hora_Inicio
'                .ColAlignment(28) = flexAlignCenterTop
                .ColWidth(29) = 0       'Hora_Termino
'                .ColAlignment(29) = flexAlignCenterTop
                .ColWidth(30) = 0       'Comida_Inicio
'                .ColAlignment(30) = flexAlignCenterTop
                .ColWidth(31) = 0       'Comida_Termino
'                .ColAlignment(31) = flexAlignCenterTop
                .ColWidth(32) = 1000     'Sábado
                .ColAlignment(32) = flexAlignCenterTop
                .ColWidth(33) = 0  'Turno_ID
                .ColWidth(34) = 0       'Hora_Inicio
'                .ColAlignment(34) = flexAlignCenterTop
                .ColWidth(35) = 0       'Hora_Termino
'                .ColAlignment(35) = flexAlignCenterTop
                .ColWidth(36) = 0       'Comida_Inicio
'                .ColAlignment(36) = flexAlignCenterTop
                .ColWidth(37) = 0       'Comida_Termino
'                .ColAlignment(37) = flexAlignCenterTop
                .ColWidth(38) = 1000     'Domingo
                .ColAlignment(38) = flexAlignCenterTop
                .ColWidth(39) = 0  'Turno_ID
                .ColWidth(40) = 0       'Hora_Inicio
'                .ColAlignment(40) = flexAlignCenterTop
                .ColWidth(41) = 0       'Hora_Termino
'                .ColAlignment(41) = flexAlignCenterTop
                .ColWidth(42) = 0       'Comida_Inicio
'                .ColAlignment(42) = flexAlignCenterTop
                .ColWidth(43) = 0       'Comida_Termino
'                .ColAlignment(43) = flexAlignCenterTop
            End With
End Sub

'*******************************************************************************************************
'NOMBRE_FUNCIÓN     : Asignar_Turno_Calendario
'DESCRIPCIÓN        : Asigna los datos de los controles a las celdas del grid
'PARÁMETROS:
'CREO               : Antonio Salvador Benavides Guardado
'FECHA_CREO         : 22/Abril/2015
'MODIFICÓ:
'FECHA_MODIFICÓ:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************************************
Private Sub Asignar_Turno_Calendario()
Dim Resultado As VbMsgBoxResult
Dim Turno As String
    If TimeValue(Dtp_Calendario_Hora_Inicio.Value) = TimeSerial(0, 0, 0) _
    And TimeValue(Dtp_Calendario_Hora_Inicio.Value) = TimeValue(Dtp_Calendario_Hora_Termino.Value) Then
        Grid_Calendarios_Configuracion_Turnos.CellBackColor = 0
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col) = ""
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1) = ""
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 2) = ""
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 3) = ""
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 4) = ""
        Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 5) = ""
    Else
        If TimeValue(Dtp_Calendario_Hora_Inicio.Value) = TimeValue(Dtp_Calendario_Hora_Termino.Value) _
        And Not Horas_Iguales_Confirmadas Then
            Resultado = MsgBox("Las Horas de Inicio y Término son iguales. Favor de confirmar.", vbExclamation + vbOKCancel, "Calendario de Turnos")
            If Resultado = vbOK Then
                Horas_Iguales_Confirmadas = True
            Else
                Horas_Iguales_Confirmadas = False
            End If
            If Horas_Iguales_Confirmadas Then
                Call Asignar_Turno_Calendario
            End If
        Else
'            Grid_Calendarios_Configuracion_Turnos.CellBackColor = Porcentaje_Rango(255, 16777215, Cmb_Calendario_Turnos.ItemData(Cmb_Calendario_Turnos.ListIndex) / (Cmb_Calendario_Turnos.ListCount * (1.3)))
            Grid_Calendarios_Configuracion_Turnos.CellBackColor = Obtener_Codigo_Color(Convertir_Cadena_A_Numero(Txt_Calendarios_Configuracion_Turno.Text), 255, 16777215)
            Turno = Txt_Calendarios_Configuracion_Turno.Text
            If Grid_Calendarios_Configuracion_Turnos.Col <= (Grid_Calendarios_Configuracion_Turnos.Cols - 5) Then
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col) = Turno
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1) = ""
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 2) = Format(Dtp_Calendario_Hora_Inicio.Value, "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 3) = Format(Dtp_Calendario_Hora_Termino.Value, "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 4) = Format(Dtp_Calendario_Inicio_Comida.Value, "HH:mm:ss")
                Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 5) = Format(Dtp_Calendario_Termino_Comida.Value, "HH:mm:ss")
            End If
        End If
    End If
End Sub

'*******************************************************************************************************
'NOMBRE_FUNCIÓN     : Obtener_Horario_Turno_Grid
'DESCRIPCIÓN        : Asigna los datos leidos del grid a los controles del formulario
'PARÁMETROS:
'CREO               : Antonio Salvador Benavides Guardado
'FECHA_CREO         : 22/Abril/2015
'MODIFICÓ:
'FECHA_MODIFICÓ:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************************************
Private Sub Obtener_Horario_Turno_Grid()
'    Call Conectar_Ayudante.Asigna_Item_Combo(Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 1), Cmb_Calendario_Turnos)
    Txt_Calendarios_Configuracion_Turno.Text = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col)
    Dtp_Calendario_Hora_Inicio.Value = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 2)
    Dtp_Calendario_Hora_Termino.Value = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 3)
    Dtp_Calendario_Inicio_Comida.Value = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 4)
    Dtp_Calendario_Termino_Comida.Value = Grid_Calendarios_Configuracion_Turnos.TextMatrix(Grid_Calendarios_Configuracion_Turnos.Row, Grid_Calendarios_Configuracion_Turnos.Col + 5)
End Sub

'*******************************************************************************************************
'NOMBRE_FUNCIÓN     : Calcular_Horas_Calendarios_Turnos
'DESCRIPCIÓN        : Calcula las Horas del turno y la comida
'PARÁMETROS:
'CREO               : Antonio Salvador Benavides Guardado
'FECHA_CREO         : 23/Abril/2015
'MODIFICÓ:
'FECHA_MODIFICÓ:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************************************
Private Sub Calcular_Horas_Calendarios_Turnos()
    'Horas turno
    If Dtp_Calendario_Hora_Inicio.Value <> Dtp_Calendario_Hora_Termino.Value Then
        If ((Val(DateDiff("n", Dtp_Calendario_Hora_Inicio.Value, Dtp_Calendario_Inicio_Comida.Value)) + Val(DateDiff("n", Dtp_Calendario_Termino_Comida.Value, Dtp_Calendario_Hora_Termino.Value))) / 60) > 0 Then
            Txt_Calendario_Horas_Turno.Text = (Val(DateDiff("n", Dtp_Calendario_Hora_Inicio.Value, Dtp_Calendario_Inicio_Comida.Value)) + Val(DateDiff("n", Dtp_Calendario_Termino_Comida.Value, Dtp_Calendario_Hora_Termino.Value))) / 60
        Else
            Txt_Calendario_Horas_Turno.Text = 24 + ((Val(DateDiff("n", Dtp_Calendario_Hora_Inicio.Value, Dtp_Calendario_Inicio_Comida.Value)) + Val(DateDiff("n", Dtp_Calendario_Termino_Comida.Value, Dtp_Calendario_Hora_Termino.Value))) / 60)
        End If
    Else
        Txt_Calendario_Horas_Turno.Text = ""
    End If
    'Horas de comida
    If Dtp_Calendario_Inicio_Comida.Value <> Dtp_Calendario_Termino_Comida.Value Then
        Txt_Calendario_Horas_Comida.Text = Val(DateDiff("n", Dtp_Calendario_Inicio_Comida.Value, Dtp_Calendario_Termino_Comida.Value)) / 60
    Else
        Txt_Calendario_Horas_Comida.Text = ""
    End If
End Sub

Private Sub Guardar_Imagen_Logo()
    On Error GoTo errSub
    
    'si el control Image no tiene una imagen sale de la rutina
    If Logo_Temp.picture = 0 Then
       MsgBox "No se puede guardar. El image debe tener una imagen", vbCritical
    End If
    With CommonDialog1
        If Dir(App.Path & "\Logos_Empresas\" & Txt_Cat_Empresas_Nombre.Text, vbDirectory) = "" Then
            MkDir (App.Path & "\Logos_Empresas\" & Txt_Cat_Empresas_Nombre.Text)
        End If
        SavePicture Logo_Temp, App.Path & "\Logos_Empresas\" & Txt_Cat_Empresas_Nombre.Text & "\" & CommonDialog1.FileTitle
       
    End With
    
    Exit Sub
    
errSub:
    MsgBox Err.Description
End Sub
'´***********************************************Inicio Equipos Almacenes Identificacion*********************************
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Equipos_Almacenes_Identificadores
    'DESCRIPCIÓN:           Consulta los equipos de la base de datos
    'PARÁMETROS :           Nombre: numero  del equipo a buscar
    'CREO       :           Flores Ramirez Yazmin
    'FECHA_CREO :           06 Diciembre 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************

Private Sub Consulta_Cat_Equipos_Almacenes_Identificadores(Nombre As String)
Dim Rs_Consulta_Cat_Equipos_Almacenes_Identificadores As rdoResultset     'Informacion de los Maquinas

Grid_Cat_Equipos_Almacenes.Rows = 0
Grid_Cat_Equipos_Almacenes.Cols = 5
'Consulta todos los roles que se encuentran dados de alta
Mi_SQL = "SELECT Equipo_ID, No_Equipo, Direccion_IP, Puerto_IP, Descripcion "
Mi_SQL = Mi_SQL & " FROM Cat_Equipos_Almacenes_Identificadores"
Mi_SQL = Mi_SQL & " WHERE No_Equipo LIKE '%" & Nombre & "%'"
Mi_SQL = Mi_SQL & " ORDER BY No_Equipo"
Set Rs_Consulta_Cat_Equipos_Almacenes_Identificadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Equipos_Almacenes_Identificadores
    If Not .EOF Then
        Grid_Cat_Equipos_Almacenes.AddItem "Equipo ID" & Chr(9) & "No Equipo" & Chr(9) & "Direccion IP" & Chr(9) & "Puerto_IP" & Chr(9) & "Descripcion"
        While Not .EOF
            Grid_Cat_Equipos_Almacenes.AddItem .rdoColumns("Equipo_ID") & Chr(9) & .rdoColumns("No_Equipo") & Chr(9) & _
                                      .rdoColumns("Direccion_IP") & Chr(9) & .rdoColumns("Puerto_IP") & Chr(9) & _
                                      .rdoColumns("Descripcion")
            .MoveNext
        Wend
        'Asigna los tamaños de las columnas del grid_roles
        Grid_Cat_Equipos_Almacenes.FixedRows = 1
        Grid_Cat_Equipos_Almacenes.ColWidth(0) = 0    'Equipo ID
        Grid_Cat_Equipos_Almacenes.ColWidth(1) = 1500    'Numero equipo
        Grid_Cat_Equipos_Almacenes.ColWidth(2) = 2000    'Direccion IP
        Grid_Cat_Equipos_Almacenes.ColWidth(3) = 0     'Puerto IP
        Grid_Cat_Equipos_Almacenes.ColWidth(4) = 3200     'Descirpcion
    End If
    .Close
End With

Set Rs_Consulta_Cat_Equipos_Almacenes_Identificadores = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Equipos_Almacenes_Identificadores
    'DESCRIPCIÓN:           Realiza el Alta de un equipo en la base de datos
    'PARÁMETROS :
    'CREO       :           Flores Ramirez Yazmin
    'FECHA_CREO :           06 Diciembre 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Alta_Cat_Equipos_Almacenes_Identificadores()
Dim Rs_Alta_Cat_Equipos_Almacenes_Identificadores As rdoResultset 'Informacion del Maquinas

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Maquina
    Set Rs_Alta_Cat_Equipos_Almacenes_Identificadores = Conectar_Ayudante.Recordset_Agregar("Cat_Equipos_Almacenes_Identificadores")
    'Llena la tabla de Cat_Maquina con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Equipos_Almacenes_Identificadores
        .AddNew
            Txt_Cat_Equipos_Almacenes_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Equipos_Almacenes_Identificadores", "Equipo_ID"), "00000")
            .rdoColumns("Equipo_ID") = Trim(Txt_Cat_Equipos_Almacenes_ID.Text)
            .rdoColumns("No_Equipo") = Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text)
            .rdoColumns("Direccion_IP") = Trim(UCase(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text))
            .rdoColumns("Puerto_IP") = Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text)
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Cat_Equipos_Almacenes_Descripcion.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    
    Set Rs_Alta_Cat_Equipos_Almacenes_Identificadores = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Equipos_Almacenes_Generales.Enabled = False
    Fra_Cat_Equipos_Almacenes.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Consultar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    If Grid_Cat_Equipos_Almacenes.Rows = 0 Then
        Grid_Cat_Equipos_Almacenes.Cols = 5
        Grid_Cat_Equipos_Almacenes.AddItem "Equipo ID" & Chr(9) & "No Equipo" & Chr(9) & "Direccion IP" & Chr(9) & "Puerto_IP" & Chr(9) & "Descripcion"
    End If
    'Llena el grid con los datos del nuevo Departamento
    Grid_Cat_Equipos_Almacenes.AddItem Trim(Txt_Cat_Equipos_Almacenes_ID.Text) & Chr(9) & Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text) & Chr(9) & _
                                      Trim(UCase(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text)) & Chr(9) & Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text) & Chr(9) & _
                                      Trim(UCase(Txt_Cat_Equipos_Almacenes_Descripcion.Text))
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Cat_Equipos_Almacenes.FixedRows = 1
    Grid_Cat_Equipos_Almacenes.ColWidth(0) = 0    'Equipo ID
    Grid_Cat_Equipos_Almacenes.ColWidth(1) = 1500    'Numero equipo
    Grid_Cat_Equipos_Almacenes.ColWidth(2) = 2000    'Direccion IP
    Grid_Cat_Equipos_Almacenes.ColWidth(3) = 0     'Puerto IP
    Grid_Cat_Equipos_Almacenes.ColWidth(4) = 3200     'Descirpcion
    Conexion_Base.CommitTrans
    Txt_Cat_Equipos_Almacenes_Direccion_IP.Text = ""
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Almacenes_Identificacion", Me)
    MsgBox "Equipo dado de alta", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
    Debug.Print Err.Description
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Equipos_Almacenes_Identificadores
    'DESCRIPCIÓN:           Realiza la modificacion de un equipo en la base de datos
    'PARÁMETROS :
    'CREO       :           Flores Ramirez Yazmin
    'FECHA_CREO :           06 Diciembre 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Sub Modifica_Cat_Equipos_Almacenes_Identificadores()
Dim Rs_Modifica_Cat_Equipos_Almacenes_Identificadores As rdoResultset 'Informacion del Maquinas
Dim Mi_SQL As String
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM Cat_Equipos_Almacenes_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Equipo_ID = '" & Trim(Txt_Cat_Equipos_Almacenes_ID.Text) & "'"
    
    'Modifica Maquina
    Set Rs_Modifica_Cat_Equipos_Almacenes_Identificadores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Modifica_Cat_Equipos_Almacenes_Identificadores
        .Edit
            .rdoColumns("No_Equipo") = Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text)
            .rdoColumns("Direccion_IP") = Trim(UCase(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text))
            .rdoColumns("Puerto_IP") = Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text)
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Cat_Equipos_Almacenes_Descripcion.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Equipos_Almacenes_Identificadores = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    Fra_Cat_Equipos_Almacenes_Generales.Enabled = False
    Fra_Cat_Equipos_Almacenes.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Consultar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    'Modifica la informacion en el grid
    Grid_Cat_Equipos_Almacenes.TextMatrix(Grid_Cat_Equipos_Almacenes.RowSel, 1) = Val(Txt_Cat_Equipos_Almacenes_No_Equipo.Text)
    Grid_Cat_Equipos_Almacenes.TextMatrix(Grid_Cat_Equipos_Almacenes.RowSel, 2) = Trim(UCase(Txt_Cat_Equipos_Almacenes_Direccion_IP.Text))
    Grid_Cat_Equipos_Almacenes.TextMatrix(Grid_Cat_Equipos_Almacenes.RowSel, 3) = Val(Txt_Cat_Equipos_Almacenes_Puerto_IP.Text)
    Grid_Cat_Equipos_Almacenes.TextMatrix(Grid_Cat_Equipos_Almacenes.RowSel, 4) = Trim(UCase(Txt_Cat_Equipos_Almacenes_Descripcion.Text))
    Conexion_Base.CommitTrans
    Txt_Cat_Equipos_Almacenes_Direccion_IP.Text = ""
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Almacenes_Identificacion", Me)
    MsgBox "Equipo Modificado", vbInformation + vbOKOnly, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er

End Sub
'*****************************************TErmino Equipos Almacenes Identificacion**********************************

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Obtener_Lista_Empleados_Seleccionados
    'DESCRIPCIÓN:           Realiza la modificacion de un equipo en la base de datos
    'PARÁMETROS :
    'CREO       :           Flores Ramirez Yazmin
    'FECHA_CREO :           06 Diciembre 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Private Function Obtener_Lista_Empleados_Seleccionados() As String
Dim Cont_I As Integer
Dim Lista_Empleados As String
    If Lst_Calendarios_Configuracion_Empleados.SelCount > 0 Then
        For Cont_I = 0 To Lst_Calendarios_Configuracion_Empleados.ListCount - 1
            If Lst_Calendarios_Configuracion_Empleados.Selected(Cont_I) Then
                Lista_Empleados = Lista_Empleados & Lst_Calendarios_Configuracion_Empleados.ItemData(Cont_I) & ","
            End If
        Next Cont_I
        If Lista_Empleados Like "*," Then
            Lista_Empleados = Mid(Lista_Empleados, 1, Len(Lista_Empleados) - 1)
        End If
    End If
    Obtener_Lista_Empleados_Seleccionados = Trim(Lista_Empleados)
End Function
