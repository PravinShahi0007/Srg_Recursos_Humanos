VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Cat_Generales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6195
   ClientLeft      =   5505
   ClientTop       =   3240
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7335
   Begin VB.PictureBox Pic_Usuarios 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5160
      Left            =   0
      ScaleHeight     =   5160
      ScaleWidth      =   7380
      TabIndex        =   16
      Top             =   360
      Width           =   7380
      Begin VB.Frame Fra_Usuarios 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   165
         TabIndex        =   23
         Top             =   2865
         Width           =   6990
         Begin MSFlexGridLib.MSFlexGrid Grid_Usuarios 
            Height          =   1920
            Left            =   75
            TabIndex        =   10
            Top             =   240
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   3387
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Usuarios 
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
         Height          =   2835
         Left            =   165
         TabIndex        =   17
         Top             =   15
         Width           =   6990
         Begin VB.ComboBox Cmb_Area_ID 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0000
            Left            =   1080
            List            =   "Frm_Cat_Generales.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   238
            Top             =   2040
            Width           =   2280
         End
         Begin VB.TextBox Txt_No_Nomina 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   4
            Top             =   960
            Width           =   915
         End
         Begin VB.ComboBox Cmb_Roles 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0027
            Left            =   1080
            List            =   "Frm_Cat_Generales.frx":0031
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   960
            Width           =   4920
         End
         Begin VB.TextBox Txt_Contraseña_Confirmar 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4980
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox Txt_Login 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1320
            Width           =   2280
         End
         Begin VB.TextBox Txt_Comentarios_Usuarios 
            Height          =   315
            Left            =   1065
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2400
            Width           =   5835
         End
         Begin VB.TextBox Txt_Contraseña 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1080
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   1680
            Width           =   2280
         End
         Begin VB.TextBox Txt_Nombre_Usuario 
            Height          =   315
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   2
            Top             =   596
            Width           =   5835
         End
         Begin VB.TextBox Txt_Usuario_ID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   255
            Width           =   2280
         End
         Begin VB.ComboBox Cmb_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":004E
            Left            =   4980
            List            =   "Frm_Cat_Generales.frx":0058
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTP_Fecha_Caducar_Usuario 
            Height          =   315
            Left            =   4980
            TabIndex        =   6
            Top             =   1305
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   124387331
            CurrentDate     =   40441
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
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
            Left            =   75
            TabIndex        =   237
            Top             =   2080
            Width           =   405
         End
         Begin VB.Label Lbl_Fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Left            =   3450
            TabIndex        =   222
            Top             =   1365
            Width           =   450
         End
         Begin VB.Label Lbl_Tipo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rol"
            Height          =   195
            Left            =   75
            TabIndex        =   41
            Top             =   1020
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirma Contraseña"
            Height          =   195
            Left            =   3450
            TabIndex        =   40
            Top             =   1740
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus"
            Height          =   195
            Left            =   3450
            TabIndex        =   39
            Top             =   300
            Width           =   525
         End
         Begin VB.Label Lbl_Login 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            Height          =   195
            Left            =   75
            TabIndex        =   22
            Top             =   1350
            Width           =   390
         End
         Begin VB.Label Lbl_Comentarios 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   21
            Top             =   2460
            Width           =   870
         End
         Begin VB.Label Lbl_Contraseña 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña"
            Height          =   195
            Left            =   75
            TabIndex        =   20
            Top             =   1740
            Width           =   810
         End
         Begin VB.Label Lbl_Nombre_Usuario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   75
            TabIndex        =   19
            Top             =   660
            Width           =   555
         End
         Begin VB.Label Lbl_Usuario_ID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario ID"
            Height          =   195
            Left            =   75
            TabIndex        =   18
            Top             =   300
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox Pic_Tipos_Notas_Credito 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7395
      TabIndex        =   210
      Top             =   360
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Frame Fra_Tipos_Notas_Credito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipos de Notas de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   60
         TabIndex        =   219
         Top             =   1485
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Tipos_Notas_Credito 
            Height          =   3120
            Left            =   75
            TabIndex        =   215
            Top             =   225
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5503
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
      Begin VB.Frame Fra_Generales_Tipos_Notas_Credito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Tipos de Notas de Credito"
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
         Height          =   1455
         Left            =   60
         TabIndex        =   211
         Top             =   0
         Width           =   7200
         Begin VB.TextBox Txt_Comentarios_Tipos_Notas_Credito 
            Height          =   315
            Left            =   1125
            MaxLength       =   250
            TabIndex        =   214
            Top             =   1050
            Width           =   5895
         End
         Begin VB.TextBox Txt_Descripcion_Tipos_Notas_Credito 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   213
            Top             =   660
            Width           =   5895
         End
         Begin VB.TextBox Txt_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   212
            Top             =   270
            Width           =   2370
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   218
            Top             =   1110
            Width           =   870
         End
         Begin VB.Label Lbl_Descripcion_Tiempos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripcion"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   217
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo Nota ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   216
            Top             =   330
            Width           =   915
         End
      End
   End
   Begin VB.CommandButton Btn_Buscar 
      Caption         =   "Buscar"
      Height          =   555
      Left            =   4404
      Picture         =   "Frm_Cat_Generales.frx":006E
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "C"
      Top             =   5550
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   30
      Picture         =   "Frm_Cat_Generales.frx":0170
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "A"
      Top             =   5550
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   5865
      Picture         =   "Frm_Cat_Generales.frx":0272
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5550
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   555
      Left            =   2946
      Picture         =   "Frm_Cat_Generales.frx":0374
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "B"
      Top             =   5550
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.PictureBox Pic_Gaps 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7395
      TabIndex        =   94
      Top             =   435
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CommandButton Btn_Ver_Gap 
         Caption         =   "Ver"
         Height          =   315
         Left            =   5835
         TabIndex        =   226
         Tag             =   "C"
         Top             =   300
         Width           =   1320
      End
      Begin VB.Frame Fra_Grid_Gaps 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tripulaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   105
         TabIndex        =   96
         Top             =   1470
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Gaps 
            Height          =   3225
            Left            =   90
            TabIndex        =   101
            Top             =   225
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5689
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
      Begin VB.Frame Fra_Generales_Gaps 
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
         Height          =   1410
         Left            =   105
         TabIndex        =   95
         Top             =   75
         Width           =   7200
         Begin VB.TextBox Txt_Comentarios_Gap 
            Height          =   315
            Left            =   1275
            MaxLength       =   250
            TabIndex        =   100
            Top             =   990
            Width           =   5790
         End
         Begin VB.TextBox Txt_Nombre_Gap 
            Height          =   315
            Left            =   1275
            MaxLength       =   100
            TabIndex        =   99
            Top             =   615
            Width           =   5790
         End
         Begin VB.TextBox Txt_Gap_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label Lbl_Comentarios_Gap 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   103
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Gap 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   102
            Top             =   675
            Width           =   660
         End
         Begin VB.Label Lbl_Ciudad_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tripulación ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   97
            Top             =   300
            Width           =   1200
         End
      End
   End
   Begin VB.PictureBox Pic_Cursos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7350
      TabIndex        =   70
      Top             =   435
      Width           =   7350
      Begin VB.Frame Fra_Generales_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Cursos"
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
         Height          =   2280
         Left            =   45
         TabIndex        =   71
         Top             =   45
         Width           =   7155
         Begin VB.TextBox Txt_Instructor_Curso 
            Height          =   315
            Left            =   1425
            MaxLength       =   100
            TabIndex        =   78
            Top             =   1485
            Width           =   5595
         End
         Begin VB.TextBox Txt_Horas_Curso 
            Height          =   315
            Left            =   1425
            MaxLength       =   250
            TabIndex        =   76
            Top             =   1110
            Width           =   2310
         End
         Begin VB.ComboBox Cmb_Tipo_Curso 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0476
            Left            =   4830
            List            =   "Frm_Cat_Generales.frx":0480
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1110
            Width           =   2205
         End
         Begin VB.TextBox Txt_Comentarios_Curso 
            Height          =   315
            Left            =   1425
            MaxLength       =   250
            TabIndex        =   79
            Top             =   1845
            Width           =   5595
         End
         Begin VB.TextBox Txt_Nombre_Curso 
            Height          =   315
            Left            =   1425
            MaxLength       =   100
            TabIndex        =   75
            Top             =   720
            Width           =   5595
         End
         Begin VB.TextBox Txt_Curso_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1425
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   315
            Width           =   2310
         End
         Begin VB.Label Lbl_Instructor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Instructor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   225
            Top             =   1530
            Width           =   660
         End
         Begin VB.Label Lbl_Horas_Curso 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Horas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   224
            Top             =   1170
            Width           =   510
         End
         Begin VB.Label Lbl_Tipo_Curso 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3825
            TabIndex        =   223
            Top             =   1170
            Width           =   390
         End
         Begin VB.Label Lbl_Comentarios_Curso 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   80
            Top             =   1905
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Curso 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Curso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   74
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Lbl_Curso_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Curso ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   72
            Top             =   375
            Width           =   750
         End
      End
      Begin VB.Frame Fra_Grid_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cursos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Left            =   60
         TabIndex        =   81
         Top             =   2370
         Width           =   7155
         Begin MSFlexGridLib.MSFlexGrid Grid_Cursos 
            Height          =   2340
            Left            =   75
            TabIndex        =   82
            Top             =   240
            Width           =   6990
            _ExtentX        =   12330
            _ExtentY        =   4128
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
   End
   Begin VB.PictureBox Pic_Unidades 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   7395
      TabIndex        =   186
      Top             =   420
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Frame Fra_Grid_Unidades 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   90
         TabIndex        =   196
         Top             =   1395
         Width           =   7245
         Begin MSFlexGridLib.MSFlexGrid Grid_Unidades 
            Height          =   3315
            Left            =   60
            TabIndex        =   197
            Top             =   225
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5847
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Unidades 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Unidades"
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
         Height          =   1350
         Left            =   75
         TabIndex        =   187
         Top             =   45
         Width           =   7245
         Begin VB.TextBox Txt_Comentarios_Unidad 
            Height          =   315
            Left            =   1155
            TabIndex        =   195
            Top             =   960
            Width           =   5970
         End
         Begin VB.TextBox Txt_Nombre_Unidad 
            Height          =   315
            Left            =   1155
            TabIndex        =   194
            Top             =   600
            Width           =   5970
         End
         Begin VB.TextBox Txt_Nombre_Corto_Unidad 
            Height          =   315
            Left            =   5115
            TabIndex        =   193
            Top             =   240
            Width           =   1995
         End
         Begin VB.TextBox Txt_Unidad_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   192
            Top             =   247
            Width           =   1995
         End
         Begin VB.Label Lbl_Comentarios_Unidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   60
            TabIndex        =   191
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Unidad 
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
            Left            =   60
            TabIndex        =   190
            Top             =   645
            Width           =   660
         End
         Begin VB.Label Lbl_Nombre_Corto_Unidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre Corto"
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
            Left            =   3435
            TabIndex        =   189
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label Lbl_Unidad_ID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unidad ID"
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
            Left            =   60
            TabIndex        =   188
            Top             =   300
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox Pic_Vendedores 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5070
      Left            =   0
      ScaleHeight     =   5070
      ScaleWidth      =   7350
      TabIndex        =   114
      Top             =   435
      Width           =   7350
      Begin VB.Frame Fra_Grid_Vendedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vendedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   60
         TabIndex        =   136
         Top             =   2595
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Vendedores 
            Height          =   2085
            Left            =   75
            TabIndex        =   128
            Top             =   240
            Width           =   6990
            _ExtentX        =   12330
            _ExtentY        =   3678
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
      Begin VB.Frame Fra_Generales_Vendedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Vendedores"
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
         Height          =   2565
         Left            =   60
         TabIndex        =   115
         Top             =   30
         Width           =   7185
         Begin VB.TextBox Txt_Comision_Oferta 
            Height          =   315
            Left            =   3060
            MaxLength       =   5
            TabIndex        =   122
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox Txt_Comision_Completa 
            Height          =   315
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   121
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox Txt_Clave_Vendedor 
            Height          =   315
            Left            =   3060
            MaxLength       =   5
            TabIndex        =   118
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox Txt_Comentarios_Vendedor 
            Height          =   315
            Left            =   1200
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   127
            Top             =   2130
            Width           =   5835
         End
         Begin VB.TextBox Txt_Telefono_Vendedor 
            Height          =   315
            Left            =   4995
            MaxLength       =   20
            TabIndex        =   126
            Top             =   1740
            Width           =   2040
         End
         Begin VB.TextBox Txt_Domicilio_Vendedor 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   124
            Top             =   1365
            Width           =   5835
         End
         Begin VB.ComboBox Cmb_Ciudad_Vendedor 
            Height          =   315
            Left            =   1200
            TabIndex        =   125
            Top             =   1740
            Width           =   2925
         End
         Begin VB.TextBox Txt_RFC_Vendedor 
            Height          =   315
            Left            =   4980
            MaxLength       =   20
            TabIndex        =   123
            Top             =   975
            Width           =   2055
         End
         Begin VB.TextBox Txt_Nombre_Vendedor 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   120
            Top             =   600
            Width           =   5835
         End
         Begin VB.ComboBox Cmb_Estatus_Vendedor 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0496
            Left            =   4980
            List            =   "Frm_Cat_Generales.frx":04A0
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   225
            Width           =   2055
         End
         Begin VB.TextBox Txt_Vendedor_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Lbl_Comision_Oferta 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Oferta"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2355
            TabIndex        =   185
            Top             =   1035
            Width           =   435
         End
         Begin VB.Label Lbl_Comision_Complete 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Com. Completa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   184
            Top             =   1035
            Width           =   1065
         End
         Begin VB.Label Lbl_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2355
            TabIndex        =   182
            Top             =   285
            Width           =   405
         End
         Begin VB.Label Lbl_Comentarios_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   135
            Top             =   2205
            Width           =   870
         End
         Begin VB.Label Lbl_Telefono_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Telefono"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4215
            TabIndex        =   134
            Top             =   1800
            Width           =   630
         End
         Begin VB.Label Lbl_Domicilio_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Domicilio"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   133
            Top             =   1425
            Width           =   630
         End
         Begin VB.Label Lbl_Ciudad_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Ciudad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   132
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label Lbl_RFC 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "RFC"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4215
            TabIndex        =   131
            Top             =   1035
            Width           =   315
         End
         Begin VB.Label Lbl_Nombre_Vendedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   130
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Lbl_Estatus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estatus"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4215
            TabIndex        =   129
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Lbl_Vendedor_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Vendedor ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   116
            Top             =   285
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox Pic_Tiempos_Muertos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5130
      Left            =   0
      ScaleHeight     =   5130
      ScaleWidth      =   7395
      TabIndex        =   198
      Top             =   405
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Frame Fra_Generales_Tiempos_Muertos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Tiempos Muertos"
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
         Height          =   1455
         Left            =   60
         TabIndex        =   200
         Top             =   15
         Width           =   7200
         Begin VB.TextBox Txt_Tiempo_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   201
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Descripcion_Tiempos 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   202
            Top             =   600
            Width           =   5895
         End
         Begin VB.TextBox Txt_Comentarios_Tiempos 
            Height          =   315
            Left            =   1125
            MaxLength       =   250
            TabIndex        =   203
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Lbl_Tiempo_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tiempo ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   207
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Lbl_Descripcion_Tiempos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripcion"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   206
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Lbl_Comentarios_Tiempos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   204
            Top             =   960
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Grid_Tiempos_Muertos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tiempos Muertos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   60
         TabIndex        =   199
         Top             =   1485
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Tiempos_Muertos 
            Height          =   3120
            Left            =   75
            TabIndex        =   205
            Top             =   225
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5503
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
   End
   Begin VB.PictureBox Pic_Giros 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   7335
      TabIndex        =   104
      Top             =   420
      Width           =   7365
      Begin VB.Frame Fra_Grid_Giros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipos Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   90
         TabIndex        =   106
         Top             =   1800
         Width           =   7110
         Begin MSFlexGridLib.MSFlexGrid Grid_Giros 
            Height          =   2805
            Left            =   75
            TabIndex        =   113
            Top             =   225
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Giros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Generales Tipos Clientes"
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
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   60
         TabIndex        =   105
         Top             =   75
         Width           =   7140
         Begin VB.TextBox Txt_Comentarios_Giros 
            Height          =   345
            Left            =   1245
            MaxLength       =   250
            TabIndex        =   111
            Top             =   1155
            Width           =   5550
         End
         Begin VB.TextBox Txt_Nombre_Giro 
            Height          =   315
            Left            =   1245
            MaxLength       =   100
            TabIndex        =   110
            Top             =   735
            Width           =   5580
         End
         Begin VB.TextBox Txt_Giro_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Lbl_Comentarios_Giros 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   112
            Top             =   1230
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Giro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   109
            Top             =   795
            Width           =   660
         End
         Begin VB.Label Lbl_Giro_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo Cliente ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   107
            Top             =   360
            Width           =   1050
         End
      End
   End
   Begin VB.PictureBox Pic_Zonas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   7350
      TabIndex        =   227
      Top             =   390
      Width           =   7350
      Begin VB.Frame Fra_Generales_Zonas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Zonas"
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
         Height          =   1620
         Left            =   60
         TabIndex        =   230
         Top             =   0
         Width           =   7185
         Begin VB.TextBox Txt_Zona_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   233
            Top             =   240
            Width           =   2340
         End
         Begin VB.TextBox Txt_Nombre_Zona 
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   232
            Top             =   675
            Width           =   5895
         End
         Begin VB.TextBox Txt_Comentarios_Zona 
            Height          =   315
            Left            =   1125
            MaxLength       =   250
            TabIndex        =   231
            Top             =   1095
            Width           =   5895
         End
         Begin VB.Label Lbl_Zona_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Zona ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   236
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Lbl_Nombre_Zona 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   235
            Top             =   735
            Width           =   660
         End
         Begin VB.Label Lbl_Comentarios_Zona 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   135
            TabIndex        =   234
            Top             =   1170
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Grid_Zonas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zonas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   75
         TabIndex        =   228
         Top             =   1695
         Width           =   7170
         Begin MSFlexGridLib.MSFlexGrid Grid_Zonas 
            Height          =   2895
            Left            =   75
            TabIndex        =   229
            Top             =   225
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5106
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
   End
   Begin VB.PictureBox Pic_Transportes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5070
      Left            =   0
      ScaleHeight     =   5070
      ScaleWidth      =   7365
      TabIndex        =   83
      Top             =   405
      Width           =   7365
      Begin VB.Frame Fra_Grid_Transportes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transportes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   90
         TabIndex        =   85
         Top             =   1830
         Width           =   7155
         Begin MSFlexGridLib.MSFlexGrid Grid_Transportes 
            Height          =   2760
            Left            =   60
            TabIndex        =   93
            Top             =   225
            Width           =   6990
            _ExtentX        =   12330
            _ExtentY        =   4868
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Transportes 
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
         Height          =   1755
         Left            =   75
         TabIndex        =   84
         Top             =   45
         Width           =   7185
         Begin VB.ComboBox Cmb_Zona 
            Height          =   315
            Left            =   1335
            TabIndex        =   90
            Top             =   990
            Width           =   5775
         End
         Begin VB.TextBox Txt_Comentarios_Transporte 
            Height          =   315
            Left            =   1335
            MaxLength       =   250
            TabIndex        =   91
            Top             =   1350
            Width           =   5775
         End
         Begin VB.TextBox Txt_Nombre_Transporte 
            Height          =   315
            Left            =   1335
            MaxLength       =   100
            TabIndex        =   89
            Top             =   615
            Width           =   5775
         End
         Begin VB.TextBox Txt_Transporte_ID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   240
            Width           =   2205
         End
         Begin VB.Label Lbl_Zona 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Zona"
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
            TabIndex        =   183
            Top             =   1050
            Width           =   450
         End
         Begin VB.Label Lbl_Comentarios__Transporte 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   92
            Top             =   1410
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Estado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   88
            Top             =   645
            Width           =   660
         End
         Begin VB.Label Lbl_Transporte_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Transporte ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   86
            Top             =   300
            Width           =   1185
         End
      End
   End
   Begin VB.PictureBox Pic_Operadores 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   7380
      TabIndex        =   168
      Top             =   420
      Width           =   7380
      Begin VB.Frame Fra_Grid_Operadores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Operadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   90
         TabIndex        =   178
         Top             =   1620
         Width           =   7155
         Begin MSFlexGridLib.MSFlexGrid Grid_Operadores 
            Height          =   3000
            Left            =   75
            TabIndex        =   179
            Top             =   240
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   5292
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
      Begin VB.Frame Fra_Generales_Operadores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos Generales"
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
         Height          =   1515
         Left            =   90
         TabIndex        =   169
         Top             =   75
         Width           =   7155
         Begin VB.ComboBox Cmb_Tipo 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":04B7
            Left            =   5535
            List            =   "Frm_Cat_Generales.frx":04C1
            Style           =   2  'Dropdown List
            TabIndex        =   220
            Top             =   300
            Width           =   1500
         End
         Begin VB.TextBox Txt_Comentarios_Operadores 
            Height          =   315
            Left            =   1080
            TabIndex        =   177
            Top             =   1095
            Width           =   5955
         End
         Begin VB.ComboBox Cmb_Estatus_Operador 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":04DA
            Left            =   3240
            List            =   "Frm_Cat_Generales.frx":04E4
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   300
            Width           =   1410
         End
         Begin VB.TextBox Txt_Nombre_Operador 
            Height          =   315
            Left            =   1080
            TabIndex        =   176
            Top             =   705
            Width           =   5955
         End
         Begin VB.TextBox Txt_Operador_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   174
            Top             =   307
            Width           =   1230
         End
         Begin VB.Label Lbl_Tipo_Operador 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo"
            Height          =   195
            Left            =   4875
            TabIndex        =   221
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Lbl_Comentarios_Operadores 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   90
            TabIndex        =   173
            Top             =   1155
            Width           =   870
         End
         Begin VB.Label Lbl_Estatus_Operadores 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   2565
            TabIndex        =   172
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Lbl_Nombre_Operador 
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
            TabIndex        =   171
            Top             =   765
            Width           =   660
         End
         Begin VB.Label Lbl_Operador_ID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Operador ID"
            Height          =   195
            Left            =   90
            TabIndex        =   170
            Top             =   360
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox Pic_Secciones 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   7365
      TabIndex        =   137
      Top             =   420
      Width           =   7365
      Begin VB.Frame Fra_Secciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Secciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   105
         TabIndex        =   139
         Top             =   1845
         Width           =   7080
         Begin MSFlexGridLib.MSFlexGrid Grid_Secciones 
            Height          =   2730
            Left            =   135
            TabIndex        =   145
            Top             =   240
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   4815
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
      Begin VB.Frame Fra_Generales_Secciones 
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
         Height          =   1710
         Left            =   90
         TabIndex        =   138
         Top             =   60
         Width           =   7110
         Begin VB.ComboBox Cmb_Seccion_Supervisor 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":04FA
            Left            =   1110
            List            =   "Frm_Cat_Generales.frx":04FC
            TabIndex        =   143
            Top             =   1185
            Width           =   5850
         End
         Begin VB.TextBox Txt_Seccion_Clave 
            Height          =   315
            Left            =   1095
            TabIndex        =   142
            Top             =   735
            Width           =   5850
         End
         Begin VB.TextBox Txt_Seccion_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1110
            TabIndex        =   141
            Top             =   292
            Width           =   1740
         End
         Begin VB.Label Lbl_Seccion_Supervisor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   135
            TabIndex        =   180
            Top             =   1230
            Width           =   750
         End
         Begin VB.Label Lbl_Seccion_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   144
            Top             =   795
            Width           =   405
         End
         Begin VB.Label Lbl_Seecion_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Seccion ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   140
            Top             =   345
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox Pic_Apl_Cat_Roles 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   7380
      TabIndex        =   24
      Top             =   405
      Visible         =   0   'False
      Width           =   7380
      Begin VB.CommandButton Btn_Acceso_Seguridad 
         Caption         =   "Control de Acceso"
         Height          =   375
         Left            =   5160
         TabIndex        =   27
         Tag             =   "C"
         Top             =   930
         Width           =   2000
      End
      Begin VB.Frame Fra_Generales_Roles 
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
         Height          =   1455
         Left            =   90
         TabIndex        =   28
         Top             =   -15
         Width           =   7140
         Begin VB.TextBox Txt_Comentarios_Rol 
            Height          =   375
            Left            =   1110
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   945
            Width           =   3945
         End
         Begin VB.TextBox Txt_Nombre_Rol 
            Height          =   285
            Left            =   1110
            MaxLength       =   100
            TabIndex        =   30
            Top             =   585
            Width           =   5940
         End
         Begin VB.TextBox Txt_Rol_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1100
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   225
            Width           =   1530
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            Height          =   195
            Left            =   135
            TabIndex        =   34
            Top             =   1035
            Width           =   870
         End
         Begin VB.Label Label28 
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
            Left            =   135
            TabIndex        =   33
            Top             =   630
            Width           =   660
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rol ID"
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
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame Fra_Acceso_Sistema_Rol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accesos del Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Left            =   90
         TabIndex        =   35
         Top             =   1410
         Visible         =   0   'False
         Width           =   7140
         Begin VB.TextBox Txt_Habilitar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7935
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   37
            Top             =   345
            Visible         =   0   'False
            Width           =   500
         End
         Begin VB.CheckBox Chk_Habilitar_Menu_Submenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Habilitar"
            Height          =   200
            Left            =   7605
            TabIndex        =   36
            Top             =   660
            Visible         =   0   'False
            Width           =   900
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Accesos_Seguridad 
            Height          =   3195
            Left            =   105
            TabIndex        =   38
            Top             =   225
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   5636
            _Version        =   393216
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Roles_Sistema 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Roles del Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Left            =   90
         TabIndex        =   25
         Top             =   1425
         Width           =   7140
         Begin MSFlexGridLib.MSFlexGrid Grid_Roles 
            Height          =   3240
            Left            =   90
            TabIndex        =   26
            Top             =   210
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   5715
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
   End
   Begin VB.PictureBox Pic_Cat_Bancos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   45
      ScaleHeight     =   5085
      ScaleWidth      =   7335
      TabIndex        =   43
      Top             =   405
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Frame Fra_Bancos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bancos"
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
         Left            =   180
         TabIndex        =   69
         Top             =   2400
         Width           =   6900
         Begin MSFlexGridLib.MSFlexGrid Grid_Bancos 
            Height          =   2190
            Left            =   135
            TabIndex        =   57
            Top             =   240
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   3863
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Fra_Generales_Bancos 
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
         Height          =   2355
         Left            =   165
         TabIndex        =   44
         Top             =   30
         Width           =   6900
         Begin VB.ComboBox Cmb_Estatus_Banco 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":04FE
            Left            =   4470
            List            =   "Frm_Cat_Generales.frx":0508
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   210
            Width           =   1065
         End
         Begin VB.ComboBox Cmb_Cuenta_Fiscal 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":051A
            Left            =   6120
            List            =   "Frm_Cat_Generales.frx":0524
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   210
            Width           =   645
         End
         Begin VB.ComboBox Cmb_Empresa 
            Height          =   315
            ItemData        =   "Frm_Cat_Generales.frx":0530
            Left            =   4470
            List            =   "Frm_Cat_Generales.frx":053D
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   930
            Width           =   2310
         End
         Begin VB.ComboBox Cmb_Formato 
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   54
            ToolTipText     =   "Formato que Usara para Impresiones."
            Top             =   1605
            Width           =   2310
         End
         Begin VB.TextBox Txt_Sucursal 
            Height          =   315
            Left            =   1050
            MaxLength       =   50
            TabIndex        =   50
            Top             =   930
            Width           =   2310
         End
         Begin VB.TextBox Txt_No_Cuenta_Banco 
            Height          =   315
            Left            =   4470
            MaxLength       =   20
            TabIndex        =   49
            Top             =   570
            Width           =   2300
         End
         Begin VB.TextBox Txt_Contacto_Banco 
            Height          =   315
            Left            =   1050
            MaxLength       =   250
            TabIndex        =   56
            Top             =   1935
            Width           =   2310
         End
         Begin VB.TextBox Txt_Banco_ID 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   225
            Width           =   2310
         End
         Begin VB.TextBox Txt_Nombre_Banco 
            Height          =   315
            Left            =   1050
            MaxLength       =   100
            TabIndex        =   48
            Top             =   570
            Width           =   2300
         End
         Begin VB.TextBox Txt_Ciudad_Banco 
            Height          =   315
            Left            =   1050
            MaxLength       =   50
            TabIndex        =   52
            Top             =   1260
            Width           =   2300
         End
         Begin VB.TextBox Txt_Estado_Banco 
            Height          =   315
            Left            =   4470
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1260
            Width           =   2310
         End
         Begin VB.TextBox Txt_Depostiar_Ah_Banco 
            Height          =   315
            Left            =   4470
            MaxLength       =   50
            TabIndex        =   55
            Top             =   1605
            Width           =   2300
         End
         Begin VB.TextBox Txt_Saldo_Banco 
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4470
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   58
            Top             =   1950
            Width           =   2300
         End
         Begin VB.Label Lbl_Estatus_Banco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Estatus"
            Height          =   195
            Left            =   3420
            TabIndex        =   209
            Top             =   270
            Width           =   525
         End
         Begin VB.Label Lbl_Fiscal 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fiscal"
            Height          =   195
            Left            =   5610
            TabIndex        =   208
            Top             =   270
            Width           =   405
         End
         Begin VB.Label Lbl_Empresa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   3420
            TabIndex        =   181
            Top             =   990
            Width           =   615
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
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
            Left            =   60
            TabIndex        =   68
            Top             =   990
            Width           =   750
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cuenta"
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
            Left            =   3420
            TabIndex        =   67
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contacto"
            Height          =   195
            Left            =   60
            TabIndex        =   66
            Top             =   1995
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco ID"
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
            Left            =   60
            TabIndex        =   65
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
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
            Left            =   60
            TabIndex        =   64
            Top             =   630
            Width           =   555
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   60
            TabIndex        =   63
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            Height          =   195
            Left            =   3420
            TabIndex        =   62
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depositar A"
            Height          =   195
            Left            =   3420
            TabIndex        =   61
            Top             =   1665
            Width           =   825
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formato"
            Height          =   195
            Left            =   60
            TabIndex        =   60
            Top             =   1665
            Width           =   570
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            Height          =   195
            Left            =   3420
            TabIndex        =   59
            Top             =   1995
            Width           =   405
         End
      End
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   555
      Left            =   1488
      Picture         =   "Frm_Cat_Generales.frx":0565
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "M"
      Top             =   5550
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.PictureBox Pic_Gerencias 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   7395
      TabIndex        =   158
      Top             =   420
      Width           =   7395
      Begin VB.Frame Fra_Grid_Gerencias 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gerencias UAP"
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
         Left            =   120
         TabIndex        =   160
         Top             =   1635
         Width           =   7155
         Begin MSFlexGridLib.MSFlexGrid Grid_Gerencias 
            Height          =   3060
            Left            =   105
            TabIndex        =   165
            Top             =   270
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   5398
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
      Begin VB.Frame Fra_Generales_Gerencias 
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
         Height          =   1530
         Left            =   120
         TabIndex        =   159
         Top             =   90
         Width           =   7155
         Begin VB.ComboBox Cmb_Supervisor_Gerencia 
            Height          =   315
            Left            =   1485
            TabIndex        =   164
            Top             =   1095
            Width           =   5520
         End
         Begin VB.TextBox Txt_Nombre_Gerencia 
            Height          =   315
            Left            =   1485
            TabIndex        =   163
            Top             =   705
            Width           =   5500
         End
         Begin VB.TextBox Txt_Gerencia_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   162
            Top             =   337
            Width           =   2175
         End
         Begin VB.Label Lbl_Supervisor_Gerencia 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Supervisor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   167
            Top             =   1155
            Width           =   750
         End
         Begin VB.Label Lbl_Nombre_Gerencia 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Gerencia"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   166
            Top             =   765
            Width           =   645
         End
         Begin VB.Label Lbl_Gerencia_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Gerencia ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   161
            Top             =   390
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Pic_Marcas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   7365
      TabIndex        =   146
      Top             =   420
      Width           =   7365
      Begin VB.Frame Fra_Grid_Marcas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Marcas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   120
         TabIndex        =   148
         Top             =   1875
         Width           =   7140
         Begin MSFlexGridLib.MSFlexGrid Grid_Marcas 
            Height          =   2715
            Left            =   90
            TabIndex        =   157
            Top             =   225
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   4789
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Generales_Marcas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Marcas"
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
         Height          =   1755
         Left            =   105
         TabIndex        =   147
         Top             =   75
         Width           =   7170
         Begin VB.TextBox Txt_Comentarios_Marcas 
            Height          =   315
            Left            =   1020
            MaxLength       =   250
            TabIndex        =   153
            Top             =   1275
            Width           =   6015
         End
         Begin VB.TextBox Txt_Nombre_Corto_Marca 
            Height          =   315
            Left            =   4590
            MaxLength       =   50
            TabIndex        =   151
            Top             =   360
            Width           =   2430
         End
         Begin VB.TextBox Txt_Nombre_Marca 
            Height          =   315
            Left            =   1020
            MaxLength       =   100
            TabIndex        =   152
            Top             =   810
            Width           =   6000
         End
         Begin VB.TextBox Txt_Marca_ID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   360
            Width           =   2400
         End
         Begin VB.Label Lbl_Comentarios_Marcas 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   156
            Top             =   1335
            Width           =   870
         End
         Begin VB.Label Lbl_Nombre_Corto_Marca 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nombre Corto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3510
            TabIndex        =   155
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Lbl_Nombre_Marca 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   154
            Top             =   870
            Width           =   660
         End
         Begin VB.Label Lbl_Marca_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Marca ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   149
            Top             =   420
            Width           =   660
         End
      End
   End
   Begin VB.Label Lbl_Titulo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Titulo del Catalogo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   105
      TabIndex        =   42
      Top             =   15
      Width           =   7155
   End
End
Attribute VB_Name = "Frm_Cat_Generales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mi_Ayudante As New Ayudante

Private Sub Btn_Buscar_Click()
Dim Nombre As String
    
    Select Case Catalogo
        Case "ROLES"
            Nombre = InputBox("Proporcione el nombre del rol a Consultar", "Consulta de Roles")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Roles (Nombre)
        Case "USUARIOS"
            Nombre = InputBox("Proporcione el nombre del Usuario a Consultar", "Consulta de Usuarios")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Usuarios (Nombre)
        Case "CURSOS"
            Nombre = InputBox("Proporcione el nombre del curso a consultar", "Consulta de Cursos")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Cursos (Nombre)
        Case "ZONAS"
            Nombre = InputBox("Proporcione el nombre de la Zona a Consultar", "Consulta de Zonas")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Zonas (Nombre)
        Case "TRANSPORTES"
            Nombre = InputBox("Proporcione el nombre del Transporte a Consultar", "Consulta de Transportes")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Transportes (Nombre)
        Case "GAPS"
            Nombre = InputBox("Proporcione el nombre de la Tripulación a Consultar", "Consulta de Tripulaciones")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Gaps (Nombre)
        Case "SECCIONES"
            Nombre = InputBox("Proporcione el nombre de la Seccion a Consultar", "Consulta de Secciones")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Secciones (Nombre)
        Case "GERENCIAS":
            Nombre = InputBox("Proporcione el nombre de la Gerencia a Consultar", "Consulta de Gerencias")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Gerencia (Nombre)
        
        'Corresponde al catalogo de tipos de clientes
        Case "TIPO_CLIENTE":
            Nombre = InputBox("Proporcione el nombre del Tipo de Cliente a Consultar", "Consulta de Tipos de Clientes")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Giros (Nombre)
        'Corresponde al catalogo de Operadores
        Case "OPERADORES":
            Nombre = InputBox("Proporcione el nombre del operador a Consultar", "Consulta de Operadores")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Operadores (Nombre)
        'Corresponde al catalogo de Marcas
        Case "MARCAS":
            Nombre = InputBox("Proporcione el nombre de la Marca a Consultar", "Consulta de Marcas")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Marcas (Nombre)
        'Corresponde al catalogo de vendedores
        Case "VENDEDORES":
            Nombre = InputBox("Proporcione el nombre del Vendedor a Consultar", "Consulta de Vendedores")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Vendedores (Nombre)
        'Corresponde al catalogo de bancos
        Case "BANCOS"
            Nombre = InputBox("Proporcione el nombre del Banco a Consultar", "Consulta de Bancos")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Bancos (Nombre)
            Cmb_Empresa.ListIndex = -1
        'Corresponde al catalogo de Unidades
        Case "UNIDADES"
            Nombre = InputBox("Proporcione el nombre de la Unidad a Consultar", "Consulta de Unidades")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Unidades (Nombre)
        'Corresponde al catalogo de Tiempos Muertos
        Case "TIEMPOS_MUERTOS"
            Nombre = InputBox("Proporcione la Descripcion del Tiempo Muerto a Consultar", "Consulta de Tiempos Muertos")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Tiempos_Muertos (Nombre)
        'Corresponde al catalogo de Tipos de Notas de Credito
        Case "TIPO_NOTAS_CREDITO"
            Nombre = InputBox("Proporcione la descripcion del tipo de nota de credito a consultar", "Consulta de tipos de notas de credito")
            Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
            Consulta_Tipos_Notas_Credito (Nombre)
    End Select
    Call Conectar_Ayudante.Limpiar_Textos(Me)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Usuarios
    'DESCRIPCIÓN: Consulta todos los Usuarios que hay en la tabla Cat_Usuarios
    '             llenando el Grid
    'PARÁMETROS:
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Usuarios(Optional Nombre As String = "%")
Dim Rs_Consulta_Cat_Usuarios As rdoResultset 'Manejo de registro
    
    'Consulta del usuario
    Mi_SQL = "SELECT Cat_Usuarios.Usuario_ID,Cat_Usuarios.Nombre AS Usuario,Cat_Roles.Nombre AS Rol"
    Mi_SQL = Mi_SQL & " FROM Cat_Usuarios,Cat_Roles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Usuarios.Rol_ID=Cat_Roles.Rol_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Usuarios.Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Usuarios.Nombre"
    Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Usuarios.EOF Then
        'Coloca un encabezado en el grid
        Grid_Usuarios.Rows = 0
        Grid_Usuarios.AddItem "Usuario ID" & Chr(9) & "Nombre" & Chr(9) & "Rol"
        While Not Rs_Consulta_Cat_Usuarios.EOF
            Grid_Usuarios.AddItem Rs_Consulta_Cat_Usuarios.rdoColumns("Usuario_ID") & Chr(9) & Rs_Consulta_Cat_Usuarios.rdoColumns("Usuario") & Chr(9) & Rs_Consulta_Cat_Usuarios.rdoColumns("Rol")
            Grid_Usuarios.FixedRows = 1
            Rs_Consulta_Cat_Usuarios.MoveNext
        Wend
        Grid_Usuarios.ColWidth(0) = 1000
        Grid_Usuarios.ColAlignment(0) = flexAlignCenterCenter
        Grid_Usuarios.ColWidth(1) = 3500
        Grid_Usuarios.ColAlignment(1) = flexAlignLeftCenter
        Grid_Usuarios.ColWidth(2) = 2000
        Grid_Usuarios.ColAlignment(2) = flexAlignLeftCenter
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Usuarios.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Usuarios
'DESCRIPCION: Hace el alta de usuarios en la tabla Cat_Usuarios
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 20-Septiembte-2010
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Usuarios()
Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Usuarios As rdoResultset            'Manejo del registro de Cat_Usuarios, da de alta al usuario
Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Alta de Usuario
    Set Rs_Alta_Cat_Usuarios = Conectar_Ayudante.Recordset_Agregar("Cat_Usuarios")
    'Llena la tabla de Cat_Usuarios con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Usuarios
        .AddNew
            Txt_Usuario_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Usuarios", "Usuario_ID"), "00000")
            .rdoColumns("Usuario_ID") = Txt_Usuario_ID.Text
            .rdoColumns("Rol_ID") = Format(Cmb_Roles.ItemData(Cmb_Roles.ListIndex), "00000")
            .rdoColumns("Area_ID") = Format(Cmb_Area_ID.ItemData(Cmb_Area_ID.ListIndex), "00000")
            .rdoColumns("No_Nomina") = Val(Txt_No_Nomina.Text)
            .rdoColumns("Nombre") = UCase(Txt_Nombre_Usuario.Text)
            .rdoColumns("Estatus") = Cmb_Estatus.Text
            .rdoColumns("Login") = Trim(Txt_Login.Text)
            .rdoColumns("Contraseña") = Trim(Txt_Contraseña.Text)
            .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
            .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Sesion_Abierta") = "NO"
            .rdoColumns("Comentarios") = UCase(Txt_Comentarios_Usuarios.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        'Guarda el password en la tabla
        Mi_SQL = "INSERT INTO Cat_Usuarios_Password (Usuario_ID, Password, Fecha_Password)"
        Mi_SQL = Mi_SQL & " VALUES('" & Trim(Txt_Usuario_ID.Text) & "'"
        Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Contraseña.Text) & "'"
        Mi_SQL = Mi_SQL & " , '" & Format(Now, "MM/dd/yyyy") & "')"
        Conexion_Base.Execute Mi_SQL
    End With
    Rs_Alta_Cat_Usuarios.Close
    'Da de alta los nuevos menus que podrian haber en el sistema
    Fra_Generales_Usuarios.Enabled = False
    Fra_Usuarios.Enabled = True
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    'Pone un encabezado en el grid
    If Grid_Usuarios.Rows = 0 Then
        Grid_Usuarios.AddItem "Usuario" & Chr(9) & "Nombre" & Chr(9) & "Login"
    End If
    'Llena el grid con los datos del nuevo usuario
    Grid_Usuarios.AddItem Txt_Usuario_ID.Text & Chr(9) & UCase(Txt_Nombre_Usuario.Text) & Chr(9) & Cmb_Roles.Text
    Conexion_Base.CommitTrans
    MsgBox "Usuario dado de alta", vbInformation
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Usuarios
'DESCRIPCION: Modifica el usuario en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 20-Septiembre-2010
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Usuarios()
Dim Mi_SQL As String                                'Obtiene los valores de la consulta
Dim Rs_Modificacion_Cat_Usuarios As rdoResultset    'Manejo de registro de la tabla Cat_Usuarios
Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de la tabla Seguridad_Sistema
Dim Rs_Consulta_Passwords_Anteriores As rdoResultset
Dim Menus As Integer                                'Contador para indicar en que numero de registro esta
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & Txt_Usuario_ID.Text & "'"
    Set Rs_Modificacion_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Usuarios.EOF Then
        'Modifica los datos de la tabla Cat_Usuarios
        With Rs_Modificacion_Cat_Usuarios
            .Edit
                .rdoColumns("Rol_ID") = Format(Cmb_Roles.ItemData(Cmb_Roles.ListIndex), "00000")
                .rdoColumns("Area_ID") = Format(Cmb_Area_ID.ItemData(Cmb_Area_ID.ListIndex), "00000")
                .rdoColumns("No_Nomina") = Val(Txt_No_Nomina.Text)
                .rdoColumns("Nombre") = UCase(Txt_Nombre_Usuario.Text)
                If .rdoColumns("Login") <> Trim(Txt_Login.Text) Then
                    .rdoColumns("Login") = Trim(Txt_Login.Text)
                End If
                If .rdoColumns("Contraseña") <> Trim(Txt_Contraseña.Text) Then
                    'Verifica que el password no sea el mismo que ya se tenia dado de alta
                    Mi_SQL = "SELECT TOP " & Historico_Password & " *"
                    Mi_SQL = Mi_SQL & " FROM Cat_Usuarios_Password"
                    Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & .rdoColumns("Usuario_ID") & "'"
                    Mi_SQL = Mi_SQL & " ORDER BY No_Partida DESC"
                    Set Rs_Consulta_Passwords_Anteriores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    While Not Rs_Consulta_Passwords_Anteriores.EOF
                        If Rs_Consulta_Passwords_Anteriores.rdoColumns("Password") = Trim(Txt_Contraseña.Text) Then
                            MsgBox "El nuevo password no puede ser el mismo que ha usado en las ultimas " & Historico_Password & " ocasiones", vbCritical
                            Rs_Consulta_Passwords_Anteriores.Close
                            Conexion_Base.RollbackTrans
                            Exit Sub
                        End If
                        Rs_Consulta_Passwords_Anteriores.MoveNext
                    Wend
                    Rs_Consulta_Passwords_Anteriores.Close
                    .rdoColumns("Contraseña") = Trim(Txt_Contraseña.Text)
                    .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
                    DTP_Fecha_Caducar_Usuario.Value = DateAdd("d", Dias_Caducidad_Contraseñas, Now)
                    .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
                    'Guarda el password en la tabla
                    Mi_SQL = "INSERT INTO Cat_Usuarios_Password (Usuario_ID, Password, Fecha_Password)"
                    Mi_SQL = Mi_SQL & " VALUES('" & Trim(Txt_Usuario_ID.Text) & "'"
                    Mi_SQL = Mi_SQL & " , '" & Trim(Txt_Contraseña.Text) & "'"
                    Mi_SQL = Mi_SQL & " , '" & Format(Now, "MM/dd/yyyy") & "')"
                    Conexion_Base.Execute Mi_SQL
                Else
                    .rdoColumns("Fecha_Caduca") = Format(DTP_Fecha_Caducar_Usuario.Value, "MM/dd/yyyy")
                End If
                .rdoColumns("Estatus") = Trim(Cmb_Estatus.Text)
                .rdoColumns("Sesion_Abierta") = "NO"
                .rdoColumns("Comentarios") = UCase(Trim(Txt_Comentarios_Usuarios.Text))
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Usuarios.Close
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Usuarios.Enabled = False
    Fra_Usuarios.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    Conexion_Base.CommitTrans
    'Hace la consulta de usuarios
    Consulta_Usuarios
    MsgBox "El usuario ha sido modificado", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Curso
'DESCRIPCION: Modifica el curso en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Modifica_Curso()
Dim Rs_Modificacion_Cat_Cursos As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta el curso a modificar
    Mi_SQL = "SELECT * FROM Cat_Cursos"
    Mi_SQL = Mi_SQL & "  WHERE Curso_ID='" & Trim(Txt_Curso_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Cursos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Cursos.EOF Then
        'Modifica los datos de la tabla Cat_Cursos
        With Rs_Modificacion_Cat_Cursos
            .Edit
                .rdoColumns("Curso_ID") = Txt_Curso_ID.Text
                .rdoColumns("Nombre") = Trim(Txt_Nombre_Curso.Text)
                .rdoColumns("Horas") = Val(Txt_Horas_Curso.Text)
                .rdoColumns("Tipo") = Cmb_Tipo_Curso.Text
                .rdoColumns("Instructor") = Trim(Txt_Instructor_Curso.Text)
                .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Curso.Text)
                .rdoColumns("Usuario_Modifico") = Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Cursos.Close
    MsgBox "El curso ha sido modificado", vbInformation
    'Configura el grid
    Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 1) = Trim(Txt_Nombre_Curso.Text)
    Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 2) = Val(Txt_Horas_Curso.Text)
    Btn_Salir_Click
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Unidades
    'DESCRIPCIÓN:Modifica la Unidad
    'PARÁMETROS:Generales de Unidades
    'CREO:Rafael Muñoz
    'FECHA_CREO:04-Sep_2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Modifica_Unidades()
Dim Rs_Modifica_Cat_Unidades As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta la unidad actual seleccionada
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Unidades"
    Mi_SQL = Mi_SQL & "  WHERE Unidad_ID ='" & Txt_Unidad_ID.Text & "'"
    Set Rs_Modifica_Cat_Unidades = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Cursos
    With Rs_Modifica_Cat_Unidades
        .Edit
            .rdoColumns("Unidad_ID") = Txt_Unidad_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Unidad.Text))
            .rdoColumns("Nombre_Corto") = Trim(UCase(Txt_Nombre_Corto_Unidad.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Unidad.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Unidades = Nothing
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Unidades.Enabled = False
    Fra_Grid_Unidades.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    'Configura el grid
    Grid_Unidades.TextMatrix(Grid_Unidades.RowSel, 0) = Txt_Unidad_ID.Text
    Grid_Unidades.TextMatrix(Grid_Unidades.RowSel, 1) = Trim(UCase(Txt_Nombre_Unidad.Text))
    Grid_Unidades.TextMatrix(Grid_Unidades.RowSel, 2) = Trim(UCase(Txt_Nombre_Corto_Unidad.Text))
    Conexion_Base.CommitTrans
    MsgBox "La Unidad ha sido modificada", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Transportes
'DESCRIPCION: Modifica el Transporte
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO :08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Transportes()
Dim Rs_Modificacion_Cat_Transportes As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta el Transporte actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Transportes"
    Mi_SQL = Mi_SQL & " WHERE Transporte_ID='" & Txt_Transporte_ID.Text & "'"
    Set Rs_Modificacion_Cat_Transportes = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Transportes.EOF Then
        With Rs_Modificacion_Cat_Transportes
            .Edit
                If Cmb_Zona.ListIndex > -1 Then
                    .rdoColumns("Zona_ID") = Format(Cmb_Zona.ItemData(Cmb_Zona.ListIndex), "00000")
                End If
                .rdoColumns("Nombre") = Trim(Txt_Nombre_Transporte.Text)
                .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Transporte.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Transportes.Close
    MsgBox "El Transporte ha sido modificado", vbInformation
    'Configura el grid
    Grid_Transportes.TextMatrix(Grid_Transportes.RowSel, 0) = Txt_Transporte_ID.Text
    Grid_Transportes.TextMatrix(Grid_Transportes.RowSel, 1) = Trim(Txt_Nombre_Transporte.Text)
    Grid_Transportes.TextMatrix(Grid_Transportes.RowSel, 2) = Trim(Cmb_Zona.Text)
    Btn_Salir_Click
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Gerencia
'DESCRIPCION: Modifica la gerencia
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Marzo-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Gerencia()
Dim Mi_SQL As String
Dim Rs_Modificacion_Cat_Gerencia As rdoResultset    'Manejo de registro
Dim Rs_Consulta_Supervisores As rdoResultset

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta la Gerencia a modificar
    Mi_SQL = "SELECT * FROM Cat_Gerencias"
    Mi_SQL = Mi_SQL & " WHERE Gerencia_ID='" & Txt_Gerencia_ID.Text & "'"
    Set Rs_Modificacion_Cat_Gerencia = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Gerencia.EOF Then
        'Modifica los datos de la tabla Cat_Gerencias
        With Rs_Modificacion_Cat_Gerencia
            .Edit
                .rdoColumns("Nombre") = Trim(Txt_Nombre_Gerencia.Text)
                .rdoColumns("Supervisor_ID") = Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000")
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Gerencia.Close
    'Actualiza los empleados con la gerencia del supervisor
    Mi_SQL = "UPDATE Cat_Empleados SET Gerencia_UAP='" & Trim(Txt_Gerencia_ID.Text) & "'"
    Mi_SQL = Mi_SQL & " WHERE (Empleado_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "')"
    Conexion_Base.Execute Mi_SQL
    'Actualiza los empleados con la gerencia del supervisor de los siguientes niveles
    Mi_SQL = "SELECT DISTINCT Empleado_ID FROM Cat_Empleados WHERE Supervisor_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "'"
    Set Rs_Consulta_Supervisores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Supervisores.EOF
        Mi_SQL = "UPDATE Cat_Empleados SET Gerencia_UAP='" & Trim(Txt_Gerencia_ID.Text) & "'"
        Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Rs_Consulta_Supervisores.rdoColumns("Empleado_ID") & "'"
        Conexion_Base.Execute Mi_SQL
        Rs_Consulta_Supervisores.MoveNext
    Wend
    Rs_Consulta_Supervisores.Close
    Conexion_Base.CommitTrans
    MsgBox "La gerencia ha sido modificada", vbInformation
    Btn_Salir_Click
    Consulta_Gerencia ""
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Modifica_Marcas
    'DESCRIPCION: Modifica la Marca
    'PARAMETROS :
    'CREO       :Sergio Ulises Durán Hernández
    'FECHA_CREO :22-Agosto-2009
    'MODIFICO   :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'*******************************************************************************
Public Sub Modifica_Marcas()
Dim Rs_Modificacion_Cat_Marcas As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta la Marca actual seleccionada
    Mi_SQL = "SELECT * FROM Cat_Marcas"
    Mi_SQL = Mi_SQL & " WHERE Marca_ID='" & Txt_Marca_ID.Text & "'"
    Set Rs_Modificacion_Cat_Marcas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Marcas
    With Rs_Modificacion_Cat_Marcas
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Marca.Text))
            .rdoColumns("Nombre_Corto") = Trim(UCase(Txt_Nombre_Corto_Marca.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Marcas.Text))
            .rdoColumns("Usuario_Modifico") = Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
    End With
    Rs_Modificacion_Cat_Marcas.Close
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Marcas.Enabled = False
    Fra_Grid_Marcas.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    'Configura el grid
    Grid_Marcas.TextMatrix(Grid_Marcas.RowSel, 0) = Txt_Marca_ID.Text
    Grid_Marcas.TextMatrix(Grid_Marcas.RowSel, 1) = Trim(UCase(Txt_Nombre_Marca.Text))
    Grid_Marcas.TextMatrix(Grid_Marcas.RowSel, 2) = Trim(UCase(Txt_Nombre_Corto_Marca.Text))
    Grid_Marcas.TextMatrix(Grid_Marcas.RowSel, 3) = Trim(UCase(Txt_Comentarios_Marcas.Text))
    Conexion_Base.CommitTrans
    MsgBox "La marca ha sido modificada", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Secciones
'DESCRIPCION: Modifica la Seccion en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 28-Septiembre-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Secciones()
Dim Rs_Modificacion_Cat_Secciones As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta el Seccion actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Secciones"
    Mi_SQL = Mi_SQL & " WHERE Seccion_ID='" & Trim(Txt_Seccion_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Secciones = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Secciones.EOF Then
        'Modifica los datos de la tabla Cat_Secciones_Envio
        With Rs_Modificacion_Cat_Secciones
            .Edit
                .rdoColumns("Supervisor_ID") = Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000")
                .rdoColumns("Clave") = Trim(Txt_Seccion_Clave.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Secciones.Close
    'Actualiza los empleados con la sección del supervisor
    Mi_SQL = "UPDATE Cat_Empleados"
    Mi_SQL = Mi_SQL & " SET Nomipaq_ID='" & Trim(Txt_Seccion_Clave.Text) & "'"
    Mi_SQL = Mi_SQL & " WHERE (Empleado_ID='" & Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000") & "')"
    Conexion_Base.Execute Mi_SQL
    MsgBox "La Seccion ha sido modificada", vbInformation
    Consulta_Secciones ("")
    Btn_Salir_Click
Exit Sub
HANDLER:    'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Operadores
    'DESCRIPCIÓN: Modifica el Operador
    'PARÁMETROS:Generales de Operadores
    'CREO:Rafael Muñoz
    'FECHA_CREO:07-Feb_2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Modifica_Operadores()
Dim Rs_Modificacion_Cat_Operadores As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta el Operador actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Operadores"
    Mi_SQL = Mi_SQL & "  WHERE Operador_ID='" & Txt_Operador_ID.Text & "'"
    Set Rs_Modificacion_Cat_Operadores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Operadores
    With Rs_Modificacion_Cat_Operadores
        .Edit
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Operador.Text))
            .rdoColumns("Tipo") = Cmb_Tipo.Text
            .rdoColumns("Estatus") = Mid(Cmb_Estatus_Operador, 1, 1)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Operadores.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Operadores = Nothing
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Operadores.Enabled = False
    Fra_Grid_Operadores.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    'Configura el grid
    Grid_Operadores.TextMatrix(Grid_Operadores.RowSel, 1) = Trim(UCase(Txt_Nombre_Operador.Text))
    Grid_Operadores.TextMatrix(Grid_Operadores.RowSel, 2) = Cmb_Tipo.Text
    Conexion_Base.CommitTrans
    'Hace la consulta de Operadores
    Consulta_Operadores ("")
    MsgBox "El Operador ha sido modificado", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Zonas
'DESCRIPCION: Modifica la Zona
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Zonas()
Dim Rs_Modificacion_Cat_Zonas As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta la Zona actual seleccionada
    Mi_SQL = "SELECT * FROM Cat_Zonas"
    Mi_SQL = Mi_SQL & "  WHERE Zona_ID='" & Txt_Zona_ID.Text & "'"
    Set Rs_Modificacion_Cat_Zonas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Zonas.EOF Then
        'Modifica los datos de la tabla Cat_Zonas
        With Rs_Modificacion_Cat_Zonas
            .Edit
                .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Zona.Text))
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Zona.Text))
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Zonas.Close
    MsgBox "La Zona ha sido modificada", vbInformation
    'Configura el grid
    Grid_Zonas.TextMatrix(Grid_Zonas.RowSel, 0) = Txt_Zona_ID.Text
    Grid_Zonas.TextMatrix(Grid_Zonas.RowSel, 1) = Trim(UCase(Txt_Nombre_Zona.Text))
    Grid_Zonas.TextMatrix(Grid_Zonas.RowSel, 2) = Trim(UCase(Txt_Comentarios_Zona.Text))
    Btn_Salir_Click
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Giros
    'DESCRIPCIÓN: Modifica el Giro_Empresarial
    'PARÁMETROS:Generales de Giros
    'CREO:Rafael Muñoz
    'FECHA_CREO:08-Ene_2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Modifica_Giros()
Dim Rs_Modificacion_Cat_Giros_Empresariales As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Giro actual seleccionado
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Giros_Empresariales"
    Mi_SQL = Mi_SQL & "  WHERE Giro_ID ='" & Txt_Giro_ID.Text & "'"
    Set Rs_Modificacion_Cat_Giros_Empresariales = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Giros_Empresariales
    With Rs_Modificacion_Cat_Giros_Empresariales
        .Edit
            .rdoColumns("Giro_ID") = Txt_Giro_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Giro.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Giros.Text))
            .rdoColumns("Usuario_Modifico") = Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Giros_Empresariales = Nothing
    
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Giros.Enabled = False
    Fra_Grid_Giros.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    'Configura el grid
    Grid_Giros.TextMatrix(Grid_Giros.RowSel, 0) = Txt_Giro_ID.Text
    Grid_Giros.TextMatrix(Grid_Giros.RowSel, 1) = Trim(UCase(Txt_Nombre_Giro.Text))
    Grid_Giros.TextMatrix(Grid_Giros.RowSel, 2) = Trim(UCase(Txt_Comentarios_Giros.Text))
    Conexion_Base.CommitTrans
    'Hace la consulta de giros
    Consulta_Giros ("")
    MsgBox "El Tipo de Cliente ha sido modificado", vbInformation
    Exit Sub

'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Gaps
'DESCRIPCION: Modifica los datos del Gap
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Gaps()
Dim Rs_Modificacion_Cat_Gaps As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta la Gap actual seleccionada
    Mi_SQL = "SELECT * FROM Cat_Gaps"
    Mi_SQL = Mi_SQL & " WHERE Gap_ID='" & Txt_Gap_ID.Text & "'"
    Set Rs_Modificacion_Cat_Gaps = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modificacion_Cat_Gaps.EOF Then
        'Modifica los datos de la tabla Cat_Gaps
        With Rs_Modificacion_Cat_Gaps
            .Edit
                .rdoColumns("Nombre") = Trim(Txt_Nombre_Gap.Text)
                .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Gap.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modificacion_Cat_Gaps.Close
    MsgBox "La tripulación ha sido modificada", vbInformation
    'Configura el grid
    Grid_Gaps.TextMatrix(Grid_Gaps.RowSel, 1) = Trim(Txt_Nombre_Gap.Text)
    Grid_Gaps.TextMatrix(Grid_Gaps.RowSel, 2) = Trim(Txt_Comentarios_Gap.Text)
    Btn_Salir_Click
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Vendedores
    'DESCRIPCIÓN: Modifica el Vendedor
    'PARÁMETROS:
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 14-Marzo-2008
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Modifica_Vendedores()
Dim Mi_SQL As String
Dim Rs_Modificacion_Cat_Vendedores As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Vendedor actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Vendedor_ID='" & Txt_Vendedor_ID.Text & "'"
    Set Rs_Modificacion_Cat_Vendedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Vendedores
    With Rs_Modificacion_Cat_Vendedores
        .Edit
            .rdoColumns("Estatus") = Mid(Cmb_Estatus_Vendedor.Text, 1, 1)
            .rdoColumns("Clave") = Trim(UCase(Txt_Clave_Vendedor.Text))
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Vendedor.Text))
            If Cmb_Gap_Vendedor.ListIndex > -1 Then
                .rdoColumns("Gap_ID") = Format(Cmb_Gap_Vendedor.ItemData(Cmb_Gap_Vendedor.ListIndex), "00000")
            End If
            .rdoColumns("RFC") = Trim(UCase(Txt_RFC_Vendedor.Text))
            .rdoColumns("Domicilio") = Trim(UCase(Txt_Domicilio_Vendedor.Text))
            .rdoColumns("Telefono_1") = Trim(Txt_Telefono_Vendedor.Text)
            .rdoColumns("Comision_Completa") = Val(Txt_Comision_Completa.Text)
            .rdoColumns("Comision_Promocion") = Val(Txt_Comision_Oferta.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Vendedor.Text))
            .rdoColumns("Usuario_Modifico") = Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Vendedores = Nothing
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Vendedores.Enabled = False
    Fra_Grid_Vendedores.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    'Actualiza el grid
    Grid_Vendedores.TextMatrix(Grid_Vendedores.RowSel, 0) = Txt_Vendedor_ID.Text
    Grid_Vendedores.TextMatrix(Grid_Vendedores.RowSel, 1) = Trim(UCase(Txt_Nombre_Vendedor.Text))
    Grid_Vendedores.TextMatrix(Grid_Vendedores.RowSel, 2) = Trim(UCase(Txt_Clave_Vendedor.Text))
    Conexion_Base.CommitTrans
    MsgBox "El Vendedor ha sido modificado", vbInformation
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Eliminar_Click()
    Select Case Catalogo
        Case "ROLES"
            'si el usuario selecciono un rol entonces elimina el rol que fue seleccionado
            If Trim(Txt_Rol_ID.Text) <> "" Then
                If Mensaje("¿Esta seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Seguridad_Sistema", "Rol_ID", Txt_Rol_ID.Text) = True Then
                        If Conectar_Ayudante.Elimina_Catalogo("Cat_Roles", "Rol_ID", Txt_Rol_ID.Text) = True Then
                            Grid_Accesos_Seguridad.Rows = 0
                            'Elimina el rol del grid
                            If Grid_Roles.Rows = 2 Then
                                Grid_Roles.Rows = 0
                            Else
                                Grid_Roles.RemoveItem Grid_Roles.RowSel
                            End If 'Grid_Roles
                            Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia los controles de la forma
                            Fra_Acceso_Sistema_Rol.Visible = False
                            Fra_Roles_Sistema.Visible = True
                            Btn_Acceso_Seguridad.Caption = "Control de Acceso"
                            MsgBox "Rol eliminado", vbInformation
                        End If 'Rol
                    End If 'Acceso
                End If
            'Si no selecciono manda un mensaje al usuario
            Else
                MsgBox "Seleccione un rol para poder eliminar", vbInformation
            End If
        Case "USUARIOS"
            MsgBox "No se puede eliminar usuarios" & Chr(13) & "  Puede cambiar el estatus" & Chr(13) & "   del usuario a INACTIVO", vbInformation
            Fra_Usuarios.Enabled = True
        Case "CURSOS"
            If Trim(Txt_Curso_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Cursos", "Curso_ID", Txt_Curso_ID.Text) = True Then
                        If Grid_Cursos.Rows = 2 Then
                            Grid_Cursos.FixedRows = 0
                            Grid_Cursos.RemoveItem Grid_Cursos.RowSel + 1
                        Else
                            Grid_Cursos.RemoveItem Grid_Cursos.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "El curso ha sido eliminado", vbInformation
                    End If
                End If
            Else
                 MsgBox "No hay datos que eliminar", vbExclamation
            End If
        Case "ZONAS"
            If Txt_Zona_ID.Text <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Zonas", "Zona_ID", Txt_Zona_ID.Text) = True Then
                        If Grid_Zonas.Rows = 2 Then
                            Grid_Zonas.FixedRows = 0
                            Grid_Zonas.RemoveItem Grid_Zonas.RowSel + 1
                        Else
                            Grid_Zonas.RemoveItem Grid_Zonas.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Zona Eliminada", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        Case "TRANSPORTES"
            If Trim(Txt_Transporte_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Transportes", "Transporte_ID", Txt_Transporte_ID.Text) = True Then
                        If Grid_Transportes.Rows = 2 Then
                            Grid_Transportes.FixedRows = 0
                            Grid_Transportes.RemoveItem Grid_Transportes.RowSel + 1
                        Else
                            Grid_Transportes.RemoveItem Grid_Transportes.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Transporte Eliminado", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        Case "GAPS"
            If Txt_Gap_ID.Text <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Gaps", "Gap_ID", Txt_Gap_ID.Text) = True Then
                        If Grid_Gaps.Rows = 2 Then
                            Grid_Gaps.FixedRows = 0
                            Grid_Gaps.RemoveItem Grid_Gaps.RowSel + 1
                        Else
                            Grid_Gaps.RemoveItem Grid_Gaps.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "La tripulación ha sido eliminada", vbInformation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        Case "SECCIONES"
            If Trim(Txt_Seccion_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Secciones", "Seccion_ID", Txt_Seccion_ID.Text) = True Then
                        If Grid_Secciones.Rows = 2 Then
                            Grid_Secciones.FixedRows = 0
                            Grid_Secciones.RemoveItem Grid_Secciones.RowSel + 1
                        Else
                            Grid_Secciones.RemoveItem Grid_Secciones.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "La Seccion ha sido eliminada", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        Case "GERENCIAS"
            If Txt_Gerencia_ID.Text <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Gerencias", "Gerencia_ID", Txt_Gerencia_ID.Text) = True Then
                        If Grid_Gerencias.Rows = 2 Then
                            Grid_Gerencias.FixedRows = 0
                            Grid_Gerencias.RemoveItem Grid_Gerencias.RowSel + 1
                        Else
                            Grid_Gerencias.RemoveItem Grid_Gerencias.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "La Gerencia ha sido Eliminada", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        
        'Corresponde al catalogo de bancos
        Case "BANCOS"
            '#  SI LA RESPUESTA ES SI
            If Mensaje("¿Esta seguro de eliminar el registro?", 3) = vbYes Then
                'SI SE SELECCIONO EL BANCO
                If Trim(Txt_Banco_ID.Text) <> "" Then
                    '#  ELIMINA EL REGISTRO DE LA TABLA Cat_Gastos
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_bancos", "banco_ID", Txt_Banco_ID.Text) = True Then
                        'Quita los datos del banco contenidos en el Grid
                        If Grid_Bancos.Rows = 2 Then
                            Grid_Bancos.FixedRows = 0
                            Grid_Bancos.RemoveItem Grid_Bancos.RowSel + 1
                        Else
                            Grid_Bancos.RemoveItem Grid_Bancos.RowSel
                        End If 'Grid
                        'Limpia los textos de la forma
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        Mensaje "Banco Eliminado"
                    End If 'Eliminar
                Else
                    Call Mensaje("No hay datos que eliminar", 2)
                End If 'Txt_banco_ID.text
            End If
        'Corresponde al catalogo de Tipos de clientes
        Case "TIPO_CLIENTE"
            If Trim(Txt_Giro_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Giros_Empresariales", "Giro_ID", Txt_Giro_ID.Text) = True Then
                        If Grid_Giros.Rows = 2 Then
                            Grid_Giros.FixedRows = 0
                            Grid_Giros.RemoveItem Grid_Giros.RowSel + 1
                        Else
                            Grid_Giros.RemoveItem Grid_Giros.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Tipo de Cliente Eliminado", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de Operadores
        Case "OPERADORES"
            If Txt_Operador_ID.Text <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Operadores", "Operador_ID", Txt_Operador_ID.Text) = True Then
                        If Grid_Operadores.Rows = 2 Then
                            Grid_Operadores.FixedRows = 0
                            Grid_Operadores.RemoveItem Grid_Operadores.RowSel + 1
                        Else
                            Grid_Operadores.RemoveItem Grid_Operadores.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Operador Eliminado", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de Marcas
        Case "MARCAS"
            If Txt_Marca_ID.Text <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Marcas", "Marca_ID", Txt_Marca_ID.Text) = True Then
                        If Grid_Marcas.Rows = 2 Then
                            Grid_Marcas.FixedRows = 0
                            Grid_Marcas.RemoveItem Grid_Marcas.RowSel + 1
                        Else
                            Grid_Marcas.RemoveItem Grid_Marcas.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "La marca ha sido eliminada", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de vendedores
        Case "VENDEDORES"
            If Trim(Txt_Vendedor_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Vendedores", "Vendedor_ID", Txt_Vendedor_ID.Text) = True Then
                        If Grid_Vendedores.Rows = 2 Then
                            Grid_Vendedores.FixedRows = 0
                            Grid_Vendedores.RemoveItem Grid_Vendedores.RowSel + 1
                        Else
                            Grid_Vendedores.RemoveItem Grid_Vendedores.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Vendedor Eliminado", vbExclamation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de Unidades
        Case "UNIDADES"
            If Trim(Txt_Unidad_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Unidades", "Unidad_ID", Txt_Unidad_ID.Text) = True Then
                        If Grid_Unidades.Rows = 2 Then
                            Grid_Unidades.FixedRows = 0
                            Grid_Unidades.RemoveItem Grid_Unidades.RowSel + 1
                        Else
                            Grid_Unidades.RemoveItem Grid_Unidades.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Unidad Eliminada", vbInformation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de Tiempos Muertos
        Case "TIEMPOS_MUERTOS"
            If Trim(Txt_Tiempo_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Tiempos_Muertos", "Tiempo_ID", Txt_Tiempo_ID.Text) = True Then
                        If Grid_Tiempos_Muertos.Rows = 2 Then
                            Grid_Tiempos_Muertos.FixedRows = 0
                            Grid_Tiempos_Muertos.RemoveItem Grid_Tiempos_Muertos.RowSel + 1
                        Else
                            Grid_Tiempos_Muertos.RemoveItem Grid_Tiempos_Muertos.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "Tiempo Muerto Eliminado", vbInformation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
        'Corresponde al catalogo de Tipos de Notas de Crédito
        Case "TIPO_NOTAS_CREDITO"
            If Trim(Txt_Tipo_Nota_Credito_ID.Text) <> "" Then
                If Mensaje("¿Está seguro de eliminar el registro?", 3) = 6 Then
                    If Conectar_Ayudante.Elimina_Catalogo("Cat_Tipos_Notas_Credito", "Tipo_Nota_Credito_ID", Txt_Tipo_Nota_Credito_ID.Text) = True Then
                        If Grid_Tipos_Notas_Credito.Rows = 2 Then
                            Grid_Tipos_Notas_Credito.FixedRows = 0
                            Grid_Tipos_Notas_Credito.RemoveItem Grid_Tipos_Notas_Credito.RowSel + 1
                        Else
                            Grid_Tipos_Notas_Credito.RemoveItem Grid_Tipos_Notas_Credito.RowSel
                        End If
                        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
                        MsgBox "El tipo de nota de crédito ha sido eliminado", vbInformation
                    End If
                End If
            Else
                MsgBox "No hay datos que eliminar", vbExclamation
            End If
    End Select
End Sub

Private Sub Btn_Modificar_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Usuarios As rdoResultset

    '#  Habilita el form para modificarse
    If Btn_Modificar.Caption = "Modificar" Then
        Select Case Catalogo
            Case "ROLES"    'Corresponde al catalogo de roles
                If Trim(Txt_Rol_ID.Text) <> "" Then
                    Btn_Acceso_Seguridad.Visible = False
                    Fra_Acceso_Sistema_Rol.Visible = True
                    Fra_Generales_Roles.Enabled = True
                    Fra_Roles_Sistema.Visible = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Salir.Caption = "Regresar"
                    Btn_Nuevo.Enabled = False
                    Btn_Eliminar.Enabled = False
                    Btn_Buscar.Enabled = False
                    Txt_Nombre_Rol.SetFocus
                    SendKeys "{Home}+{End}"
                Else
                    MsgBox "Debe seleccionar un rol para poder modificar", vbInformation
                    Exit Sub
                End If
            Case "USUARIOS"     'Corresponde al catalogo de Usuarios
                'Revisa que exista un registro a modificar
                If Trim(Txt_Nombre_Usuario.Text) <> "" Then
                    Fra_Generales_Usuarios.Enabled = True
                    Fra_Usuarios.Enabled = False
                Else
                    MsgBox "Seleccione los datos de las listas", vbInformation
                    Exit Sub
                End If
            Case "CURSOS"       'Corresponde al catalogo de Cursos
                If Trim(Txt_Curso_ID.Text) <> "" Then
                    Fra_Generales_Cursos.Enabled = True
                    Fra_Grid_Cursos.Enabled = False
                Else
                    MsgBox "Seleccione un curso de la lista", vbInformation
                    Exit Sub
                End If
            Case "ZONAS"        'Corresponde al catalogo de zonas
                If Trim(Txt_Nombre_Zona.Text) <> "" Then
                    Fra_Generales_Zonas.Enabled = True
                    Fra_Grid_Zonas.Enabled = False
                Else
                    MsgBox "Seleccione una zona de la lista", vbInformation
                    Exit Sub
                End If
            Case "TRANSPORTES"  'Corresponde al catalogo de Transportes
                If Trim(Txt_Nombre_Transporte.Text) <> "" Then
                    Fra_Generales_Transportes.Enabled = True
                    Fra_Grid_Transportes.Enabled = False
                Else
                    MsgBox "Seleccione un Transporte de la lista", vbInformation
                    Exit Sub
                End If
            Case "GAPS"         'Corresponde al catalogo de Gaps
                If Trim(Txt_Nombre_Gap.Text) <> "" Then
                    Fra_Generales_Gaps.Enabled = True
                    Fra_Grid_Gaps.Enabled = False
                Else
                    MsgBox "Seleccione una Gap de la lista", vbInformation
                    Exit Sub
                End If
            Case "SECCIONES"    'Corresponde al catalogo de secciones
                If Trim(Txt_Seccion_ID.Text) <> "" Then
                    Fra_Generales_Secciones.Enabled = True
                    Fra_Secciones.Enabled = False
                Else
                    MsgBox "Seleccione una sección de la lista", vbExclamation
                    Exit Sub
                End If
            Case "GERENCIAS":
                If Trim(Txt_Gerencia_ID.Text) <> "" Then
                    Fra_Generales_Gerencias.Enabled = True
                    Fra_Grid_Gerencias.Enabled = False
                Else
                    MsgBox "Seleccione una gerencia de la lista", vbInformation
                    Exit Sub
                End If
            
            'Corresponde al catalogo de Tipos de Clientes
            Case "TIPO_CLIENTE":
                If Trim(Txt_Giro_ID.Text) <> "" Then
                    Fra_Generales_Giros.Enabled = True
                    Fra_Grid_Giros.Enabled = False
                Else
                    MsgBox "Seleccione un tipo de cliente de la lista", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Operadores
            Case "OPERADORES":
                If Trim(Txt_Nombre_Operador.Text) <> "" Then
                    Fra_Generales_Operadores.Enabled = True
                    Fra_Grid_Operadores.Enabled = False
                Else
                    MsgBox "Seleccione un Operador de la lista", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Marcas
            Case "MARCAS":
                If Trim(Txt_Nombre_Marca.Text) <> "" Then
                    Fra_Generales_Marcas.Enabled = True
                    Fra_Grid_Marcas.Enabled = False
                Else
                    MsgBox "Selecciones una marca de la lista", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Vendedores
            Case "VENDEDORES":
                If Trim(Txt_Vendedor_ID.Text) <> "" Then
                    Fra_Generales_Vendedores.Enabled = True
                    Fra_Grid_Vendedores.Enabled = False
                Else
                    MsgBox "Seleccione un vendedor de la lista", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de bancos
            Case "BANCOS"
                'Revisa que exista un registro a modificar
                If Trim(Txt_Nombre_Banco.Text) <> "" Then
                    Fra_Generales_Bancos.Enabled = True
                    Fra_Bancos.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                Else
                    MsgBox "Seleccione los datos de las listas", vbInformation
                    Exit Sub
                End If
            Case "UNIDADES"
                If Trim(Txt_Unidad_ID.Text) <> "" Then
                    Fra_Generales_Unidades.Enabled = True
                    Fra_Grid_Unidades.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                Else
                    MsgBox "Seleccione los datos de las listas", vbInformation
                    Exit Sub
                End If
            Case "TIEMPOS_MUERTOS"
               If Trim(Txt_Tiempo_ID.Text) <> "" Then
                    Fra_Generales_Tiempos_Muertos.Enabled = True
                    Fra_Grid_Tiempos_Muertos.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                Else
                    MsgBox "Seleccione los datos de las listas", vbInformation
                    Exit Sub
                End If
            Case "TIPO_NOTAS_CREDITO"
               If Trim(Txt_Tipo_Nota_Credito_ID.Text) <> "" Then
                    Fra_Generales_Tipos_Notas_Credito.Enabled = True
                    Fra_Tipos_Notas_Credito.Enabled = False
                    Btn_Nuevo.Enabled = False
                    Btn_Modificar.Caption = "Actualizar"
                    Btn_Eliminar.Enabled = False
                    Btn_Buscar.Enabled = False
                Else
                    MsgBox "Seleccione los datos de las listas", vbExclamation
                    Exit Sub
                End If
        End Select
        Btn_Nuevo.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Modificar.Caption = "Actualizar"
        Btn_Salir.Caption = "Regresar"
    Else '#  Para modificar el registro
        Select Case Catalogo
            Case "ROLES"        'Corresponde al catalogo de roles
                If Trim(Txt_Nombre_Rol.Text) <> "" Then
                    If Consulta_Nombre_Rol = False Then
                        Modifica_Rol 'Modifica el registro del rol
                    Else
                        MsgBox "El nombre del rol ya esta dado de alta" & Chr(13) & Chr(13) & "favor de introducirlo nuevamente", vbInformation
                        Txt_Nombre_Rol.SetFocus
                        SendKeys "{Home}+{End}"
                        Exit Sub
                    End If
                Else 'Si falta el nombre
                    If Trim(Txt_Nombre_Rol.Text) = "" Then
                        MsgBox "Proporcione el nombre del rol", vbInformation
                        Txt_Nombre_Rol.SetFocus
                        Exit Sub
                    Else 'Si falta seleccionar el tipo
                        If Cmb_Tipo_Rol.ListIndex > -1 Then
                            MsgBox "Seleccione el Tipo de Rol", vbInformation
                            Cmb_Tipo_Rol.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            Case "USUARIOS"
                'Actualiza catálogo de usuarios revisando si los campos obligatorios estan llenos
                If Trim(Txt_Nombre_Usuario.Text) <> "" And Cmb_Estatus.ListIndex > -1 And Cmb_Roles.ListIndex > -1 And Cmb_Area_ID.ListIndex > -1 And Txt_Login.Text <> "" Then
                    If Trim(Txt_Contraseña.Text) <> "" And Trim(Txt_Contraseña_Confirmar.Text) <> "" And (Trim(Txt_Contraseña.Text) = Trim(Txt_Contraseña_Confirmar.Text)) Then
                        'Valida qu la contraseña sea de por lo menos el parámetro de caracteres
                        If Len(Txt_Contraseña.Text) >= Longitud_Minima_Password Then
                            'Valida si el nombre de usuario no ha sido usado para otro registro
                            Mi_SQL = "SELECT Login FROM Cat_Usuarios WHERE Login='" & Trim(Txt_Login.Text) & "'"
                            Mi_SQL = Mi_SQL & " AND Usuario_ID<>'" & Txt_Usuario_ID.Text & "'"
                            Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Cat_Usuarios.EOF Then
                                MsgBox "El LOGIN ya esta siendo usado por otro usuario", vbExclamation
                                Exit Sub
                            End If
                            Modifica_Usuarios
                        Else
                            MsgBox "La longitud del password debe ser por lo menos de " & Longitud_Minima_Password & " caracteres", vbCritical
                            Exit Sub
                        End If
                    Else
                        '#  si no se ah metido la contrasela
                        If Trim(Txt_Contraseña.Text) = "" Then
                            Mensaje ("Introduce la contraseña")
                            Txt_Contraseña.SetFocus
                            Exit Sub
                        End If
                        '#  si no se ah metido la contrasela
                        If Trim(Txt_Contraseña_Confirmar.Text) = "" Then
                            Mensaje ("Confirma la contraseña")
                            Txt_Contraseña_Confirmar.SetFocus
                            Exit Sub
                        End If
                        Mensaje ("No Coinciden las contraseñas")
                        Txt_Contraseña.SetFocus
                        Exit Sub
                    End If
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "CURSOS"       'Catálogo de cursos
                If Trim(Txt_Nombre_Curso.Text) <> "" And Trim(Txt_Horas_Curso.Text) <> "" Then
                    Modifica_Curso
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "ZONAS"        'Corresponde al catalogo de zonas
                If Trim(Txt_Nombre_Zona.Text) <> "" Then
                    Modifica_Zonas
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "TRANSPORTES"  'Corresponde al catalogo de Transportes
                If Trim(Txt_Nombre_Transporte.Text) <> "" <> "" And Cmb_Zona.ListIndex > -1 Then
                    Modifica_Transportes
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "GAPS"         'Corresponde al catalogo de Gaps
                If Trim(Txt_Nombre_Gap.Text) <> "" Then
                    Modifica_Gaps
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "SECCIONES"    'Corresponde al catalogo de secciones
                If Trim(Txt_Seccion_Clave.Text) <> "" And Cmb_Seccion_Supervisor.ListIndex > -1 Then
                    Modifica_Secciones
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            Case "GERENCIAS"
                If Trim(Txt_Nombre_Gerencia.Text) <> "" And Cmb_Supervisor_Gerencia.ListIndex > -1 Then
                    Modifica_Gerencia
                Else
                    MsgBox "Faltan datos para actualizar la gerencia", vbExclamation
                    Exit Sub
                End If
            
            'Corresponde al catalogo de tipos de clientes
            Case "TIPO_CLIENTE"
                If Trim(Txt_Nombre_Giro.Text) <> "" Then
                    Modifica_Giros
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de Operadores
            Case "OPERADORES"
                If Trim(Txt_Nombre_Operador.Text) <> "" Then
                    Modifica_Operadores
                Else
                    MsgBox "Faltan datos para actualizar", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Marcas
            Case "MARCAS"
                If Trim(Txt_Nombre_Marca.Text) <> "" Then
                    Modifica_Marcas
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de vendedores
            Case "VENDEDORES"
                If Trim(Txt_Nombre_Vendedor.Text) <> "" Then
                    Modifica_Vendedores
                Else
                    MsgBox "Faltan datos para actualizar", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de bancos
            Case "BANCOS"
                'Actualiza catálogo de bancos revisando si los campos obligatorios estan llenos
                If Trim(Txt_Nombre_Banco.Text) <> "" And Trim(Txt_Sucursal.Text) <> "" And Trim(Txt_No_Cuenta_Banco.Text) <> "" Then
                    Modifica_Banco
                Else
                    MsgBox "Faltan datos para poder Actualizar el Registro", vbExclamation
                    Exit Sub
                End If
            Case "UNIDADES"
                If Trim(Txt_Nombre_Corto_Unidad.Text) <> "" And Trim(Txt_Nombre_Unidad.Text) <> "" Then
                    Modifica_Unidades
                Else
                    MsgBox "Faltan datos para poder actualizar el registro", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Tiempos muertos
            Case "TIEMPOS_MUERTOS"
                 If Trim(Txt_Descripcion_Tiempos.Text) <> "" Then
                    Modifica_Tiempos_Muertos
                Else
                    MsgBox "Faltan datos para poder actualizar el registro", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Tiempos muertos
            Case "TIPO_NOTAS_CREDITO"
                 If Trim(Txt_Descripcion_Tipos_Notas_Credito.Text) <> "" Then
                    Modifica_Tipos_Notas_Credito
                Else
                    MsgBox "Faltan datos para poder actualizar el registro", vbExclamation
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub Btn_Nuevo_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Usuarios As rdoResultset

    '#  si se desea registrar un nuevo dato
    If Btn_Nuevo.Caption = "Nuevo" Then
        '#  Limpia las cajas de texto
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        'Muestra el picture del catalogo seleccionado
        Select Case Catalogo
            Case "ROLES":       'Corresponde al catalogo de roles
                Txt_Rol_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Roles", "Rol_ID"), "00000")
                Btn_Acceso_Seguridad.Visible = False
                Fra_Acceso_Sistema_Rol.Visible = True
                Fra_Generales_Roles.Enabled = True
                Fra_Roles_Sistema.Visible = False
                Consulta_Configuracion 'Consulta los menus y submenus que se encuentren en el sistema
            Case "USUARIOS":    'Muestra el picture de Usuarios
                'Llama al último registro de la tabla y asigna el siguiente
                Txt_Usuario_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Usuarios", "Usuario_ID"), "00000")
                Fra_Generales_Usuarios.Enabled = True
                Fra_Usuarios.Enabled = False
                Cmb_Estatus.ListIndex = 0
                DTP_Fecha_Caducar_Usuario.Value = DateAdd("d", Dias_Caducidad_Contraseñas, Now)
                Txt_Nombre_Usuario.SetFocus
            Case "GAPS":        'Corresponde al catalogo de Gaps
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Gap_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Gaps", "Gap_ID"), "00000")
                Fra_Generales_Gaps.Enabled = True
                Fra_Grid_Gaps.Enabled = False
                Txt_Nombre_Gap.SetFocus
            Case "CURSOS":      'Corresponde al catálogo de cursos
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Curso_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Cursos", "Curso_ID"), "00000")
                Fra_Generales_Cursos.Enabled = True
                Fra_Grid_Cursos.Enabled = False
                Txt_Nombre_Curso.SetFocus
            Case "ZONAS":       'Corresponde al catalogo de Zonas
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Zona_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Zonas", "Zona_ID"), "00000")
                Fra_Generales_Zonas.Enabled = True
                Fra_Grid_Zonas.Enabled = False
                Txt_Nombre_Zona.SetFocus
            Case "TRANSPORTES": 'Corresponde al catalogo de Transportes
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Transporte_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Transportes", "Transporte_ID"), "00000")
                Call Conectar_Ayudante.Llena_Combo_Item("Zona_ID,Nombre", "Cat_Zonas", Cmb_Zona, 1, "Nombre")
                Fra_Generales_Transportes.Enabled = True
                Fra_Grid_Transportes.Enabled = False
                Txt_Nombre_Transporte.SetFocus
            Case "SECCIONES":   'Corresponde al catalogo de secciones
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Seccion_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Secciones", "Seccion_ID"), "00000")
                Cmb_Seccion_Supervisor.Text = ""
                Fra_Generales_Secciones.Enabled = True
                Fra_Secciones.Enabled = False
                Txt_Seccion_Clave.SetFocus
            Case "GERENCIAS":
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Gerencia_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Gerencias", "Gerencia_ID"), "00000")
                Fra_Generales_Gerencias.Enabled = True
                Fra_Grid_Gerencias.Enabled = False
                Txt_Nombre_Gerencia.SetFocus
                
            'Corresponde al catalogo de Tipos Clientes
            Case "TIPO_CLIENTE":
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Giro_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Giros_Empresariales", "Giro_ID"), "00000")
                Fra_Generales_Giros.Enabled = True
                Fra_Grid_Giros.Enabled = False
                Txt_Nombre_Giro.SetFocus
            'Corresponde al catalogo de Operadores
            Case "OPERADORES":
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Operador_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Operadores", "Operador_ID"), "00000")
                Fra_Generales_Operadores.Enabled = True
                Fra_Grid_Operadores.Enabled = False
                Cmb_Estatus_Operador.ListIndex = 0
                Cmb_Tipo.ListIndex = 0
                Cmb_Tipo.SetFocus
            'Corresponde al catalogo de Marcas
            Case "MARCAS":
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Marca_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Marcas", "Marca_ID"), "00000")
                Fra_Generales_Marcas.Enabled = True
                Fra_Grid_Marcas.Enabled = False
                Txt_Nombre_Corto_Marca.SetFocus
            'Corresponde al catalogo de Vendedores
            Case "VENDEDORES":
                'Llama al ultimo registro de la tabla y asigna el siguiente
                Txt_Vendedor_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Vendedores", "Vendedor_ID"), "00000")
                Fra_Generales_Vendedores.Enabled = True
                Fra_Grid_Vendedores.Enabled = False
                Cmb_Gap_Vendedor.ListIndex = -1
                Cmb_Estatus_Vendedor.ListIndex = 0
                Txt_Clave_Vendedor.SetFocus
            Case "BANCOS"
                'Limpia los texts de la forma
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Banco_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Bancos", "Banco_ID"), "00000")
                Fra_Generales_Bancos.Enabled = True
                Fra_Bancos.Enabled = False
                Cmb_Empresa.ListIndex = 0
                Cmb_Cuenta_Fiscal.ListIndex = 0
                Cmb_Formato.ListIndex = -1
                Txt_Nombre_Banco.SetFocus
            Case "UNIDADES"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Unidad_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Unidades", "Unidad_ID"), "00000")
                Fra_Generales_Unidades.Enabled = True
                Fra_Grid_Unidades.Enabled = False
                Txt_Nombre_Corto_Unidad.SetFocus
            Case "TIEMPOS_MUERTOS"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Tiempo_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tiempos_Muertos", "Tiempo_ID"), "00000")
                Fra_Generales_Tiempos_Muertos.Enabled = True
                Fra_Grid_Tiempos_Muertos.Enabled = False
                Txt_Descripcion_Tiempos.SetFocus
            Case "TIPO_NOTAS_CREDITO"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Txt_Tipo_Nota_Credito_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Notas_Credito", "Tipo_Nota_Credito_ID"), "00000")
                Fra_Generales_Tipos_Notas_Credito.Enabled = True
                Fra_Tipos_Notas_Credito.Enabled = False
                Txt_Descripcion_Tipos_Notas_Credito.SetFocus
        End Select
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Salir.Caption = "Regresar"
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Buscar.Enabled = False
    Else    '#  Para guardar los datos
        Select Case Catalogo
            Case "ROLES"    'Corresponde al catalogo de roles
                If Trim(Txt_Nombre_Rol.Text) <> "" Then
                    If Consulta_Nombre_Rol = False Then
                        Alta_Rol
                    Else
                        MsgBox "Proporcione otro nombre para el rol" & Chr(13) & Chr(13) & _
                               "porque ya se encuentra asignado.", vbInformation
                        Txt_Nombre_Rol.SetFocus
                        Exit Sub
                    End If
                Else 'Si falta el nombre
                    If Trim(Txt_Nombre_Rol.Text) = "" Then
                        MsgBox "Proporcione el nombre del rol", vbInformation
                        Txt_Nombre_Rol.SetFocus
                        Exit Sub
                    Else 'Si falta seleccionar el tipo
                        If Cmb_Tipo_Rol.ListIndex > -1 Then
                            MsgBox "Seleccione el Tipo de Rol", vbInformation
                            Cmb_Tipo_Rol.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            Case "USUARIOS" 'Corresponde al catalogo de usuarios
                If Txt_Nombre_Usuario.Text <> "" And Cmb_Estatus.ListIndex > -1 And Cmb_Roles.ListIndex > -1 And Cmb_Area_ID.ListIndex > -1 And Txt_Login.Text <> "" Then
                    If Trim(Txt_Contraseña.Text) <> "" And Trim(Txt_Contraseña_Confirmar.Text) <> "" And (Trim(Txt_Contraseña.Text) = Trim(Txt_Contraseña_Confirmar.Text)) Then
                        'Valida qu la contraseña sea de por lo menos el parámetro de caracteres
                        If Len(Txt_Contraseña.Text) >= Longitud_Minima_Password Then
                            'Valida si el nombre de usuario no ha sido usado para otro registro
                            Mi_SQL = "SELECT Login FROM Cat_Usuarios WHERE Login='" & Trim(Txt_Login.Text) & "'"
                            Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Cat_Usuarios.EOF Then
                                MsgBox "El LOGIN ya esta siendo usado por otro usuario", vbExclamation
                                Exit Sub
                            End If
                            Rs_Consulta_Cat_Usuarios.Close
                            Alta_Usuarios
                        Else
                            MsgBox "La longitud del password debe ser por lo menos de " & Longitud_Minima_Password & " caracteres", vbCritical
                            Exit Sub
                        End If
                    Else
                        'no se ha capturado la contraseña
                        If Trim(Txt_Contraseña.Text) = "" Then
                            Mensaje ("Introduce la contraseña")
                            Txt_Contraseña.SetFocus
                            Exit Sub
                        End If
                        'falta la confirmación de la contraseña
                        If Trim(Txt_Contraseña_Confirmar.Text) = "" Then
                            Mensaje ("Confirma la contraseña")
                            Txt_Contraseña_Confirmar.SetFocus
                            Exit Sub
                        End If
                        Mensaje ("No Coinciden las contraseñas")
                        Txt_Contraseña.SetFocus
                        Exit Sub
                    End If
                Else
                    MsgBox "Faltan datos para registrar la informacion", vbExclamation
                    Exit Sub
                End If
            Case "GAPS"     'Corresponde al catalogo de Gaps
                If Trim(Txt_Nombre_Gap.Text) <> "" Then
                    Alta_Gaps
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "SECCIONES"    'Corresponde al catalogo de secciones
                If Trim(Txt_Seccion_Clave.Text) <> "" And Cmb_Seccion_Supervisor.ListIndex > -1 Then
                    Alta_Secciones
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "CURSOS"       'Corresponde al catalogo de cursos
                If Trim(Txt_Nombre_Curso.Text) <> "" And Trim(Txt_Horas_Curso.Text) <> "" Then
                    Alta_Curso
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "ZONAS"        'Corresponde al catalogo de zonas
                If Trim(Txt_Nombre_Zona.Text) <> "" Then
                    Alta_Zonas
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "TRANSPORTES"  'Corresponde al catalogo de Transportes
                If Trim(Txt_Nombre_Transporte.Text) <> "" And Cmb_Zona.ListIndex > -1 Then
                    Alta_Transportes
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "GERENCIAS"
                If Trim(Txt_Nombre_Gerencia.Text) <> "" And Cmb_Supervisor_Gerencia.ListIndex > -1 Then
                    Alta_Gerencia
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            
            'Corresponde al catalogo de Giros
            Case "TIPO_CLIENTE"
                If Trim(Txt_Giro_ID.Text) <> "" Then
                    Alta_Giro
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de Operadores
            Case "OPERADORES"
                If Trim(Txt_Nombre_Operador.Text) <> "" Then
                    Alta_Operadores
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            
            'Corresponde al catalogo de Marcas
            Case "MARCAS"
                If Trim(Txt_Nombre_Marca.Text) <> "" Then
                    Alta_Marcas
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de Vendedores
            Case "VENDEDORES"
                If Trim(Txt_Nombre_Vendedor.Text) <> "" Then
                    Alta_Vendedores
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            'Corresponde al catalogo de bancos
            Case "BANCOS"
                'Validacion de datos requeridos
                If Trim(Txt_Nombre_Banco.Text) <> "" And Trim(Txt_Sucursal.Text) <> "" And Trim(Txt_No_Cuenta_Banco.Text) <> "" Then
                    Alta_Banco
                Else
                    MsgBox "Faltan datos para registrar la información", vbExclamation
                    Exit Sub
                End If
            Case "UNIDADES"
                If Trim(Txt_Nombre_Corto_Unidad.Text) <> "" And Trim(Txt_Nombre_Unidad.Text) <> "" Then
                    Alta_Unidades
                Else
                    MsgBox "Faltan datos para registrar la infomación", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Tiempos Muertos
            Case "TIEMPOS_MUERTOS"
                If Trim(Txt_Descripcion_Tiempos.Text) <> "" Then
                    Alta_Tiempos_Muertos
                Else
                    MsgBox "Faltan datos para registrar la infomación", vbInformation
                    Exit Sub
                End If
            'Corresponde al catalogo de Tipos de Notas de Credito
            Case "TIPO_NOTAS_CREDITO"
                If Trim(Txt_Descripcion_Tipos_Notas_Credito.Text) <> "" Then
                    Alta_Tipos_Notas_Credito
                Else
                    MsgBox "Faltan datos para registrar la infomación", vbExclamation
                    Exit Sub
                End If
        End Select
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Modificar.Enabled = True
        Btn_Eliminar.Enabled = True
        Btn_Buscar.Enabled = True
        Btn_Salir.Caption = "Salir"
    End If
End Sub

Private Sub Btn_Salir_Click()
   If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else 'si el letrero es para cancelar una operacion
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Modificar.Enabled = True
        Btn_Buscar.Enabled = True
        Btn_Eliminar.Enabled = True
        Btn_Salir.Caption = "Salir"
        Call Conectar_Ayudante.Limpiar_Textos(Frm_Cat_Generales)
        Select Case Catalogo
            Case "ROLES"
                Btn_Acceso_Seguridad.Visible = True
                Fra_Acceso_Sistema_Rol.Visible = False
                Fra_Generales_Roles.Enabled = False
                Fra_Roles_Sistema.Visible = True
            Case "USUARIOS"
                Fra_Generales_Usuarios.Enabled = False
                Fra_Usuarios.Enabled = True
            Case "CURSOS"
                Fra_Generales_Cursos.Enabled = False
                Fra_Grid_Cursos.Enabled = True
            Case "ZONAS"
                Fra_Generales_Zonas.Enabled = False
                Fra_Grid_Zonas.Enabled = True
            Case "TRANSPORTES"
                Fra_Generales_Transportes.Enabled = False
                Fra_Grid_Transportes.Enabled = True
            Case "GAPS"
                Fra_Generales_Gaps.Enabled = False
                Fra_Grid_Gaps.Enabled = True
            Case "SECCIONES"
                Fra_Generales_Secciones.Enabled = False
                Fra_Secciones.Enabled = True
                Cmb_Seccion_Supervisor.Text = ""
            Case "GERENCIAS":
                Fra_Generales_Gerencias.Enabled = False
                Fra_Grid_Gerencias.Enabled = True
                Cmb_Supervisor_Gerencia.Text = ""
            
            'Corresponde al catalogo de tipos de clientes
            Case "TIPO_CLIENTE":
                Fra_Generales_Giros.Enabled = False
                Fra_Grid_Giros.Enabled = True
            'Corresponde al catalogo de Operadores
            Case "OPERADORES":
                Fra_Generales_Operadores.Enabled = False
                Fra_Grid_Operadores.Enabled = True
            'Corresponde al catalogo de vendedores
            Case "VENDEDORES":
                Fra_Generales_Vendedores.Enabled = False
                Fra_Grid_Vendedores.Enabled = True
                Cmb_Gap_Vendedor.ListIndex = -1
                Cmb_Estatus_Vendedor.ListIndex = -1
            'Corresponde al catalogo de Marcas
            Case "MARCAS":
                Fra_Generales_Marcas.Enabled = False
                Fra_Grid_Marcas.Enabled = True
            'Corresponde al catalogo de bancos
            Case "BANCOS"
                'Limpia los texts de la forma
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Cmb_Empresa.ListIndex = -1
                Fra_Generales_Bancos.Enabled = False
                Fra_Bancos.Enabled = True
                Btn_Salir.SetFocus
            Case "UNIDADES"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Fra_Generales_Unidades.Enabled = False
                Fra_Grid_Unidades.Enabled = True
            'Corresponde al catalogo de Tiempos Muertos
            Case "TIEMPOS_MUERTOS"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Fra_Generales_Tiempos_Muertos.Enabled = False
                Fra_Grid_Tiempos_Muertos.Enabled = True
            Case "TIPO_NOTAS_CREDITO"
                Call Conectar_Ayudante.Limpiar_Textos(Me)
                Fra_Generales_Tipos_Notas_Credito.Enabled = False
                Fra_Tipos_Notas_Credito.Enabled = True
        End Select
    End If
End Sub

Private Sub Btn_Ver_Gap_Click()
    If Trim(Txt_Gap_ID.Text) <> "" Then
        Unload Frm_Cat_Gaps_Layout
        Load Frm_Cat_Gaps_Layout
        Frm_Cat_Gaps_Layout.Lbl_Nombre_Gap.Caption = Trim(Txt_Nombre_Gap.Text)
        Frm_Cat_Gaps_Layout.Lbl_Comentarios_Gap.Caption = Trim(Txt_Comentarios_Gap.Text)
        Call Frm_Cat_Gaps_Layout.Consulta_Empleados_Gap(Trim(Txt_Gap_ID.Text))
    Else
        MsgBox "Selecciona un gap de la lista", vbExclamation
    End If
End Sub


Private Sub Cmb_Seccion_Supervisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Seccion_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Seccion_Supervisor.Text) & "%' OR " & "Apellido_Paterno LIKE '%" & Trim(Cmb_Seccion_Supervisor.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Seccion_Supervisor.Text) & "%')", False, "")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Supervisor_Gerencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(Cmb_Supervisor_Gerencia.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A' AND No_Tarjeta=" & Trim(Cmb_Supervisor_Gerencia.Text) & "", Cmb_Supervisor_Gerencia, 0, "No_Tarjeta")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Supervisor_Gerencia.Text) & "%' OR " & "Apellido_Paterno LIKE '%" & Trim(Cmb_Supervisor_Gerencia.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Supervisor_Gerencia.Text) & "%')", Cmb_Supervisor_Gerencia, 0, "Apellido_Paterno")
        End If
        If Cmb_Supervisor_Gerencia.ListCount > 0 Then Cmb_Supervisor_Gerencia.ListIndex = 0 Else Cmb_Supervisor_Gerencia.Text = ""
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Zona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Zona_ID,Nombre", "Cat_Zonas", Cmb_Zona, 1, "Nombre")
        Cmb_Zona.Text = ""
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Form_Initialize()
    Set Mi_Ayudante = New Ayudante
    Set Mi_Ayudante.Forma = Me
End Sub

Private Sub Form_Load()
    Me.Height = 6750
    Me.Width = 7410
    Call Conectar_Ayudante.Llena_Combo_Item("Rol_ID,Nombre", "Cat_Roles", Cmb_Roles, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Area_ID,Nombre", "Cat_Areas", Cmb_Area_ID, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo("Nombre", "Cfg_Formatos", Cmb_Formato, 1, "Nombre")
    DTP_Fecha_Caducar_Usuario.Value = DateAdd("d", Dias_Caducidad_Contraseñas, Now)
End Sub

Private Sub Form_Resize()
    Mi_Ayudante.Redimensionar_Controles
End Sub

Private Sub Grid_Accesos_Seguridad_Click()
Dim Renglon As Integer          'Indica que renglon se esta consulltando
Dim Tamaño_Renglon As Integer   'Indica el tamaño que deben tener los renglones del grid
    
On Error GoTo HANDLER
    'Oculta los controles del grid
    Chk_Habilitar_Menu_Submenu.Visible = False
    Txt_Habilitar.Visible = False
    'Valida que el renglon seleccionado sea mayor a 1
    If Grid_Accesos_Seguridad.Rows > 1 Then
        'Si la columna seleccionada es la 0 la cual contiene el signo (-/+)
        If Grid_Accesos_Seguridad.Col = 1 Then
            'Valida que se haya seleccionado un renglon que contenga menu
            If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) <> "" Then
                'Si el renglon esta expandido cambia el signo y coloca en 0 el alto de los renglones que dependen de el
                If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0) = "-" Then
                    Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0) = "+" 'Cambia el signo del renglon
                    Tamaño_Renglon = 0 'Tamaño para contraer los renglones
                Else 'Si el renglon esta contraido cambia el signo y coloca el tamaño normal del alto de los renglones que dependen de el
                    Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0) = "-"
                    Tamaño_Renglon = 240 'Tamaño para expander los renglones
                End If
                'Realiza el ciclo para cambiar el tamaño de los renglones, comenzando en el siguiente renglon del que se selecciono
                For Renglon = (Grid_Accesos_Seguridad.RowSel + 1) To Grid_Accesos_Seguridad.Rows - 1
                    'Valida que sea un renglon de submenu para cambiar el tamaño
                    If Trim(Grid_Accesos_Seguridad.TextMatrix(Renglon, 0)) = "" Then
                        Grid_Accesos_Seguridad.RowHeight(Renglon) = Tamaño_Renglon 'Cambia el tamaño del renglon
                    Else 'Si no es renglon de submenu sale del ciclo
                        Exit For
                    End If
                Next
            End If
        Else 'Si la columna seleccionado la 5 la cual contiene el check para habilitar o deshabilitar el menu
            'Valida que si el boton de nuevo o modificar en "Dar de Alta" o "Actualizar" respectivamente para
            'mostrar el control o cambiar de dato en el celda
            If Btn_Salir.Caption = "Regresar" Then
                If Grid_Accesos_Seguridad.Col = 5 Then
                    Chk_Habilitar_Menu_Submenu.BackColor = &HFFFFFF 'Coloca el fondo del check en blanco
                    'Si el valor de la celda es S coloca en TRUE el check
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                        Chk_Habilitar_Menu_Submenu.Value = 1
                    Else 'Si el valor de la celda es N coloca en FALSE el check
                        Chk_Habilitar_Menu_Submenu.Value = 0
                    End If
                    'Mueve el check a la columna y renglon seleccionados
                    Call Conectar_Ayudante.Mover_Control_Grid_CheckBox(Grid_Accesos_Seguridad, Chk_Habilitar_Menu_Submenu)
                Else 'Si la columna seleccionado esta entre la 6 y la 9 las cuales contienen las opciones de alta, cambio, eliminar y consultar
                    If Grid_Accesos_Seguridad.Col > 5 And Grid_Accesos_Seguridad.Col <= 9 Then
                        'Valida que sea un renglon de submenu para realizar este proceso
                        If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) = "" Then
                            'Cambia el color de la celda seleccionada de blanco a amarillo
                            Grid_Accesos_Seguridad.CellBackColor = &HC0FFFF
                        End If
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Grid_Bancos_RowColChange()
    Grid_Bancos_Click
End Sub

Private Sub Grid_Gaps_Click()
Dim Rs_Consulta_Cat_Gaps As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Gaps.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Gaps"
        Mi_SQL = Mi_SQL & " WHERE Gap_ID='" & Trim(Grid_Gaps.TextMatrix(Grid_Gaps.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Gaps = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Gaps.EOF Then
            With Rs_Consulta_Cat_Gaps
                Txt_Gap_ID.Text = .rdoColumns("Gap_ID")
                Txt_Nombre_Gap.Text = .rdoColumns("Nombre")
                Txt_Comentarios_Gap.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Gaps.Close
    End If
End Sub

Private Sub Grid_Gaps_RowColChange()
    Grid_Gaps_Click
End Sub

Private Sub Grid_Secciones_Click()
Dim Rs_Consulta_Cat_Secciones As rdoResultset

    If Grid_Secciones.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Secciones"
        Mi_SQL = Mi_SQL & " WHERE Seccion_ID='" & Trim(Grid_Secciones.TextMatrix(Grid_Secciones.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Secciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Secciones.EOF Then
            With Rs_Consulta_Cat_Secciones
                Txt_Seccion_ID.Text = .rdoColumns("Seccion_ID")
                Txt_Seccion_Clave.Text = .rdoColumns("Clave")
                Cmb_Seccion_Supervisor.Text = ""
                Call Cmb_Seccion_Supervisor_KeyPress(13)
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Seccion_Supervisor)
            End With
        End If
        Rs_Consulta_Cat_Secciones.Close
    End If
End Sub

Private Sub Grid_Secciones_EnterCell()
    Grid_Secciones_Click
End Sub

Private Sub Grid_Transportes_Click()
Dim Rs_Consulta_Cat_Transportes As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Transportes.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Transportes"
        Mi_SQL = Mi_SQL & " WHERE Transporte_ID='" & Grid_Transportes.TextMatrix(Grid_Transportes.RowSel, 0) & "'"
        Set Rs_Consulta_Cat_Transportes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Transportes.EOF Then
            With Rs_Consulta_Cat_Transportes
                Txt_Transporte_ID.Text = .rdoColumns("Transporte_ID")
                Txt_Nombre_Transporte.Text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Zona_ID")) Then
                    Cmb_Zona.Text = .rdoColumns("Zona_ID")
                    Call Conectar_Ayudante.Llena_Combo_Item("Zona_ID,Nombre", "Cat_Zonas", Cmb_Zona, 1, "Zona_ID")
                    Cmb_Zona.ListIndex = 0
                Else
                    Cmb_Zona.ListIndex = -1
                End If
                Txt_Comentarios_Transporte.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Transportes.Close
    End If
End Sub

Private Sub Grid_Transportes_RowColChange()
    Grid_Transportes_Click
End Sub

Private Sub Grid_Giros_Click()
Dim Rs_Consulta_Cat_Giros_Empresariales As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Giros.Rows > 1 Then
        Txt_Giro_ID.Text = Grid_Giros.TextMatrix(Grid_Giros.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Giros_Empresariales"
        Mi_SQL = Mi_SQL & "  WHERE Giro_ID ='" & Txt_Giro_ID.Text & "'"
        Set Rs_Consulta_Cat_Giros_Empresariales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Giros_Empresariales.EOF Then
            With Rs_Consulta_Cat_Giros_Empresariales
                Txt_Giro_ID.Text = .rdoColumns("Giro_ID")
                Txt_Nombre_Giro.Text = .rdoColumns("Nombre")
                Txt_Comentarios_Giros.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Giros_Empresariales.Close
    End If
End Sub

Private Sub Grid_Giros_RowColChange()
    Grid_Giros_Click
End Sub

Private Sub Grid_Marcas_Click()
Dim Rs_Consulta_Cat_Marcas As rdoResultset             'Manejo de registro

    If Grid_Marcas.Rows > 1 Then
        'Consulta la marca
        Mi_SQL = "SELECT * FROM Cat_Marcas"
        Mi_SQL = Mi_SQL & " WHERE Marca_ID='" & Trim(Grid_Marcas.TextMatrix(Grid_Marcas.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Marcas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Marcas.EOF Then
            With Rs_Consulta_Cat_Marcas
                Txt_Marca_ID.Text = .rdoColumns("Marca_ID")
                Txt_Nombre_Marca.Text = .rdoColumns("Nombre")
                Txt_Nombre_Corto_Marca.Text = .rdoColumns("Nombre_Corto")
                Txt_Comentarios_Marcas.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Marcas.Close
    End If
End Sub

Private Sub Grid_Marcas_EnterCell()
    Grid_Marcas_Click
End Sub

Private Sub Grid_Operadores_Click()
Dim Rs_Consulta_Cat_Operadores As rdoResultset             'Manejo de registro

    If Grid_Operadores.Rows > 1 Then
        'Consulta el operador
        Mi_SQL = "SELECT * FROM Cat_Operadores"
        Mi_SQL = Mi_SQL & " WHERE Operador_ID='" & Trim(Grid_Operadores.TextMatrix(Grid_Operadores.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Operadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Operadores.EOF Then
            With Rs_Consulta_Cat_Operadores
                Txt_Operador_ID.Text = .rdoColumns("Operador_ID")
                If Not IsNull(.rdoColumns("Tipo")) Then
                    Cmb_Tipo.Text = .rdoColumns("Tipo")
                Else
                    Cmb_Tipo.ListIndex = 0
                End If
                If .rdoColumns("Estatus") = "A" Then
                    Cmb_Estatus_Operador.ListIndex = 0
                Else
                    Cmb_Estatus_Operador.ListIndex = 1
                End If
                Txt_Nombre_Operador.Text = .rdoColumns("Nombre")
                Txt_Comentarios_Operadores.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Operadores.Close
    End If
End Sub

Private Sub Grid_Operadores_RowColChange()
    Grid_Operadores_Click
End Sub

Private Sub Grid_Roles_RowColChange()
    Grid_Roles_Click
End Sub

Private Sub Grid_Tiempos_Muertos_Click()
    Dim Rs_Consulta_Cat_Tiempos_Muertos As rdoResultset             'Manejo de registro
    'Selecciona el tiempo muerto que esta en la tabla
    If Grid_Tiempos_Muertos.Rows > 1 Then
        Txt_Tiempo_ID.Text = Grid_Tiempos_Muertos.TextMatrix(Grid_Tiempos_Muertos.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Tiempos_Muertos"
        Mi_SQL = Mi_SQL & "  WHERE Tiempo_ID ='" & Txt_Tiempo_ID.Text & "'"
        Set Rs_Consulta_Cat_Tiempos_Muertos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Tiempos_Muertos.EOF Then
            With Rs_Consulta_Cat_Tiempos_Muertos
                Txt_Tiempo_ID.Text = .rdoColumns("Tiempo_ID")
                Txt_Descripcion_Tiempos.Text = .rdoColumns("Descripcion")
                Txt_Comentarios_Tiempos.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Tiempos_Muertos.Close
    End If
End Sub

Private Sub Grid_Cursos_Click()
Dim Rs_Consulta_Cat_Cursos As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Cursos.Rows > 1 Then
        Txt_Curso_ID.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Cursos"
        Mi_SQL = Mi_SQL & "  WHERE Curso_ID ='" & Txt_Curso_ID.Text & "'"
        Set Rs_Consulta_Cat_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Cursos.EOF Then
            With Rs_Consulta_Cat_Cursos
                Txt_Curso_ID.Text = .rdoColumns("Curso_ID")
                Txt_Nombre_Curso.Text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Horas")) Then Txt_Horas_Curso.Text = .rdoColumns("Horas")
                If Not IsNull(.rdoColumns("Tipo")) Then Cmb_Tipo_Curso.Text = .rdoColumns("Tipo")
                If Not IsNull(.rdoColumns("Instructor")) Then Txt_Instructor_Curso.Text = .rdoColumns("Instructor")
                Txt_Comentarios_Curso.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Cursos.Close
    End If
End Sub

Private Sub Grid_Cursos_RowColChange()
    Grid_Cursos_Click
End Sub

Private Sub Grid_Gerencias_Click()
Dim Rs_Consulta_Cat_Gerencias As rdoResultset             'Manejo de registro de la tabla Cat_Gerencias

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Gerencias.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Gerencias"
        Mi_SQL = Mi_SQL & " WHERE Gerencia_ID='" & Trim(Grid_Gerencias.TextMatrix(Grid_Gerencias.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Gerencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Consulta_Cat_Gerencias.EOF Then
            With Rs_Consulta_Cat_Gerencias
                Txt_Gerencia_ID.Text = .rdoColumns("Gerencia_ID")
                Txt_Nombre_Gerencia.Text = .rdoColumns("Nombre")
                Cmb_Supervisor_Gerencia.Text = ""
                Call Cmb_Supervisor_Gerencia_KeyPress(13)
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Supervisor_ID"), Cmb_Supervisor_Gerencia)
            End With
        End If
        Rs_Consulta_Cat_Gerencias.Close
    End If
End Sub

Private Sub Grid_Tipos_Notas_Credito_Click()
Dim Rs_Consulta_Cat_Tipos_Notas_Credito As rdoResultset
    If Grid_Tipos_Notas_Credito.Rows > 1 Then
        Txt_Tipo_Nota_Credito_ID.Text = Grid_Tipos_Notas_Credito.TextMatrix(Grid_Tipos_Notas_Credito.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Tipos_Notas_Credito"
        Mi_SQL = Mi_SQL & "  WHERE Tipo_Nota_Credito_ID='" & Txt_Tipo_Nota_Credito_ID.Text & "'"
        Set Rs_Consulta_Cat_Tipos_Notas_Credito = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Tipos_Notas_Credito.EOF Then
            With Rs_Consulta_Cat_Tipos_Notas_Credito
                Txt_Tipo_Nota_Credito_ID.Text = .rdoColumns("Tipo_Nota_Credito_ID")
                Txt_Descripcion_Tipos_Notas_Credito.Text = .rdoColumns("Descripcion")
                Txt_Comentarios_Tipos_Notas_Credito.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Tipos_Notas_Credito.Close
    End If
End Sub

Private Sub Grid_Unidades_Click()
Dim Rs_Consulta_Cat_Unidades As rdoResultset             'Manejo de registro

    'Selecciona las Unidades que estan en la Tabla
    If Grid_Unidades.Rows > 1 Then
        Txt_Unidad_ID.Text = Grid_Unidades.TextMatrix(Grid_Unidades.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Unidades"
        Mi_SQL = Mi_SQL & "  WHERE Unidad_ID ='" & Txt_Unidad_ID.Text & "'"
        Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Unidades.EOF Then
            With Rs_Consulta_Cat_Unidades
                Txt_Unidad_ID.Text = .rdoColumns("Unidad_ID")
                Txt_Nombre_Unidad.Text = .rdoColumns("Nombre")
                Txt_Nombre_Corto_Unidad.Text = .rdoColumns("Nombre_Corto")
                Txt_Comentarios_Unidad.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Unidades.Close
    End If
End Sub

Private Sub Grid_Usuarios_Click()
Dim Mi_SQL As String                                    'Obtiene los valores de la consulta
Dim Rs_Llenado_Cat_Usuarios As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios
Dim Rs_Consulta_Seguridad_Sistema As rdoResultset       'Manejo de registro de la tabla Seguridad_Sistema
Dim Ctr As Control                                      'Sirve para identificar que tipo de control es el que se esta consultando

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Usuarios.Rows > 1 Then
        Cmb_Roles.ListIndex = -1
        Cmb_Area_ID.ListIndex = -1
        Cmb_Estatus.ListIndex = -1
        Txt_Usuario_ID.Text = Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Usuarios"
        Mi_SQL = Mi_SQL & "  WHERE Usuario_ID ='" & Txt_Usuario_ID.Text & "'"
        Set Rs_Llenado_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Llenado_Cat_Usuarios.EOF Then
            With Rs_Llenado_Cat_Usuarios
                Txt_Nombre_Usuario.Text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Estatus")) Then Cmb_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Estatus"), Cmb_Estatus)
                If Not IsNull(.rdoColumns("Rol_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Rol_ID"), Cmb_Roles)
                If Not IsNull(.rdoColumns("Area_ID")) Then Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Area_ID"), Cmb_Area_ID)
                If Not IsNull(.rdoColumns("No_Nomina")) Then Txt_No_Nomina.Text = .rdoColumns("No_Nomina") Else Txt_No_Nomina.Text = ""
                Txt_Login.Text = .rdoColumns("Login")
                Txt_Contraseña.Text = .rdoColumns("Contraseña")
                Txt_Contraseña_Confirmar.Text = .rdoColumns("Contraseña")
                If Not IsNull(.rdoColumns("Fecha_Caduca")) Then DTP_Fecha_Caducar_Usuario.Value = Format(.rdoColumns("Fecha_Caduca")) Else DTP_Fecha_Caducar_Usuario.Value = Now
                Txt_Comentarios_Usuarios.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Llenado_Cat_Usuarios.Close
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Roles
    'DESCRIPCIÓN: Consulta los roles que se tienen dados de alta
    'PARÁMETROS : Nombre: Indica el nombre del rol que se pretende buscar
    'CREO       : Yazmin A. Delgado Gómez
    'FECHA_CREO : 28-MAYO-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Roles(Nombre As String)
Dim Rs_Consulta_Apl_Roles As rdoResultset 'Consulta los roles que se encuentran dados de alta
    
On Error GoTo HANDLER
    Grid_Roles.Rows = 0
    Grid_Roles.Cols = 3
    'Consulta todos los roles que se encuentran dados de alta
    Mi_SQL = "SELECT * FROM Cat_Roles"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Apl_Roles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Apl_Roles.EOF Then
        With Rs_Consulta_Apl_Roles
            Grid_Roles.AddItem "Rol ID" & Chr(9) & "Nombre del Rol" & Chr(9) & "Comentarios"
            While Not .EOF
                Grid_Roles.AddItem .rdoColumns("Rol_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Comentarios")
                Grid_Roles.FixedRows = 1
                .MoveNext
            Wend
        End With
        'Asigna los tamaños de las columnas del grid_roles
        Grid_Roles.ColWidth(0) = 1500
        Grid_Roles.ColAlignment(0) = flexAlignCenterCenter
        Grid_Roles.ColWidth(1) = 5000
        Grid_Roles.ColAlignment(1) = flexAlignLeftCenter
        Grid_Roles.ColWidth(2) = 0
    End If
    Rs_Consulta_Apl_Roles.Close
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Bancos
    'DESCRIPCIÓN: Consulta los bancos
    'PARÁMETROS : Nombre: Indica el nombre del banco que se pretende buscar
    'CREO       : Ricardo Soria
    'FECHA_CREO : 29-Dic-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Bancos(Nombre As String)
Dim Rs_Consulta_Cat_Bancos As rdoResultset '#  Consulta el banco
    
    On Error GoTo HANDLER
    'Consulta el Banco
    Mi_SQL = "SELECT * FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Bancos.EOF Then
        'Coloca un encabezado en el grid
        Grid_Bancos.Rows = 0
        Grid_Bancos.AddItem "Banco ID" & Chr(9) & "Nombre" & Chr(9) & "No. Cuenta"
        While Not Rs_Consulta_Cat_Bancos.EOF
            Grid_Bancos.AddItem Rs_Consulta_Cat_Bancos!Banco_ID & Chr(9) & Rs_Consulta_Cat_Bancos!Nombre & Chr(9) & Rs_Consulta_Cat_Bancos!No_Cuenta
            Grid_Bancos.FixedRows = 1
            Rs_Consulta_Cat_Bancos.MoveNext
        Wend
        Grid_Bancos.ColWidth(0) = 1200
        Grid_Bancos.ColAlignment(0) = flexAlignCenterCenter
        Grid_Bancos.ColWidth(1) = 3000
        Grid_Bancos.ColAlignment(1) = flexAlignLeftCenter
        Grid_Bancos.ColWidth(2) = 2000
        Grid_Bancos.ColAlignment(2) = flexAlignLeftCenter
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Bancos.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Cursos
'DESCRIPCION: Consulta los cursos registrados en la base de datos
'PARAMETROS : Nombre: Indica el nombre del curso a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Cursos(Nombre As String)
Dim Rs_Consulta_Cat_Cursos As rdoResultset         'Manejo del Registro
    
On Error GoTo HANDLER
    Grid_Cursos.Rows = 0
    Grid_Cursos.AddItem "Curso ID" & Chr(9) & "Nombre" & Chr(9) & "Horas"
    'Consulta el curso
    Mi_SQL = "SELECT Curso_ID,Nombre,Horas"
    Mi_SQL = Mi_SQL & " FROM Cat_Cursos"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Cat_Cursos.EOF
        Grid_Cursos.AddItem Rs_Consulta_Cat_Cursos.rdoColumns("Curso_ID") & Chr(9) & Rs_Consulta_Cat_Cursos.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Cursos.rdoColumns("Horas")
        Grid_Cursos.FixedRows = 1
        Rs_Consulta_Cat_Cursos.MoveNext
    Wend
    Grid_Cursos.ColWidth(0) = 1000
    Grid_Cursos.ColWidth(1) = 4500
    Grid_Cursos.ColWidth(2) = 1000
    Rs_Consulta_Cat_Cursos.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Unidades
    'DESCRIPCIÓN:           Consulta las Unidades
    'PARÁMETROS :           Nombre: Indica el nombre de la unidad a buscar
    'CREO       :           Rafael Muñoz
    'FECHA_CREO :           04-Sep-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Unidades(Nombre As String)
Dim Rs_Consulta_Cat_Unidades As rdoResultset         'Manejo del Registro
    
On Error GoTo HANDLER
    'Consulta la unidad
    Mi_SQL = "SELECT Unidad_ID,Nombre,Nombre_Corto"
    Mi_SQL = Mi_SQL & " FROM Cat_Unidades"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Unidades = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Unidades.EOF Then
        'Coloca un encabezado en el grid
        Grid_Unidades.Rows = 0
        Grid_Unidades.AddItem "Unidad ID" & Chr(9) & "Nombre" & Chr(9) & "Nombre Corto"
        While Not Rs_Consulta_Cat_Unidades.EOF
            Grid_Unidades.AddItem Rs_Consulta_Cat_Unidades.rdoColumns("Unidad_ID") & Chr(9) & Rs_Consulta_Cat_Unidades.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Unidades.rdoColumns("Nombre_Corto")
            Grid_Unidades.FixedRows = 1
            Rs_Consulta_Cat_Unidades.MoveNext
        Wend
        Grid_Unidades.ColWidth(0) = 1500
        Grid_Unidades.ColWidth(1) = 2500
        Grid_Unidades.ColWidth(2) = 2500
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Unidades.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Transportes
'DESCRIPCION: Consulta los Transportes
'PARAMETROS : Nombre: Indica el nombre del Transporte a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Transportes(Nombre As String)
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Transportes As rdoResultset         'Manejo del Registro
Dim Rs_Consulta_Cat_Zonas As rdoResultset

On Error GoTo HANDLER
    'Consulta el Transporte
    Mi_SQL = "SELECT * FROM Cat_Transportes"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Transportes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Transportes.EOF Then
        'Coloca un encabezado en el grid
        Grid_Transportes.Rows = 0
        Grid_Transportes.AddItem "Transporte ID" & Chr(9) & "Nombre" & Chr(9) & "Zona"
        While Not Rs_Consulta_Cat_Transportes.EOF
            If Not IsNull(Rs_Consulta_Cat_Transportes.rdoColumns("Zona_ID")) Then
                Mi_SQL = "SELECT Zona_ID,Nombre FROM Cat_Zonas"
                Mi_SQL = Mi_SQL & " WHERE Zona_ID='" & Rs_Consulta_Cat_Transportes.rdoColumns("Zona_ID") & "'"
                Set Rs_Consulta_Cat_Zonas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Cat_Zonas.EOF Then
                    Grid_Transportes.AddItem Rs_Consulta_Cat_Transportes.rdoColumns("Transporte_ID") & Chr(9) & Rs_Consulta_Cat_Transportes.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Zonas.rdoColumns("Nombre")
                Else
                    Grid_Transportes.AddItem Rs_Consulta_Cat_Transportes.rdoColumns("Transporte_ID") & Chr(9) & Rs_Consulta_Cat_Transportes.rdoColumns("Nombre")
                End If
                Rs_Consulta_Cat_Zonas.Close
            Else
                Grid_Transportes.AddItem Rs_Consulta_Cat_Transportes.rdoColumns("Transporte_ID") & Chr(9) & Rs_Consulta_Cat_Transportes.rdoColumns("Nombre")
            End If
            Grid_Transportes.FixedRows = 1
            Rs_Consulta_Cat_Transportes.MoveNext
        Wend
        Grid_Transportes.ColWidth(0) = 1000
        Grid_Transportes.ColWidth(1) = 3000
        Grid_Transportes.ColWidth(2) = 2600
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Transportes.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Gerencias
'DESCRIPCION: Consulta las Gerencias
'PARAMETROS : Nombre: Indica el nombre de la gerencia a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Marzo-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Gerencia(Nombre As String)
Dim Rs_Consulta_Cat_Gerencia As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Coloca un encabezado en el grid
    Grid_Gerencias.Rows = 0
    Grid_Gerencias.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Supervisor"
    'Consulta las gerencias
    Mi_SQL = "SELECT Cat_Gerencias.*,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Supervisor"
    Mi_SQL = Mi_SQL & " FROM Cat_Gerencias,Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Cat_Gerencias.Supervisor_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Gerencias.Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Gerencias.Nombre"
    Set Rs_Consulta_Cat_Gerencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    While Not Rs_Consulta_Cat_Gerencias.EOF
        Grid_Gerencias.AddItem Rs_Consulta_Cat_Gerencias.rdoColumns("Gerencia_ID") _
            & Chr(9) & Rs_Consulta_Cat_Gerencias.rdoColumns("Nombre") _
            & Chr(9) & Rs_Consulta_Cat_Gerencias.rdoColumns("Supervisor")
        Grid_Gerencias.FixedRows = 1
        Rs_Consulta_Cat_Gerencias.MoveNext
    Wend
    Rs_Consulta_Cat_Gerencias.Close
    Grid_Gerencias.ColWidth(0) = 1000
    Grid_Gerencias.ColAlignment(0) = flexAlignCenterCenter
    Grid_Gerencias.ColWidth(1) = 1500
    Grid_Gerencias.ColAlignment(1) = flexAlignLeftCenter
    Grid_Gerencias.ColWidth(2) = 4000
    Grid_Gerencias.ColAlignment(2) = flexAlignLeftCenter
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Consulta_Marcas
    'DESCRIPCION: Consulta las Marcas
    'PARAMETROS : Nombre: Indica el nombre de la Marca a buscar
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 22-Agosto-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'******************************************************************************
Public Sub Consulta_Marcas(Nombre As String)
Dim Rs_Consulta_Cat_Marcas As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Consulta la Marca
    Mi_SQL = "SELECT * FROM Cat_Marcas"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Marcas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Marcas.EOF Then
        'Coloca un encabezado en el grid
        Grid_Marcas.Rows = 0
        Grid_Marcas.AddItem "Marca ID" & Chr(9) & "Nombre" & Chr(9) & "Nombre Corto" & Chr(9) & "Comentarios"
        While Not Rs_Consulta_Cat_Marcas.EOF
            Grid_Marcas.AddItem Rs_Consulta_Cat_Marcas.rdoColumns("Marca_ID") & Chr(9) & Rs_Consulta_Cat_Marcas.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Marcas.rdoColumns("Nombre_Corto") & Chr(9) & Rs_Consulta_Cat_Marcas.rdoColumns("Comentarios")
            Grid_Marcas.FixedRows = 1
            Rs_Consulta_Cat_Marcas.MoveNext
        Wend
        Grid_Marcas.ColWidth(0) = 1000
        Grid_Marcas.ColWidth(1) = 3000
        Grid_Marcas.ColWidth(2) = 2000
        Grid_Marcas.ColAlignment(2) = flexAlignLeftCenter
        Grid_Marcas.ColWidth(3) = 0
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Marcas.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Secciones
'DESCRIPCION: Consulta las Secciones y llena el grid
'PARAMETROS : Nombre: Indica el nombre del Seccion a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 28-Septiembre-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Secciones(Nombre As String)
Dim Rs_Consulta_Cat_Secciones As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Coloca un encabezado en el grid
    Grid_Secciones.Rows = 0
    Grid_Secciones.AddItem "Seccion ID" & Chr(9) & "Clave" & Chr(9) & "Supervisor"
    'Consulta la Seccion
    Mi_SQL = "SELECT Cat_Secciones.*,(Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) AS Supervisor"
    Mi_SQL = Mi_SQL & " FROM Cat_Secciones,Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Cat_Secciones.Supervisor_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Secciones.Clave LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Secciones.Clave"
    Set Rs_Consulta_Cat_Secciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    While Not Rs_Consulta_Cat_Secciones.EOF
        Grid_Secciones.AddItem Rs_Consulta_Cat_Secciones.rdoColumns("Seccion_ID") & Chr(9) & Rs_Consulta_Cat_Secciones.rdoColumns("Clave") & Chr(9) & Rs_Consulta_Cat_Secciones.rdoColumns("Supervisor")
        Grid_Secciones.FixedRows = 1
        Rs_Consulta_Cat_Secciones.MoveNext
    Wend
    Rs_Consulta_Cat_Secciones.Close
    Grid_Secciones.ColWidth(0) = 1000
    Grid_Secciones.ColAlignment(0) = flexAlignCenterCenter
    Grid_Secciones.ColWidth(1) = 1000
    Grid_Secciones.ColAlignment(1) = flexAlignCenterCenter
    Grid_Secciones.ColWidth(2) = 4000
    Grid_Secciones.ColAlignment(2) = flexAlignLeftCenter
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Operadores
    'DESCRIPCIÓN: Consulta los Operadores
    'PARÁMETROS : Nombre: Indica el nombre del Operador a buscar
    'CREO       : Rafael Muñoz
    'FECHA_CREO : 07-Feb-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Operadores(Nombre As String)
Dim Rs_Consulta_Cat_Operadores As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Consulta el Operador
    Mi_SQL = "SELECT * FROM Cat_Operadores"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Operadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Operadores.EOF Then
        'Coloca un encabezado en el grid
        Grid_Operadores.Rows = 0
        Grid_Operadores.AddItem "Operador ID" & Chr(9) & "Nombre" & Chr(9) & "Tipo"
        While Not Rs_Consulta_Cat_Operadores.EOF
            Grid_Operadores.AddItem Rs_Consulta_Cat_Operadores.rdoColumns("Operador_ID") & Chr(9) & Trim(Rs_Consulta_Cat_Operadores.rdoColumns("Nombre")) & Chr(9) & Rs_Consulta_Cat_Operadores.rdoColumns("Tipo")
            Grid_Operadores.FixedRows = 1
            Rs_Consulta_Cat_Operadores.MoveNext
        Wend
        Grid_Operadores.ColWidth(0) = 1000
        Grid_Operadores.ColWidth(1) = 4000
        Grid_Operadores.ColWidth(2) = 1500
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Operadores.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Zonas
'DESCRIPCION: Consulta las Zonas
'PARAMETROS : Nombre: Indica el nombre de la zona a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Zonas(Nombre As String)
Dim Rs_Consulta_Cat_Zonas As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Consulta la Zona
    Mi_SQL = "SELECT * FROM Cat_Zonas"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Zonas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Zonas.EOF Then
        'Coloca un encabezado en el grid
        Grid_Zonas.Rows = 0
        Grid_Zonas.AddItem "Zona ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
        While Not Rs_Consulta_Cat_Zonas.EOF
            Grid_Zonas.AddItem Rs_Consulta_Cat_Zonas!Zona_ID & Chr(9) & Rs_Consulta_Cat_Zonas!Nombre & Chr(9) & Rs_Consulta_Cat_Zonas!Comentarios
            Grid_Zonas.FixedRows = 1
            Rs_Consulta_Cat_Zonas.MoveNext
        Wend
        Grid_Zonas.ColWidth(0) = 1500
        Grid_Zonas.ColWidth(1) = 2500
        Grid_Zonas.ColWidth(2) = 2000
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Zonas.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Transportes
    'DESCRIPCIÓN: Consulta los Transportes
    'PARÁMETROS : Nombre: Indica el nombre del Transporte a buscar
    'CREO       : Rafael Muñoz
    'FECHA_CREO : 08-Ene-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Giros(Nombre As String)
Dim Rs_Consulta_Cat_Giros_Empresariales As rdoResultset         'Manejo del Registro

On Error GoTo HANDLER
    'Consulta el giro
    Mi_SQL = "SELECT * FROM Cat_Giros_Empresariales"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Giros_Empresariales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Giros_Empresariales.EOF Then
        'Coloca un encabezado en el grid
        Grid_Giros.Rows = 0
        Grid_Giros.AddItem "Tipo Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
        While Not Rs_Consulta_Cat_Giros_Empresariales.EOF
            Grid_Giros.AddItem Rs_Consulta_Cat_Giros_Empresariales!Giro_ID & Chr(9) & Rs_Consulta_Cat_Giros_Empresariales!Nombre & Chr(9) & Rs_Consulta_Cat_Giros_Empresariales!Comentarios
            Grid_Giros.FixedRows = 1
            Rs_Consulta_Cat_Giros_Empresariales.MoveNext
        Wend
        Grid_Giros.ColWidth(0) = 1200
        Grid_Giros.ColWidth(1) = 3000
        Grid_Giros.ColWidth(2) = 0
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Giros_Empresariales.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Gaps
'DESCRIPCION: Consulta las Gaps de la base de datos y los pone en el grid
'PARAMETROS : Nombre: Indica el nombre de la Gap a buscar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Gaps(Nombre As String)
Dim Rs_Consulta_Cat_Gaps As rdoResultset         'Manejo del Registro
    
On Error GoTo HANDLER
    'Consulta la Gaps
    Mi_SQL = "SELECT * FROM Cat_Gaps"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Gaps = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    Grid_Gaps.Rows = 0
    Grid_Gaps.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
    While Not Rs_Consulta_Cat_Gaps.EOF
        Grid_Gaps.AddItem Rs_Consulta_Cat_Gaps.rdoColumns("Gap_ID") & Chr(9) & Rs_Consulta_Cat_Gaps.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Gaps.rdoColumns("Comentarios")
        Grid_Gaps.FixedRows = 1
        Rs_Consulta_Cat_Gaps.MoveNext
    Wend
    Rs_Consulta_Cat_Gaps.Close
    Grid_Gaps.ColWidth(0) = 1000
    Grid_Gaps.ColAlignment(0) = flexAlignCenterCenter
    Grid_Gaps.ColWidth(1) = 3000
    Grid_Gaps.ColAlignment(1) = flexAlignLeftCenter
    Grid_Gaps.ColWidth(2) = 2500
    Grid_Gaps.ColAlignment(2) = flexAlignLeftCenter
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Vendedores
    'DESCRIPCIÓN: Consulta los Vendedores
    'PARÁMETROS : Nombre: Indica el nombre del vendedor a buscar
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Marzo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Vendedores(Nombre As String)
Dim Rs_Consulta_Cat_Vendedores As rdoResultset         'Manejo del Registro
    
On Error GoTo HANDLER
    'Consulta el Vendedor
    Mi_SQL = "SELECT Vendedor_ID,Nombre,Clave FROM Cat_Vendedores"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Cat_Vendedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Vendedores.EOF Then
        'Coloca un encabezado en el grid
        Grid_Vendedores.Rows = 0
        Grid_Vendedores.AddItem "Vendedor ID" & Chr(9) & "Nombre" & Chr(9) & "Clave"
        While Not Rs_Consulta_Cat_Vendedores.EOF
            Grid_Vendedores.AddItem Rs_Consulta_Cat_Vendedores.rdoColumns("Vendedor_ID") & Chr(9) & Rs_Consulta_Cat_Vendedores.rdoColumns("Nombre") & Chr(9) & Rs_Consulta_Cat_Vendedores.rdoColumns("Clave")
            Grid_Vendedores.FixedRows = 1
            Rs_Consulta_Cat_Vendedores.MoveNext
        Wend
        Grid_Vendedores.FixedCols = 1
        Grid_Vendedores.ColWidth(0) = 1000
        Grid_Vendedores.ColWidth(1) = 4000
        Grid_Vendedores.ColAlignment(1) = flexAlignLeftCenter
        Grid_Vendedores.ColWidth(2) = 1600
        Grid_Vendedores.ColAlignment(2) = flexAlignLeftCenter
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Vendedores.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Configuracion
    'DESCRIPCIÓN: Consulta todos los Menus del sistema en el MDI colocandolos en el grid
    'PARÁMETROS :
    'CREO       : Jorge Razo
    'FECHA_CREO :
    'MODIFICO          : Yazmin Delgado Gómez
    'FECHA_MODIFICO    : 28-Mayo-2007
    'CAUSA_MODIFICACIÓN: Porque se modifico la forma de accesar al sistema
'*******************************************************************************
Private Sub Consulta_Configuracion()
'On Error GoTo HANDLER
Dim Ctl As Control               'Indica que control es el que se esta consultando en el sistema
Dim Contador_Columnas As Integer 'Indica que columna del grid se esta consultando

    'Limpia el grid de los menus
    Grid_Accesos_Seguridad.Rows = 0
    'Grid_Accesos_Seguridad.Font.Name = "Arial"
    'Grid_Accesos_Seguridad.Font.Size = 8
    'Pone en el encabezado del grid los nombres de la columnas
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "Menu" & Chr(9) & "Submenu" & _
        Chr(9) & "Nombre" & Chr(9) & "Tipo" & Chr(9) & "Habilitar" & _
        Chr(9) & "Alta" & Chr(9) & "Cambio" & Chr(9) & "Elimina" & Chr(9) & "Consulta"
    'Agrega los menus y submenus que se encuentran en el MDI del sistema
    For Each Ctl In MDIFrm_Apl_Principal.Controls
        'Si es un menu o submenu lo coloca en el grid
        If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Or UCase(Mid(Ctl.Name, 1, 7)) = UCase("SUBMENU") Then
            'Si es un menu o encabezado
            If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                'Coloca los datos del menu en el grid
                Grid_Accesos_Seguridad.AddItem "-" & Chr(9) & UCase(Ctl.Caption) _
                & Chr(9) & "" & Chr(9) & Ctl.Name & Chr(9) & "Encabezado" _
                & Chr(9) & "S" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                'Se posiciona en el ultimo renglon para pintarlo de gris
                Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                'Agrega el color gris a la fila que tiene el encabezado
                For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                    Grid_Accesos_Seguridad.Col = Contador_Columnas
                    Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                Next Contador_Columnas
                'Vuelve a colocar el signo - en el renglon agregado
                Grid_Accesos_Seguridad.TextMatrix((Grid_Accesos_Seguridad.Rows - 1), 0) = "-"
            Else 'Si es un Submenu
                'Coloca los datos del submenu en el grid
                Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "" & Chr(9) & UCase(Ctl.Caption) & _
                Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & Chr(9) & "S" & _
                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "S"
            End If
        End If
    Next Ctl
    'Configura el tamaño de las columnas del grid_accesos_seguridad
    Grid_Accesos_Seguridad.FixedRows = 1
    Grid_Accesos_Seguridad.ColWidth(0) = 200  '-
    Grid_Accesos_Seguridad.ColWidth(1) = 2150 'Menu
    Grid_Accesos_Seguridad.ColWidth(2) = 2700 'SubMenu
    Grid_Accesos_Seguridad.ColWidth(3) = 0    'Nombre Menu/Submenu
    Grid_Accesos_Seguridad.ColWidth(4) = 0    'Tipo
    Grid_Accesos_Seguridad.ColWidth(5) = 900  'Habilitar
    Grid_Accesos_Seguridad.ColAlignment(5) = 3
    Grid_Accesos_Seguridad.ColWidth(6) = 700  'Alta
    Grid_Accesos_Seguridad.ColAlignment(6) = 3
    Grid_Accesos_Seguridad.ColWidth(7) = 700  'Cambio
    Grid_Accesos_Seguridad.ColAlignment(7) = 3
    Grid_Accesos_Seguridad.ColWidth(8) = 700  'Eliminar
    Grid_Accesos_Seguridad.ColAlignment(8) = 3
    Grid_Accesos_Seguridad.ColWidth(9) = 700  'Consultar
    Grid_Accesos_Seguridad.ColAlignment(9) = 3
    'Coloca como contraidos los renglones del grid de menus
    Call Collapse_Grid
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Modifica_Rol
    'DESCRIPCIÓN: Modifica los datos del registro del rol que selecciono el
    '             usuario así como elimina y da de alta los menus y submenus
    '             que tiene dados de alta el sistema
    'PARÁMETROS :
    'CREO       : Yazmin A Delgado Gómez
    'FECHA_CREO : 28-Mayo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Rol()
    Dim Rs_Modifica_Apl_Cat_Roles As rdoResultset 'Modifica el registro del rol que fue seleccionado por el usuario
    Dim Rs_Alta_Apl_Cat_Accesos As rdoResultset   'Da de alta los accesos que va a contener el rol
    Dim Menus As Integer                          'Contador que sirve para ver en que posición me encuentro en el grid
    Dim Ctl As Control                            'Indica que tipo de control es el que se esta consultando de la pantalla principal

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
        'Consulta los datos que tiene asignado el rol que fue seleccionado por el usuario
        Mi_SQL = "SELECT * FROM Cat_Roles"
        Mi_SQL = Mi_SQL & " WHERE Rol_ID = '" & Trim(Txt_Rol_ID.Text) & "'"
        Set Rs_Modifica_Apl_Cat_Roles = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Modifica_Apl_Cat_Roles.EOF Then
            With Rs_Modifica_Apl_Cat_Roles
                .Edit
                    .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Rol.Text))
                    .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Rol.Text))
                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                    .rdoColumns("Fecha_Modifico") = Now
                .Update
            End With
        End If
        Rs_Modifica_Apl_Cat_Roles.Close
        'Si se elimino correctamente los menus y submenus que se tenian asignados entonces
        'da de alta nuevamente estos mismos
        If Conectar_Ayudante.Elimina_Catalogo("Seguridad_Sistema", "Rol_ID", Trim(Txt_Rol_ID)) = True Then
            'Da de alta los menus y submenus al cual va a tener acceso el rol
            Set Rs_Alta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Agregar("Seguridad_Sistema")
            'Llena el Grid con los datos actualizados
            For Menus = 1 To Grid_Accesos_Seguridad.Rows - 1
                With Rs_Alta_Apl_Cat_Accesos
                    .AddNew
                        .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.Text)
                        If Grid_Accesos_Seguridad.TextMatrix(Menus, 1) <> "" Then
                            .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 1))
                        End If
                        If Grid_Accesos_Seguridad.TextMatrix(Menus, 2) <> "" Then
                            .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 2))
                        End If
                        .rdoColumns("Nombre_Sistema") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 3))
                        .rdoColumns("Tipo") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 4))
                        .rdoColumns("Habilitar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 5))
                        .rdoColumns("Alta") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 6))
                        .rdoColumns("Cambio") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 7))
                        .rdoColumns("Eliminar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 8))
                        .rdoColumns("Consultar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 9))
                    .Update
                End With
            Next Menus
            Rs_Alta_Apl_Cat_Accesos.Close
        End If
    Conexion_Base.CommitTrans
    
    'Coloca los nuevos datos en el renglon del grid seleccionado
    Grid_Roles.TextMatrix(Grid_Roles.RowSel, 1) = Trim(UCase(Txt_Nombre_Rol.Text))
    Grid_Roles.TextMatrix(Grid_Roles.RowSel, 2) = Trim(UCase(Txt_Comentarios_Rol.Text))
    Fra_Roles_Sistema.Visible = True
    Fra_Acceso_Sistema_Rol.Visible = False
    Fra_Generales_Roles.Enabled = False
    Btn_Buscar.Enabled = True
    Btn_Acceso_Seguridad.Visible = True
    Btn_Acceso_Seguridad.Caption = "Control de Acceso"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Roles_Usuarios", Me)
    MsgBox "El rol " & Trim(UCase(Txt_Nombre_Rol.Text)) & Chr(13) & Chr(13) & _
           "ha sido modificado", vbInformation
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Eliminar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Rol
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del rol que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Yazmin A Delgado Gómez
    'FECHA_CREO : 28-Mayo-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Rol()
Dim Rs_Alta_Apl_Cat_Roles As rdoResultset    'Da de alta el  nuevo rol en la base de datos
Dim Rs_Alta_Apl_Cat_Accesos As rdoResultset  'Manejo de registro de Apl_Cat_Accesos, guarda a que menus son lo que va a tener acceso el rol en el sistema
Dim Menus As Integer                         'Contador que sirve para ver en que posición me encuentro en el grid
Dim Ctl As Control                           'Indica que tipo de control es el que se esta consultando de la pantalla principal

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Da de alta el rol en la base de datos
    Set Rs_Alta_Apl_Cat_Roles = Conectar_Ayudante.Recordset_Agregar("Cat_Roles")
    With Rs_Alta_Apl_Cat_Roles
        .AddNew
            .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.Text)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Rol.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Rol.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Apl_Cat_Roles.Close

    'Llena el Grid con los datos actualizados
    For Menus = 1 To Grid_Accesos_Seguridad.Rows - 1
        'Da de alta los menus y submenus al cual va a tener acceso el rol
        Set Rs_Alta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Agregar("Seguridad_Sistema")
        With Rs_Alta_Apl_Cat_Accesos
            .AddNew
                .rdoColumns("Rol_ID") = Trim(Txt_Rol_ID.Text)
                If Grid_Accesos_Seguridad.TextMatrix(Menus, 1) <> "" Then
                    .rdoColumns("Menu_Habilitado") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 2))
                End If
                .rdoColumns("Nombre_Sistema") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 3))
                .rdoColumns("Tipo") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 4))
                .rdoColumns("Habilitar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 5))
                .rdoColumns("Alta") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 6))
                .rdoColumns("Cambio") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 7))
                .rdoColumns("Eliminar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 8))
                .rdoColumns("Consultar") = Trim(Grid_Accesos_Seguridad.TextMatrix(Menus, 9))
            .Update
        End With
        Rs_Alta_Apl_Cat_Accesos.Close
    Next Menus
    Conexion_Base.CommitTrans
    If Grid_Roles.Rows = 0 Then
        Grid_Roles.AddItem "Rol ID" & Chr(9) & "Nombre del Rol" & Chr(9) & "Comentarios"
    End If
    Grid_Roles.AddItem Txt_Rol_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Rol.Text)) & Chr(9) & Trim(UCase(Txt_Comentarios_Rol.Text))
    'Asigna los tamaños de las columnas del grid_roles
    Grid_Roles.FixedRows = 1
    Grid_Roles.ColWidth(0) = 1550
    Grid_Roles.ColWidth(1) = 6550
    Grid_Roles.ColAlignment(1) = 1
    Grid_Roles.ColWidth(2) = 0
    Fra_Roles_Sistema.Visible = True
    Fra_Acceso_Sistema_Rol.Visible = False
    Fra_Generales_Roles.Enabled = False
    Btn_Acceso_Seguridad.Visible = True
    Btn_Acceso_Seguridad.Caption = "Control de Acceso"
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Roles_Usuarios", Me)
    MsgBox "El rol " & Trim(UCase(Txt_Nombre_Rol.Text)) & Chr(13) & Chr(13) & _
           "ha sido dado de alta", vbInformation
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Buscar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Btn_Buscar.Enabled = True
    Btn_Salir.Caption = "Salir"
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Private Sub Chk_Habilitar_Menu_Submenu_Click()
On Error GoTo HANDLER
Dim Renglon As Integer       'Indica que renglon es el que se esta consultando

'Si el check esta visible realiza el proceso que implica cambiar de true a false y viceverza
If Chk_Habilitar_Menu_Submenu.Visible = True Then
    'Cambia el dato de la celda segun el cambio en el check
    'Si el check cambia a FALSE
    If Chk_Habilitar_Menu_Submenu.Value = 0 Then
        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "N"
    Else 'Si el check cambia a TRUE
        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S"
    End If
    
    'Valida si el renglon seleccionado es un menu para cambiar los datos de los renglones que dependen de este
    If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) <> "" Then
        'Realiza el ciclo de los renglones que pertenecen al menu comenzando del renglon siguiente al seleccionado
        For Renglon = (Grid_Accesos_Seguridad.RowSel + 1) To Grid_Accesos_Seguridad.Rows - 1
            'Valida que sea un renglon de submenu
            If Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "" Then
                'Si el check esta en FALSE coloca en N todas las
                If Chk_Habilitar_Menu_Submenu.Value = 0 Then
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 5) = "N" 'Habilitar
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 6) = "N" 'Alta
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 7) = "N" 'Cambio
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 8) = "N" 'Eliminar
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 9) = "N" 'Consultar
                Else 'Si el check esta en TRUE coloca los permisos minimos
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 5) = "S" 'Habilitar
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 6) = "N" 'Alta
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 7) = "N" 'Cambio
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 8) = "N" 'Eliminar
                    Grid_Accesos_Seguridad.TextMatrix(Renglon, 9) = "S" 'Consultar
                End If
            Else 'Si no es un submenu sale del ciclo for
                Exit For
            End If
        Next
    Else 'Si es un submenu cambia los valores del renglon seleccionado
        'Si el check esta en FALSE coloca en N todas las
        If Chk_Habilitar_Menu_Submenu.Value = 0 Then
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "N" 'Habilitar
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "N" 'Alta
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "N" 'Cambio
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "N" 'Eliminar
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "N" 'Consultar
        Else 'Si el check esta en TRUE coloca los permisos minimos
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" 'Habilitar
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 6) = "N" 'Alta
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 7) = "N" 'Cambio
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 8) = "N" 'Eliminar
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 9) = "S" 'Consultar
        End If
    End If
End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Grid_Usuarios_EnterCell()
    Grid_Usuarios_Click
End Sub

Private Sub Grid_Vendedores_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Vendedores As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona los usuarios que estan en la Tabla
    If Grid_Vendedores.Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Vendedores"
        Mi_SQL = Mi_SQL & " WHERE Vendedor_ID='" & Trim(Grid_Vendedores.TextMatrix(Grid_Vendedores.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Vendedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Vendedores.EOF Then
            With Rs_Consulta_Cat_Vendedores
                Txt_Vendedor_ID.Text = .rdoColumns("Vendedor_ID")
                'Selecciona el estatus del combo dependiendo el valor guardado en la bd
                If Not IsNull(.rdoColumns("Estatus")) Then
                    If .rdoColumns("Estatus") = "A" Then
                        Cmb_Estatus_Vendedor.ListIndex = 0
                    Else
                        Cmb_Estatus_Vendedor.ListIndex = 1
                    End If
                End If
                Txt_Nombre_Vendedor.Text = .rdoColumns("Nombre")
                'Llena el combo con la Gap dependiendo el Id de la Gap
                If Not IsNull(.rdoColumns("Gap_ID")) Then
                    Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Ciudad_ID"), Cmb_Ciudad_Vendedor)
                Else
                    Cmb_Gap_Vendedor.ListIndex = -1
                End If
                If Not IsNull(.rdoColumns("Clave")) Then
                    Txt_Clave_Vendedor.Text = .rdoColumns("Clave")
                Else
                    Txt_Clave_Vendedor.Text = ""
                End If
                If Not IsNull(.rdoColumns("RFC")) Then
                    Txt_RFC_Vendedor.Text = .rdoColumns("RFC")
                Else
                    Txt_RFC_Vendedor.Text = ""
                End If
                If Not IsNull(.rdoColumns("Domicilio")) Then
                    Txt_Domicilio_Vendedor.Text = .rdoColumns("Domicilio")
                Else
                    Txt_Domicilio_Vendedor.Text = ""
                End If
                If Not IsNull(.rdoColumns("Telefono_1")) Then
                    Txt_Telefono_Vendedor.Text = .rdoColumns("Telefono_1")
                Else
                    Txt_Telefono_Vendedor.Text = ""
                End If
                If Not IsNull(.rdoColumns("Comision_Completa")) Then
                    Txt_Comision_Completa.Text = .rdoColumns("Comision_Completa")
                Else
                    Txt_Comision_Completa.Text = ""
                End If
                If Not IsNull(.rdoColumns("Comision_Promocion")) Then
                    Txt_Comision_Oferta.Text = .rdoColumns("Comision_Promocion")
                Else
                    Txt_Comision_Oferta.Text = ""
                End If
                If Not IsNull(.rdoColumns("Comentarios")) Then
                    Txt_Comentarios_Vendedor.Text = .rdoColumns("Comentarios")
                Else
                    Txt_Comentarios_Vendedor.Text = ""
                End If
            End With
        End If
        Rs_Consulta_Cat_Vendedores.Close
    End If
End Sub

Private Sub Grid_Vendedores_RowColChange()
    Grid_Vendedores_Click
End Sub

Private Sub Grid_Zonas_Click()
Dim Rs_Consulta_Cat_Zonas As rdoResultset             'Manejo de registro de la tabla Cat_Usuarios

    'Selecciona la zona que esta en la Tabla
    If Grid_Zonas.Rows > 1 Then
        Txt_Zona_ID.Text = Grid_Zonas.TextMatrix(Grid_Zonas.RowSel, 0)
        Mi_SQL = "SELECT *"
        Mi_SQL = Mi_SQL & " FROM Cat_Zonas"
        Mi_SQL = Mi_SQL & "  WHERE Zona_ID ='" & Txt_Zona_ID.Text & "'"
        Set Rs_Consulta_Cat_Zonas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Zonas.EOF Then
            With Rs_Consulta_Cat_Zonas
                Txt_Zona_ID.Text = .rdoColumns("Zona_ID")
                Txt_Nombre_Zona.Text = .rdoColumns("Nombre")
                Txt_Comentarios_Zona.Text = .rdoColumns("Comentarios")
            End With
        End If
        Rs_Consulta_Cat_Zonas.Close
    End If
End Sub

Private Sub Grid_Zonas_RowColChange()
    Grid_Zonas_Click
End Sub



Private Sub Txt_Comentarios_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comentarios_Gap_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comentarios_Transporte_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comentarios_Usuarios_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comentarios_Zona_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Comision_Completa_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Comision_Completa.Text, True)
End Sub

Private Sub Txt_Comision_Oferta_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Comision_Oferta.Text, True)
End Sub

Private Sub Txt_Contacto_Banco_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Contraseña_Confirmar_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Contraseña_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Depostiar_Ah_Banco_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Estado_Banco_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Habilitar_Click()
On Error GoTo HANDLER
    'Si la columna de Habilitar esta en "S" realiza el proceso de cambio de valor
    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
        If Trim(Txt_Habilitar.Text) = "S" Then
            Txt_Habilitar.Text = "N"
        Else
            Txt_Habilitar.Text = "S"
        End If
        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = Txt_Habilitar.Text
    End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
Private Sub Txt_Habilitar_KeyDown(KeyCode As Integer, Shift As Integer)
If Grid_Accesos_Seguridad.Rows > 1 Then
    If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) = "" Then
        If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
            If KeyCode > 37 Then Grid_Accesos_Seguridad.SetFocus
                If KeyCode = 37 Then
                    If Txt_Habilitar.SelStart = 0 Then
                        Grid_Accesos_Seguridad.SetFocus
                        If Grid_Accesos_Seguridad.Col > 5 Then
                            Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel - 1
                        End If
                    End If
                End If
                If Grid_Accesos_Seguridad.Row > 2 Then
                    If KeyCode = 38 Then Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel - 1
                    If KeyCode = 40 Then
                        If Grid_Accesos_Seguridad.Row < Grid_Accesos_Seguridad.Rows - 1 Or Grid_Accesos_Seguridad.Row = 1 Then
                            Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    If Grid_Accesos_Seguridad.Col >= 6 And Grid_Accesos_Seguridad.Col < 9 Then
                        If KeyCode = 39 Then Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel + 1
                    End If
                Else
                    If KeyCode = 40 Then
                        If Grid_Accesos_Seguridad.Row < Grid_Accesos_Seguridad.Rows - 1 Or Grid_Accesos_Seguridad.Row > 2 Then
                            Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.RowSel + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    If Grid_Accesos_Seguridad.Col >= 6 And Grid_Accesos_Seguridad.Col < 9 Then
                        If KeyCode = 39 Then Grid_Accesos_Seguridad.Col = Grid_Accesos_Seguridad.ColSel + 1
                    End If
                End If
                If Txt_Habilitar.Visible = True Then
                    Txt_Habilitar.SetFocus
                    SendKeys "{Home}+{End}"
                End If
            End If
        Else
            Txt_Habilitar.Visible = False
        End If
    Else
        Txt_Habilitar.Visible = False
    End If
End Sub

Private Sub Txt_Habilitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
            If Trim(Txt_Habilitar.Text) = "S" Then
                Txt_Habilitar.Text = "N"
            Else
                Txt_Habilitar.Text = "S"
            End If
            Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = Txt_Habilitar.Text
        End If
    End If
End Sub

Private Sub Grid_Accesos_Seguridad_DblClick()
    'Valida que si el boton de nuevo o modificar en "Dar de Alta" o "Actualizar" respectivamente para
    'mostrar el control o cambiar de dato en el celda
    If Btn_Nuevo.Caption = "Dar de Alta" Or Btn_Modificar.Caption = "Actualizar" Then
        'Llama el evento click del grid
        Grid_Accesos_Seguridad_Click

        'Si la columna seleccionado esta entre la 6 y la 9 las cuales contienen las opciones de alta, cambio, eliminar y consultar
        If Grid_Accesos_Seguridad.Col > 5 And Grid_Accesos_Seguridad.Col <= 9 Then
            'Valida que sea un renglon de submenu para realizar este proceso
            If Trim(Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 0)) = "" Then
                'Valida que la columna de Habilitar este en "S"
                If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, 5) = "S" Then
                    'Segun el renglon y la columna seleccionada cambia el valor
                    If Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = "S" Then
                        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = "N"
                    Else
                        Grid_Accesos_Seguridad.TextMatrix(Grid_Accesos_Seguridad.RowSel, Grid_Accesos_Seguridad.ColSel) = "S"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Grid_Accesos_Seguridad_EnterCell()
    'Al entrar en una celda del grid llama el evento click del grid si es diferente a la columna del signo (-/+)
    If Grid_Accesos_Seguridad.ColSel > 1 Then Grid_Accesos_Seguridad_Click
End Sub

Private Sub Grid_Accesos_Seguridad_KeyPress(KeyAscii As Integer)
    'Si es el enter llama el evento de doble click del grid para cambiar el valor de la celda
    If KeyAscii = 13 Then
        Grid_Accesos_Seguridad_DblClick
    End If
End Sub

Private Sub Grid_Accesos_Seguridad_LeaveCell()
    'Si el color de la celda es amarillo la cambia a blanco
    If Grid_Accesos_Seguridad.CellBackColor = &HC0FFFF Then Grid_Accesos_Seguridad.CellBackColor = &HFFFFFF
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Collapse_Grid
    'DESCRIPCIÓN: Hace las filas del grid con respecto al height igual a 0
    'PARÁMETROS :
    'CREO       : Oscar Alcantara
    'FECHA_CREO :
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Collapse_Grid()
On Error GoTo HANDLER
Dim Renglon As Integer  'Permite recorrer el grid de los menus

    'Valida que el grid tenga renglones aparte del encabezado
    If Grid_Accesos_Seguridad.Rows > 1 Then
    
        'Coloca el primer renglon en gris y fijo
        Grid_Accesos_Seguridad.FixedRows = 1
        
        'Recorre el grid de los menus para contraer los renglones que dependen de otro
        For Renglon = 1 To Grid_Accesos_Seguridad.Rows - 1
            'Valida si el renglon es un menu y tiene el signo - lo contrae
            If Grid_Accesos_Seguridad.TextMatrix(Renglon, 0) = "-" Then
                'Selecciona la primer columna y el renglon del grid
                Grid_Accesos_Seguridad.Col = 1
                Grid_Accesos_Seguridad.Row = Renglon
                'Llama el evento click del grid para contraer los renglones pertenecientes al seleccionado
                Call Grid_Accesos_Seguridad_Click
            End If
        Next Renglon
    End If
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Grid_Roles_Click
    'DESCRIPCIÓN: Se consulta los menus y submenus que tiene asignado el rol
    '             que el usuario selecciono así como agrega los datos del rol
    '             en los controles correspondientes
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 28-Abril-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Grid_Roles_Click()
Dim Contador_Columnas As Integer                'Indica que columna del grid se esta consultando
Dim Ctl As Control                              'Indica que control es el que se esta consultando en el sistema
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta los menus y submnus que tiene asignados el usuario
    
On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    If Grid_Roles.Rows > 1 Then
        'Asigna los valores correspondientes a los controles de la forma
        Txt_Rol_ID.Text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 0))
        Txt_Nombre_Rol.Text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 1))
        Txt_Comentarios_Rol.Text = Trim(Grid_Roles.TextMatrix(Grid_Roles.RowSel, 2))
        Grid_Accesos_Seguridad.Rows = 0
        Grid_Accesos_Seguridad.Cols = 10
        'Agrega el encabezado el grid_seguridad
        Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "Menu" & Chr(9) & "Submenu" & _
            Chr(9) & "Nombre" & Chr(9) & "Tipo" & Chr(9) & "Habilitar" & _
            Chr(9) & "Alta" & Chr(9) & "Cambio" & Chr(9) & "Elimina" & Chr(9) & "Consulta"
        'Consulta todos los controles que tiene la pantalla MDIFrm_Apl_Principal
        For Each Ctl In MDIFrm_Apl_Principal.Controls
            If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Or UCase(Mid(Ctl.Name, 1, 7)) = UCase("SUBMENU") Then
                'Consulta si el usuario que fue seleccionado tiene habilitado el menu o submenu
                'que se esta consultando
                Mi_SQL = "SELECT * FROM Seguridad_Sistema "
                Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
                Mi_SQL = Mi_SQL & " AND Nombre_Sistema = '" & Ctl.Name & "'"
                Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                'Si no se encuentra el menu que se esta consultando entonces este menu o submenu lo
                'agrega al grid_seguridad y con estatus deshabilitaho
                If Rs_Consulta_Apl_Cat_Accesos.EOF Then
                    If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                        Grid_Accesos_Seguridad.AddItem "-" & Chr(9) & _
                        UCase(Ctl.Caption) & Chr(9) & "" & _
                        Chr(9) & Ctl.Name & Chr(9) & "Encabezado" & _
                        Chr(9) & "N" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                        Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                        'Agrega el color gris a la fila que tiene el encabezado
                        For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                            Grid_Accesos_Seguridad.Col = Contador_Columnas
                            Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                        Next Contador_Columnas
                    Else
                        Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                        "" & Chr(9) & UCase(Ctl.Caption) & _
                        Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & _
                        Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                        
                        If Ctl.Name = "Submenu_Recursos_Humanos" Then
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Importacion Asistencias") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Validacion Tiempo Trabajo") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Solicitud de Permisos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Mantenimiento Asistencias") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Correo de Validacion") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Correo_Validacion" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Asistencia de Empleados") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Exportación a Compaq") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Compaq" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Visor de Registros") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Visor_Registros" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            '**********Catalogos
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Empresas") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empresas" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Departamentos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Motivos Baja") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Puestos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Puestos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Equipo de Identificacion") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Nivel de Estudios") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Dias No Laborales") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Turnos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Turnos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Empleados") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empleados" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Paramentros") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Parametros" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            '********Reportes
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Asistencias") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Historico Faltas y Retardos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Historico de Permisos") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Horas Trabajadas Empleado") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Empleados No Validados") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Empleados de Baja") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
                                "" & Chr(9) & UCase("Empleados de Alta") & _
                                Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta" & Chr(9) & "SubMenu" & _
                                Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
                            
                        End If
                    End If
                'Si lo encuentra entonces agrega el menu o submenu al grid_seguridad con el estatus que tiene
                'asignado
                Else
                    If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
                        Grid_Accesos_Seguridad.AddItem "-" & _
                        Chr(9) & UCase(Ctl.Caption) & Chr(9) & "" & _
                        Chr(9) & Ctl.Name & Chr(9) & "Encabezado" & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Habilitar") & _
                        Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                        'Agrega el color gris a la fila que tiene el encabezado
                        Grid_Accesos_Seguridad.Row = Grid_Accesos_Seguridad.Rows - 1
                        For Contador_Columnas = 0 To Grid_Accesos_Seguridad.Cols - 1
                            Grid_Accesos_Seguridad.Col = Contador_Columnas
                            Grid_Accesos_Seguridad.CellBackColor = vbButtonFace
                        Next Contador_Columnas
                    Else
                        Grid_Accesos_Seguridad.AddItem "" & Chr(9) & "" & _
                        Chr(9) & UCase(Ctl.Caption) & _
                        Chr(9) & Ctl.Name & Chr(9) & "SubMenu" & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Habilitar") & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Alta") & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Cambio") & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Eliminar") & _
                        Chr(9) & Rs_Consulta_Apl_Cat_Accesos.rdoColumns("Consultar")
                        
                        If Ctl.Name = "Submenu_Recursos_Humanos" Then
                            Agrega_Submenus_RH
                        End If
                        
                    End If
                End If
                Rs_Consulta_Apl_Cat_Accesos.Close
            End If 'Menu/Submenu
        Next Ctl
        'Configura el tamaño de las columnas del grid_accesos_seguridad
        If Grid_Accesos_Seguridad.Rows > 1 Then
            Grid_Accesos_Seguridad.FixedRows = 1
            Grid_Accesos_Seguridad.ColWidth(0) = 200  '-
            Grid_Accesos_Seguridad.ColWidth(1) = 2150 'Menu
            Grid_Accesos_Seguridad.ColWidth(2) = 2700 'SubMenu
            Grid_Accesos_Seguridad.ColWidth(3) = 0    'Nombre Menu/Submenu
            Grid_Accesos_Seguridad.ColWidth(4) = 0    'Tipo
            Grid_Accesos_Seguridad.ColWidth(5) = 900  'Habilitar
            Grid_Accesos_Seguridad.ColAlignment(5) = 3
            Grid_Accesos_Seguridad.ColWidth(6) = 700  'Alta
            Grid_Accesos_Seguridad.ColAlignment(6) = 3
            Grid_Accesos_Seguridad.ColWidth(7) = 700  'Cambio
            Grid_Accesos_Seguridad.ColAlignment(7) = 3
            Grid_Accesos_Seguridad.ColWidth(8) = 700  'Eliminar
            Grid_Accesos_Seguridad.ColAlignment(8) = 3
            Grid_Accesos_Seguridad.ColWidth(9) = 700  'Consultar
            Grid_Accesos_Seguridad.ColAlignment(9) = 3
            'Coloca como contraidos los renglones del grid de menus
            Call Collapse_Grid
        End If
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Acceso_Seguridad_Click()
    On Error GoTo HANDLER
        If Fra_Roles_Sistema.Visible = True Then
            Fra_Acceso_Sistema_Rol.Visible = True
            Fra_Roles_Sistema.Visible = False
            Btn_Acceso_Seguridad.Caption = "Roles"
        Else
            Fra_Acceso_Sistema_Rol.Visible = False
            Fra_Roles_Sistema.Visible = True
            Btn_Acceso_Seguridad.Caption = "Control de Acceso"
        End If
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Asigna_Item_Combo
    'DESCRIPCIÓN: 'Selecciona un registro especifico de un combo
    'PARÁMETROS:
    '             1. Valor: Clave oculta del registro en el combo
    '             2. Cmb_Dato: Combo que contiene la lista de registros
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Asigna_Item_Combo(Valor As String, Cmb_Dato As ComboBox, Optional Tam_Format As Integer = 5)
    Dim I As Integer
    Dim format_string As String
    
    format_string = "00000"
    If Tam_Format <> 5 Then
        format_string = ""
        For I = 1 To Tam_Format
            format_string = format_string & "0"
        Next
    End If
    Cmb_Dato.ListIndex = -1
    For I = 0 To Cmb_Dato.ListCount - 1
        If Cmb_Dato.List(I) = Valor Or Format(Cmb_Dato.ItemData(I), format_string) = Valor Then
            Cmb_Dato.ListIndex = I
            Exit For
        End If
    Next I
End Sub

Private Sub Grid_Bancos_Click()
Dim Rs_Consulta_Bancos As rdoResultset  '#  Obtine los datos del Banco
Dim Rs_Consulta_Adm_Movimientos As rdoResultset '#  Consulta el saldo del banco
    
    'valida que existan registros en el grid
    If Grid_Bancos.Rows > 1 Then
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Txt_Saldo_Banco.Text = 0
        Mi_SQL = "SELECT * FROM Cat_Bancos"
        Mi_SQL = Mi_SQL & " WHERE Banco_ID='" & Trim(Grid_Bancos.TextMatrix(Grid_Bancos.RowSel, 0)) & "'"
        Set Rs_Consulta_Bancos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        '#  Si obtuvo el registro del Catalogo de Bancos
        If Not Rs_Consulta_Bancos.EOF Then
            With Rs_Consulta_Bancos
                Txt_Banco_ID.Text = .rdoColumns("Banco_ID")
                Txt_Nombre_Banco.Text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Sucursal")) Then Txt_Sucursal.Text = .rdoColumns("Sucursal")
                If Not IsNull(.rdoColumns("Contacto")) Then Txt_Contacto_Banco.Text = .rdoColumns("Contacto")
                Txt_No_Cuenta_Banco.Text = .rdoColumns("No_Cuenta")
                If Not IsNull(.rdoColumns("Fiscal")) Then
                    Cmb_Cuenta_Fiscal.Text = .rdoColumns("Fiscal")
                Else
                    Cmb_Cuenta_Fiscal.Text = "SI"
                End If
                If Not IsNull(.rdoColumns("Estatus")) Then
                    If Trim(.rdoColumns("Estatus")) <> "" Then
                        Cmb_Estatus_Banco.Text = .rdoColumns("Estatus")
                    Else
                        Cmb_Estatus_Banco.Text = "ACTIVO"
                    End If
                Else
                    Cmb_Estatus_Banco.Text = "ACTIVO"
                End If
                If Not IsNull(.rdoColumns("Depositar_Ah")) Then Txt_Depostiar_Ah_Banco.Text = .rdoColumns("Depositar_Ah")
                If Not IsNull(.rdoColumns("Gap")) Then Txt_Gap_Banco.Text = .rdoColumns("Gap")
                If Not IsNull(.rdoColumns("Transporte")) Then Txt_Transporte_Banco.Text = .rdoColumns("Transporte")
                Select Case .rdoColumns("Empresa")
                    Case 1
                        Cmb_Empresa.Text = "NATURAL HEALTH"
                    Case 2
                        Cmb_Empresa.Text = "PRONACEN"
                    Case 4
                        Cmb_Empresa.Text = "GRUPO NH"
                End Select
                If Not IsNull(.rdoColumns("Formato")) Then Cmb_Formato.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(.rdoColumns("Formato"), Cmb_Formato)
                If Not IsNull(.rdoColumns("Saldo")) Then Txt_Saldo_Banco.Text = .rdoColumns("Saldo")
            End With
        End If
        Rs_Consulta_Bancos.Close
        Txt_Saldo_Banco.Text = Format(Txt_Saldo_Banco.Text, "#,##0.00")
    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Banco
'DESCRIPCION: Da de alta el banco en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 13-Mayo-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Banco()
Dim Rs_Alta_Banco As rdoResultset  'Manejo de Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Banco = Conectar_Ayudante.Recordset_Agregar("Cat_Bancos")
    With Rs_Alta_Banco
        .AddNew
            Txt_Banco_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Bancos", "Banco_ID"), "00000")
            .rdoColumns("Banco_ID") = Txt_Banco_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Banco.Text))
            .rdoColumns("Sucursal") = Trim(Txt_Sucursal.Text)
            .rdoColumns("Contacto") = Trim(UCase(Txt_Contacto_Banco.Text))
            .rdoColumns("No_Cuenta") = Trim(Txt_No_Cuenta_Banco.Text)
            .rdoColumns("Fiscal") = Cmb_Cuenta_Fiscal.Text
            .rdoColumns("Estatus") = Cmb_Estatus_Banco.Text
            .rdoColumns("Depositar_Ah") = Trim(UCase(Txt_Depostiar_Ah_Banco.Text))
            .rdoColumns("Gap") = Trim(UCase(Txt_Gap_Banco.Text))
            .rdoColumns("Transporte") = Trim(UCase(Txt_Transporte_Banco.Text))
            Select Case Cmb_Empresa.Text
                Case "NATURAL HEALTH"
                    .rdoColumns("Empresa") = 1
                Case "PRONACEN"
                    .rdoColumns("Empresa") = 2
                Case "GRUPO NH"
                    .rdoColumns("Empresa") = 4
            End Select
            .rdoColumns("Formato") = Trim(Cmb_Formato.Text)
            .rdoColumns("Saldo") = 0
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Rs_Alta_Banco.Close
    MsgBox "El banco ha sido dado de alta", vbInformation
    If Grid_Bancos.Rows = 0 Then
        Grid_Bancos.AddItem "banco ID" & Chr(9) & "Nombre" & Chr(9) & "No. Cuenta"
    End If
    Grid_Bancos.AddItem Txt_Banco_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Banco.Text)) & Chr(9) & Txt_No_Cuenta_Banco.Text
    Grid_Bancos.FixedRows = 1
    Grid_Bancos.ColWidth(0) = 1200
    Grid_Bancos.ColAlignment(0) = flexAlignCenterCenter
    Grid_Bancos.ColWidth(1) = 3000
    Grid_Bancos.ColAlignment(1) = flexAlignLeftCenter
    Grid_Bancos.ColWidth(2) = 2000
    Grid_Bancos.ColAlignment(2) = flexAlignLeftCenter
    Btn_Salir_Click
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Catalogo_Bancos", Frm_Cat_Generales)
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Curso
'DESCRIPCION: Da de alta el curso en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Curso()
Dim Rs_Alta_Curso As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Curso = Conectar_Ayudante.Recordset_Agregar("Cat_Cursos")
    With Rs_Alta_Curso
        .AddNew
            Txt_Curso_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Cursos", "Curso_ID"), "00000")
            .rdoColumns("Curso_ID") = Txt_Curso_ID.Text
            .rdoColumns("Nombre") = Trim(Txt_Nombre_Curso.Text)
            .rdoColumns("Horas") = Val(Txt_Horas_Curso.Text)
            .rdoColumns("Tipo") = Cmb_Tipo_Curso.Text
            .rdoColumns("Instructor") = Trim(Txt_Instructor_Curso.Text)
            .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Curso.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Curso.Close
    MsgBox "El curso ha sido dado de alta", vbInformation
    Fra_Generales_Cursos.Enabled = False
    Fra_Grid_Cursos.Enabled = True
    'Pone el encabezado al grid de Tipo Producto
    If Grid_Cursos.Rows = 0 Then
        Grid_Cursos.AddItem "Curso ID" & Chr(9) & "Nombre" & Chr(9) & "Horas"
    End If
    'Agrega los datos
    Grid_Cursos.AddItem Txt_Curso_ID.Text & Chr(9) & Trim(Txt_Nombre_Curso.Text) & Chr(9) & Val(Txt_Horas_Curso.Text)
    Grid_Cursos.FixedRows = 1
    'Da formato al grid
    Grid_Cursos.ColWidth(0) = 1000
    Grid_Cursos.ColWidth(1) = 4500
    Grid_Cursos.ColWidth(2) = 1000
    Btn_Salir_Click
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Unidades
'DESCRIPCIÓN: Da de alta el registro de la Unidad en la base de datos
'PARAMETROS :
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Unidades()
Dim Rs_Alta_Unidades As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Prepara el recordset para la alta
    Set Rs_Alta_Unidades = Conectar_Ayudante.Recordset_Agregar("Cat_Unidades")
    With Rs_Alta_Unidades
        .AddNew
            Txt_Unidad_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Unidades", "Unidad_ID"), "00000")
            .rdoColumns("Unidad_ID") = Txt_Unidad_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Unidad.Text))
            .rdoColumns("Nombre_Corto") = Trim(UCase(Txt_Nombre_Corto_Unidad.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Unidad.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Fra_Generales_Unidades.Enabled = False
    Fra_Grid_Unidades.Enabled = True
    'Pone el encabezado al grid de Tipo Producto
    If Grid_Unidades.Rows = 0 Then
        Grid_Unidades.AddItem "Unidad ID" & Chr(9) & "Nombre" & Chr(9) & "Nombre Corto"
    End If
    'Agrega los datos
    Grid_Unidades.AddItem Txt_Unidad_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Unidad.Text)) & Chr(9) & Trim(UCase(Txt_Nombre_Corto_Unidad.Text))
    Grid_Unidades.FixedRows = 1
    'Da formato al grid
    Grid_Unidades.ColWidth(0) = 1500
    Grid_Unidades.ColWidth(1) = 2500
    Grid_Unidades.ColWidth(2) = 2500
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    MsgBox "Unidad dada de Alta", vbInformation
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Transportes
'DESCRIPCION: Da de alta el registro de Transportes en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Transportes()
Dim Rs_Alta_Transportes As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Transportes = Conectar_Ayudante.Recordset_Agregar("Cat_Transportes")
    With Rs_Alta_Transportes
        .AddNew
            Txt_Transporte_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Transportes", "Transporte_ID"), "00000")
            .rdoColumns("Transporte_ID") = Txt_Transporte_ID.Text
            If Cmb_Zona.ListIndex > -1 Then
                .rdoColumns("Zona_ID") = Format(Cmb_Zona.ItemData(Cmb_Zona.ListIndex), "00000")
            End If
            .rdoColumns("Nombre") = Trim(Txt_Nombre_Transporte.Text)
            .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Transporte.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Transportes.Close
    'Pone el encabezado al grid de Transportes
    If Grid_Transportes.Rows = 0 Then
        Grid_Transportes.AddItem "Transporte ID" & Chr(9) & "Nombre" & Chr(9) & "Zona"
    End If
    'Agrega los datos
    Grid_Transportes.AddItem Txt_Transporte_ID.Text & Chr(9) & Trim(Txt_Nombre_Transporte.Text) & Chr(9) & Trim(Cmb_Zona.Text)
    Grid_Transportes.FixedRows = 1
    'Da formato al grid
    Grid_Transportes.ColWidth(0) = 1000
    Grid_Transportes.ColWidth(1) = 3000
    Grid_Transportes.ColWidth(2) = 2600
    MsgBox "Transporte dado de Alta", vbInformation
    Btn_Salir_Click
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Gerencia
'DESCRIPCION: Da de alta el Tipo de Pago en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 21-Marzo-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Gerencia()
Dim Rs_Alta_Gerencia As rdoResultset           'Manejo del Registro
Dim Rs_Consulta_Supervisores As rdoResultset
    
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Prepara el recordset para la alta
    Set Rs_Alta_Gerencia = Conectar_Ayudante.Recordset_Agregar("Cat_Gerencias")
    With Rs_Alta_Gerencia
        .AddNew
            Txt_Gerencia_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Gerencias", "Gerencia_ID"), "00000")
            .rdoColumns("Gerencia_ID") = Txt_Gerencia_ID.Text
            .rdoColumns("Nombre") = Trim(Txt_Nombre_Gerencia.Text)
            .rdoColumns("Supervisor_ID") = Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000")
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Gerencia.Close
    'Actualiza los empleados con la gerencia del supervisor
    Mi_SQL = "UPDATE Cat_Empleados SET Gerencia_UAP='" & Trim(Txt_Gerencia_ID.Text) & "'"
    Mi_SQL = Mi_SQL & " WHERE (Empleado_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "')"
    Conexion_Base.Execute Mi_SQL
    'Actualiza los empleados con la gerencia del supervisor de los siguientes niveles
    Mi_SQL = "SELECT DISTINCT Empleado_ID FROM Cat_Empleados WHERE Supervisor_ID='" & Format(Cmb_Supervisor_Gerencia.ItemData(Cmb_Supervisor_Gerencia.ListIndex), "00000") & "'"
    Set Rs_Consulta_Supervisores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Supervisores.EOF
        Mi_SQL = "UPDATE Cat_Empleados SET Gerencia_UAP='" & Trim(Txt_Gerencia_ID.Text) & "'"
        Mi_SQL = Mi_SQL & " WHERE Supervisor_ID='" & Rs_Consulta_Supervisores.rdoColumns("Empleado_ID") & "'"
        Conexion_Base.Execute Mi_SQL
        Rs_Consulta_Supervisores.MoveNext
    Wend
    Rs_Consulta_Supervisores.Close
    Conexion_Base.CommitTrans
    MsgBox "La gerencia ha sido dada de alta", vbInformation
    Btn_Salir_Click
    Consulta_Gerencia ""
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Alta_Marcas
    'DESCRIPCION: Da de alta el registro de Marcas en la base de datos
    'PARAMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 22-Agosto-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'*******************************************************************************
Private Sub Alta_Marcas()
Dim Rs_Alta_Marcas As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Marcas = Conectar_Ayudante.Recordset_Agregar("Cat_Marcas")
    With Rs_Alta_Marcas
        .AddNew
            Txt_Marca_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Marcas", "Marca_ID"), "00000")
            .rdoColumns("Marca_ID") = Txt_Marca_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Marca.Text))
            .rdoColumns("Nombre_Corto") = Trim(UCase(Txt_Nombre_Corto_Marca.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Marcas.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Fra_Generales_Marcas.Enabled = False
    Fra_Grid_Marcas.Enabled = True
    'Pone el encabezado al grid de marcas
    If Grid_Marcas.Rows = 0 Then
        Grid_Marcas.AddItem "Marca ID" & Chr(9) & "Nombre" & Chr(9) & "Nombre Corto" & Chr(9) & "Comentarios"
    End If
    'Agrega los datos
    Grid_Marcas.AddItem Txt_Marca_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Marca.Text)) & Chr(9) & Trim(UCase(Txt_Nombre_Corto_Marca.Text)) & Chr(9) & Trim(UCase(Txt_Comentarios_Marcas.Text))
    Grid_Marcas.FixedRows = 1
    'Da formato al grid
    Grid_Marcas.ColWidth(0) = 1000
    Grid_Marcas.ColWidth(1) = 3000
    Grid_Marcas.ColWidth(2) = 2000
    Grid_Marcas.ColAlignment(2) = flexAlignLeftCenter
    Grid_Marcas.ColWidth(3) = 0
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    MsgBox "La marca ha sido dada de alta", vbExclamation
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Secciones
'DESCRIPCION: Da de alta el registro de Seccion en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 28-Septiembre-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Secciones()
Dim Rs_Alta_Secciones As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Secciones = Conectar_Ayudante.Recordset_Agregar("Cat_Secciones")
    With Rs_Alta_Secciones
        .AddNew
            Txt_Seccion_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Secciones", "Seccion_ID"), "00000")
            .rdoColumns("Seccion_ID") = Txt_Seccion_ID.Text
            .rdoColumns("Supervisor_ID") = Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000")
            .rdoColumns("Clave") = Trim(Txt_Seccion_Clave.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    'Actualiza los empleados con la sección del supervisor
    Mi_SQL = "UPDATE Cat_Empleados"
    Mi_SQL = Mi_SQL & " SET Nomipaq_ID='" & Trim(Txt_Seccion_Clave.Text) & "'"
    Mi_SQL = Mi_SQL & " WHERE (Empleado_ID='" & Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " OR Supervisor_ID='" & Format(Cmb_Seccion_Supervisor.ItemData(Cmb_Seccion_Supervisor.ListIndex), "00000") & "')"
    Conexion_Base.Execute Mi_SQL
    MsgBox "La Seccion ha sido dada de Alta", vbInformation
    Consulta_Secciones ""
    Btn_Salir_Click
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbExclamation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Operadores
    'DESCRIPCIÓN: Da de alta el registro de Operadores en la base de datos
    'PARÁMETROS :
    'CREO       : Rafael Muñoz
    'FECHA_CREO : 07-Feb-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Alta_Operadores()
Dim Rs_Alta_Operadores As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Operadores = Conectar_Ayudante.Recordset_Agregar("Cat_Operadores")
    With Rs_Alta_Operadores
        .AddNew
            Txt_Operador_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Operadores", "Operador_ID"), "00000")
            .rdoColumns("Operador_ID") = Txt_Operador_ID.Text
            .rdoColumns("Tipo") = Cmb_Tipo.Text
            .rdoColumns("Estatus") = Mid(Cmb_Estatus_Operador, 1, 1)
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Operador.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Operadores.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Fra_Generales_Operadores.Enabled = False
    Fra_Grid_Operadores.Enabled = True
    'Pone el encabezado al grid de Operadores
    If Grid_Operadores.Rows = 0 Then
        Grid_Operadores.AddItem "Operador ID" & Chr(9) & "Nombre" & Chr(9) & "Tipo"
    End If
    'Agrega los datos
    Grid_Operadores.AddItem Txt_Operador_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Operador.Text)) & Chr(9) & Cmb_Tipo
    Grid_Operadores.FixedRows = 1
    'Da formato al grid
    Grid_Operadores.ColWidth(0) = 1000
    Grid_Operadores.ColWidth(1) = 4000
    Grid_Operadores.ColWidth(2) = 1500
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    MsgBox "Operador dado de Alta", vbExclamation
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Zonas
'DESCRIPCION: Da de alta el registro de la zona en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Zonas()
Dim Rs_Alta_Zonas As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Zonas = Conectar_Ayudante.Recordset_Agregar("Cat_Zonas")
    With Rs_Alta_Zonas
        .AddNew
            Txt_Zona_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Zonas", "Zona_ID"), "00000")
            .rdoColumns("Zona_ID") = Txt_Zona_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Zona.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Zona.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Zonas.Close
    MsgBox "Zona dada de Alta", vbInformation
    'Pone el encabezado al grid de Zonas
    If Grid_Zonas.Rows = 0 Then
        Grid_Zonas.AddItem "Zona ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
    End If
    'Agrega los datos
    Grid_Zonas.AddItem Txt_Zona_ID.Text & Chr(9) & Trim(Txt_Nombre_Zona.Text) & Chr(9) & Trim(Txt_Comentarios_Zona.Text)
    Grid_Zonas.FixedRows = 1
    'Da formato al grid
    Grid_Zonas.ColWidth(0) = 1500
    Grid_Zonas.ColWidth(1) = 2500
    Grid_Zonas.ColWidth(2) = 2000
    Btn_Salir_Click
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbExclamation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Giro
    'DESCRIPCIÓN: Da de alta el registro del giro en la base de datos
    'PARÁMETROS :
    'CREO       : Rafael Muñoz
    'FECHA_CREO : 08-Ene-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Alta_Giro()
Dim Rs_Alta_Cat_Giros_Empresariales As rdoResultset           'Manejo del Registro
    
    On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Prepara el recordset para la alta
    Set Rs_Alta_Cat_Giros_Empresariales = Conectar_Ayudante.Recordset_Agregar("Cat_Giros_Empresariales")
    With Rs_Alta_Cat_Giros_Empresariales
        .AddNew
            Txt_Giro_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Giros_Empresariales", "Giro_ID"), "00000")
            .rdoColumns("Giro_ID") = Txt_Giro_ID.Text
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Giro.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Giros.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Fra_Generales_Giros.Enabled = False
    Fra_Grid_Giros.Enabled = True
    
    'Pone el encabezado al grid de giros empresariales
    If Grid_Giros.Rows = 0 Then
        Grid_Giros.AddItem "Tipo Cliente ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
    End If
    'Agrega los datos
    Grid_Giros.AddItem Txt_Giro_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Giro.Text)) & Chr(9) & Trim(UCase(Txt_Comentarios_Giros.Text))
    Grid_Giros.FixedRows = 1
    'Da formato al grid
    Grid_Giros.ColWidth(0) = 1200
    Grid_Giros.ColWidth(1) = 3000
    Grid_Giros.ColWidth(2) = 2300
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    MsgBox "Tipo Cliente dado de Alta", vbExclamation
    Exit Sub
    
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Gaps
'DESCRIPCION: Da de alta el registro de Gaps en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Alta_Gaps()
Dim Rs_Alta_Gaps As rdoResultset            'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Gaps = Conectar_Ayudante.Recordset_Agregar("Cat_Gaps")
    With Rs_Alta_Gaps
        .AddNew
            Txt_Gap_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Gaps", "Gap_ID"), "00000")
            .rdoColumns("Gap_ID") = Txt_Gap_ID.Text
            .rdoColumns("Nombre") = Trim(Txt_Nombre_Gap.Text)
            .rdoColumns("Comentarios") = Trim(Txt_Comentarios_Gap.Text)
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Gaps.Close
    MsgBox "La tripulación ha sido dada de Alta", vbInformation
    'Da formato al grid
    If Grid_Gaps.Rows = 0 Then
        Grid_Gaps.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Comentarios"
        Grid_Gaps.ColWidth(0) = 1000
        Grid_Gaps.ColAlignment(0) = flexAlignCenterCenter
        Grid_Gaps.ColWidth(1) = 3000
        Grid_Gaps.ColAlignment(1) = flexAlignLeftCenter
        Grid_Gaps.ColWidth(2) = 2500
        Grid_Gaps.ColAlignment(2) = flexAlignLeftCenter
    End If
    'Agrega los datos
    Grid_Gaps.AddItem Txt_Gap_ID.Text & Chr(9) & Trim(Txt_Nombre_Gap.Text) & Chr(9) & Trim(Txt_Comentarios_Gap.Text)
    Grid_Gaps.FixedRows = 1
    Btn_Salir_Click
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbExclamation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Vendedores
    'DESCRIPCIÓN: Da de alta el registro del vendedor en la base de datos
    'PARÁMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 14-Marzo-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Alta_Vendedores()
Dim Rs_Alta_Cat_Vendedores As rdoResultset            'Manejo del Registro
    
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Prepara el recordset para la alta
    Set Rs_Alta_Cat_Vendedores = Conectar_Ayudante.Recordset_Agregar("Cat_Vendedores")
    With Rs_Alta_Cat_Vendedores
        .AddNew
            Txt_Vendedor_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Vendedores", "Vendedor_ID"), "00000")
            .rdoColumns("Vendedor_ID") = Txt_Vendedor_ID.Text
            .rdoColumns("Estatus") = Mid(Cmb_Estatus_Vendedor.Text, 1, 1)
            .rdoColumns("Clave") = Trim(UCase(Txt_Clave_Vendedor.Text))
            .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Vendedor.Text))
            If Cmb_Gap_Vendedor.ListIndex > -1 Then
                .rdoColumns("Gap_ID") = Format(Cmb_Gap_Vendedor.ItemData(Cmb_Gap_Vendedor.ListIndex), "00000")
            End If
            .rdoColumns("RFC") = Trim(UCase(Txt_RFC_Vendedor.Text))
            .rdoColumns("Domicilio") = Trim(UCase(Txt_Domicilio_Vendedor.Text))
            .rdoColumns("Telefono_1") = Trim(Txt_Telefono_Vendedor.Text)
            .rdoColumns("Comision_Completa") = Val(Txt_Comision_Completa.Text)
            .rdoColumns("Comision_Promocion") = Val(Txt_Comision_Oferta.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Vendedor.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Format(Now, "MM/dd/yyyy")
        .Update
    End With
    Fra_Generales_Vendedores.Enabled = False
    Fra_Grid_Vendedores.Enabled = True
    'Pone el encabezado al grid de Vendedores
    If Grid_Vendedores.Rows = 0 Then
        Grid_Vendedores.AddItem "Vendedor ID" & Chr(9) & "Nombre" & Chr(9) & "Clave"
    End If
    Rs_Alta_Cat_Vendedores.Close
    'Agrega los datos
    Grid_Vendedores.AddItem Txt_Vendedor_ID.Text & Chr(9) & Trim(UCase(Txt_Nombre_Vendedor.Text)) & Chr(9) & Trim(UCase(Txt_Clave_Vendedor.Text))
    Grid_Vendedores.FixedRows = 1
    'Da formato al grid
    Grid_Vendedores.ColWidth(0) = 1000
    Grid_Vendedores.ColWidth(1) = 4000
    Grid_Vendedores.ColAlignment(1) = flexAlignLeftCenter
    Grid_Vendedores.ColWidth(2) = 1600
    Grid_Vendedores.ColAlignment(2) = flexAlignLeftCenter
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    MsgBox "Vendedor dada de alta", vbInformation
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Modifica_Banco
'DESCRIPCION: Actualiza los datos del banco en la base de datos
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 13-Mayo-2009
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Modifica_Banco()
Dim Rs_Modifica_Banco As rdoResultset  'Manejo del Registro
    
On Error GoTo HANDLER
    'Modifica la banco actual en ventana
    Mi_SQL = "SELECT * FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Banco_ID='" & Txt_Banco_ID.Text & "'"
    Set Rs_Modifica_Banco = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Banco.EOF Then
        With Rs_Modifica_Banco
            .Edit
                .rdoColumns("Nombre") = Trim(UCase(Txt_Nombre_Banco.Text))
                .rdoColumns("Sucursal") = Trim(Txt_Sucursal.Text)
                .rdoColumns("Contacto") = Trim(UCase(Txt_Contacto_Banco.Text))
                .rdoColumns("No_Cuenta") = Trim(Txt_No_Cuenta_Banco.Text)
                .rdoColumns("Fiscal") = Cmb_Cuenta_Fiscal.Text
                .rdoColumns("Estatus") = Cmb_Estatus_Banco.Text
                .rdoColumns("Depositar_Ah") = Trim(UCase(Txt_Depostiar_Ah_Banco.Text))
                .rdoColumns("Gap") = Trim(UCase(Txt_Gap_Banco.Text))
                .rdoColumns("Transporte") = Trim(UCase(Txt_Transporte_Banco.Text))
                Select Case Cmb_Empresa.Text
                    Case "NATURAL HEALTH"
                        .rdoColumns("Empresa") = 1
                    Case "PRONACEN"
                        .rdoColumns("Empresa") = 2
                    Case "GRUPO NH"
                        .rdoColumns("Empresa") = 4
                End Select
                .rdoColumns("Formato") = Trim(Cmb_Formato.Text)
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modifica_Banco.Close
    MsgBox "El banco ha sido modificado", vbInformation
    Grid_Bancos.TextMatrix(Grid_Bancos.RowSel, 0) = Trim(UCase(Txt_Banco_ID.Text))
    Grid_Bancos.TextMatrix(Grid_Bancos.RowSel, 1) = Trim(UCase(Txt_Nombre_Banco.Text))
    Grid_Bancos.TextMatrix(Grid_Bancos.RowSel, 2) = Trim(UCase(Txt_No_Cuenta_Banco.Text))
    Btn_Salir_Click
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Catalogo_Bancos", Frm_Cat_Generales)
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Tiempos_Muertos
    'DESCRIPCIÓN:           Consulta Tiempos Muertos
    'PARÁMETROS :           Nombre: Indica el nombre del Tiempo Muerto a buscar
    'CREO       :           Julio Cruz
    'FECHA_CREO :           10-Dic-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'******************************************************************************
Public Sub Consulta_Tiempos_Muertos(Descripcion As String)
Dim Rs_Consulta_Cat_Tiempos_Muertos As rdoResultset         'Manejo del Registro
    
On Error GoTo HANDLER
    'Consulta el Tiempo Muerto
    Mi_SQL = "SELECT Tiempo_ID,Descripcion, Comentarios"
    Mi_SQL = Mi_SQL & " FROM Cat_Tiempos_Muertos"
    Mi_SQL = Mi_SQL & " WHERE Descripcion LIKE '%" & Descripcion & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Descripcion"
    Set Rs_Consulta_Cat_Tiempos_Muertos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Cat_Tiempos_Muertos.EOF Then
        'Coloca un encabezado en el grid
        Grid_Tiempos_Muertos.Rows = 0
        Grid_Tiempos_Muertos.AddItem "Tiempo Muerto ID" & Chr(9) & "Descripción"
        While Not Rs_Consulta_Cat_Tiempos_Muertos.EOF
            Grid_Tiempos_Muertos.AddItem Rs_Consulta_Cat_Tiempos_Muertos.rdoColumns("Tiempo_ID") & Chr(9) & Rs_Consulta_Cat_Tiempos_Muertos.rdoColumns("Descripcion")
            Grid_Tiempos_Muertos.FixedRows = 1
            Rs_Consulta_Cat_Tiempos_Muertos.MoveNext
        Wend
        Grid_Tiempos_Muertos.ColWidth(0) = 1500
        Grid_Tiempos_Muertos.ColWidth(1) = 4990
        Grid_Tiempos_Muertos.ColWidth(2) = 0
        Grid_Tiempos_Muertos.ColWidth(3) = 0
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Cat_Tiempos_Muertos.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Consulta_Tipos_Notas_Credito
    'DESCRIPCION: Consulta los tipos de notas de crédito de acuerdo al parámetro
    'PARAMETROS : Nombre: Indica el nombre del tipo de nota de crédito a buscar
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 02-Junio-2009
    'MODIFICO   :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'******************************************************************************
Public Sub Consulta_Tipos_Notas_Credito(Descripcion As String)
Dim Rs_Consulta_Cat_Tipos_Notas_Credito As rdoResultset
    
On Error GoTo HANDLER
    Mi_SQL = "SELECT * FROM Cat_Tipos_Notas_Credito"
    Mi_SQL = Mi_SQL & " WHERE Descripcion LIKE '%" & Descripcion & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Descripcion"
    Set Rs_Consulta_Cat_Tipos_Notas_Credito = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Tipos_Notas_Credito.EOF Then
        Grid_Tipos_Notas_Credito.Rows = 0
        Grid_Tipos_Notas_Credito.AddItem "Tipo Nota ID" & Chr(9) & "Descripción"
        While Not Rs_Consulta_Cat_Tipos_Notas_Credito.EOF
            Grid_Tipos_Notas_Credito.AddItem Rs_Consulta_Cat_Tipos_Notas_Credito.rdoColumns("Tipo_Nota_Credito_ID") & Chr(9) & Rs_Consulta_Cat_Tipos_Notas_Credito.rdoColumns("Descripcion")
            Grid_Tipos_Notas_Credito.FixedRows = 1
            Rs_Consulta_Cat_Tipos_Notas_Credito.MoveNext
        Wend
        Grid_Tipos_Notas_Credito.ColWidth(0) = 1500
        Grid_Tipos_Notas_Credito.ColWidth(1) = 4990
        Grid_Tipos_Notas_Credito.ColWidth(2) = 0
        Grid_Tipos_Notas_Credito.ColWidth(3) = 0
    End If
    Rs_Consulta_Cat_Tipos_Notas_Credito.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Alta_Tiempos_Muertos
    'DESCRIPCIÓN:           Da de alta el registro del Tiempos Muertos en la base de datos
    'PARÁMETROS :
    'CREO       :           Julio Cruz
    'FECHA_CREO :           10-Dic-2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Alta_Tiempos_Muertos()
Dim Rs_Alta_Tiempos_Muertos As rdoResultset           'Manejo del Registro
    
    On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Prepara el recordset para la alta
    Set Rs_Alta_Tiempos_Muertos = Conectar_Ayudante.Recordset_Agregar("Cat_Tiempos_Muertos")
    With Rs_Alta_Tiempos_Muertos
        .AddNew
            Txt_Tiempo_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tiempos_Muertos", "Tiempo_ID"), "00000")
            .rdoColumns("Tiempo_ID") = Txt_Tiempo_ID.Text
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Descripcion_Tiempos.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Tiempos.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Fra_Generales_Tiempos_Muertos.Enabled = False
    Fra_Grid_Tiempos_Muertos.Enabled = True
    
    'Pone el encabezado al grid de Tipo Producto
    If Grid_Tiempos_Muertos.Rows = 0 Then
        Grid_Tiempos_Muertos.AddItem "Tiempo ID" & Chr(9) & "Descripcion"
    End If
    'Agrega los datos
    Grid_Tiempos_Muertos.AddItem Txt_Tiempo_ID.Text & Chr(9) & Trim(UCase(Txt_Descripcion_Tiempos.Text))
    Grid_Tiempos_Muertos.FixedRows = 1
    'Da formato al grid
    Grid_Tiempos_Muertos.ColWidth(0) = 1500
    Grid_Tiempos_Muertos.ColWidth(1) = 4990
    Grid_Tiempos_Muertos.ColWidth(2) = 0
    Grid_Tiempos_Muertos.ColWidth(3) = 0
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    MsgBox "Tiempo Muerto dado de Alta", vbInformation
    Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Alta_Tiempos_Muertos
    'DESCRIPCION: Da de alta el registro de tipos de notas de crédito
    'PARAMETROS :
    'CREO       : Sergio Ulises Durán Hernández
    'FECHA_CREO : 02-Junio-2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACION:
'*******************************************************************************
Public Sub Alta_Tipos_Notas_Credito()
Dim Rs_Alta_Tipos_Notas_Credito As rdoResultset           'Manejo del Registro
    
On Error GoTo HANDLER
    'Prepara el recordset para la alta
    Set Rs_Alta_Tipos_Notas_Credito = Conectar_Ayudante.Recordset_Agregar("Cat_Tipos_Notas_Credito")
    With Rs_Alta_Tipos_Notas_Credito
        .AddNew
            Txt_Tipo_Nota_Credito_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Tipos_Notas_Credito", "Tipo_Nota_Credito_ID"), "00000")
            .rdoColumns("Tipo_Nota_Credito_ID") = Txt_Tipo_Nota_Credito_ID.Text
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Descripcion_Tipos_Notas_Credito.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Tipos_Notas_Credito.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Tipos_Notas_Credito.Close
    'Pone el encabezado al grid de Tipo Producto
    If Grid_Tipos_Notas_Credito.Rows = 0 Then
        Grid_Tipos_Notas_Credito.AddItem "Tipo Nota ID" & Chr(9) & "Descripcion"
    End If
    'Agrega los datos
    Grid_Tipos_Notas_Credito.AddItem Txt_Tipo_Nota_Credito_ID.Text & Chr(9) & UCase(Trim(Txt_Descripcion_Tipos_Notas_Credito.Text))
    Grid_Tipos_Notas_Credito.FixedRows = 1
    'Da formato al grid
    Fra_Generales_Tipos_Notas_Credito.Enabled = False
    Fra_Tipos_Notas_Credito.Enabled = True
    Grid_Tipos_Notas_Credito.ColWidth(0) = 1500
    Grid_Tipos_Notas_Credito.ColWidth(1) = 4990
    Grid_Tipos_Notas_Credito.ColWidth(2) = 0
    Grid_Tipos_Notas_Credito.ColWidth(3) = 0
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    MsgBox "El tipo de nota de credito ha sido dado de alta", vbInformation
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Tiempos_Muertos
    'DESCRIPCIÓN:           Modifica Tiempos Muertos
    'PARÁMETROS:
    'CREO:                  Julio cruz
    'FECHA_CREO:            10-Dic-2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Modifica_Tiempos_Muertos()
Dim Rs_Modifica_Cat_Tiempos_Muertos As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta la Tiempo_Muerto actual seleccionado
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Tiempos_Muertos"
    Mi_SQL = Mi_SQL & "  WHERE Tiempo_ID ='" & Txt_Tiempo_ID.Text & "'"
    Set Rs_Modifica_Cat_Tiempos_Muertos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Tiempos_Muertos
    With Rs_Modifica_Cat_Tiempos_Muertos
        .Edit
            .rdoColumns("Tiempo_ID") = Txt_Tiempo_ID.Text
            .rdoColumns("Descripcion") = Trim(UCase(Txt_Descripcion_Tiempos.Text))
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Tiempos.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modifica_Cat_Tiempos_Muertos = Nothing
    
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Tiempos_Muertos.Enabled = False
    Fra_Grid_Tiempos_Muertos.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    
    'Configura el grid
    Grid_Tiempos_Muertos.TextMatrix(Grid_Tiempos_Muertos.RowSel, 0) = Txt_Tiempo_ID.Text
    Grid_Tiempos_Muertos.TextMatrix(Grid_Tiempos_Muertos.RowSel, 1) = Trim(UCase(Txt_Descripcion_Tiempos.Text))
    Grid_Tiempos_Muertos.TextMatrix(Grid_Tiempos_Muertos.RowSel, 2) = Trim(UCase(Txt_Comentarios_Tiempos.Text))
    Conexion_Base.CommitTrans
    'Hace la consulta de los tiempos Muertos
'    Consulta_Curso ("")
    MsgBox "El Tiempo Muerto ha sido modificado", vbInformation
    Exit Sub

'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCION: Modifica_Tipos_Notas_Credito
    'DESCRIPCION: Actualiza los datos de las notas de crédito
    'PARAMETROS:
    'CREO: Sergio Ulises Durán Hernández
    'FECHA_CREO: 02-Junio-2009
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACION:
'*******************************************************************************
Public Sub Modifica_Tipos_Notas_Credito()
Dim Rs_Modifica_Cat_Tipos_Notas_Credito As rdoResultset    'Manejo de registro

On Error GoTo HANDLER
    'Consulta la Tiempo_Muerto actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Tipos_Notas_Credito"
    Mi_SQL = Mi_SQL & "  WHERE Tipo_Nota_Credito_ID='" & Txt_Tipo_Nota_Credito_ID.Text & "'"
    Set Rs_Modifica_Cat_Tipos_Notas_Credito = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Cat_Tipos_Notas_Credito.EOF Then
        'Modifica los datos de la tabla Cat_Tiempos_Muertos
        With Rs_Modifica_Cat_Tipos_Notas_Credito
            .Edit
                .rdoColumns("Tipo_Nota_Credito_ID") = Txt_Tipo_Nota_Credito_ID.Text
                .rdoColumns("Descripcion") = Trim(UCase(Txt_Descripcion_Tipos_Notas_Credito.Text))
                .rdoColumns("Comentarios") = Trim(UCase(Txt_Comentarios_Tipos_Notas_Credito.Text))
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
    End If
    Rs_Modifica_Cat_Tipos_Notas_Credito.Close
    'Configura el grid
    Grid_Tipos_Notas_Credito.TextMatrix(Grid_Tipos_Notas_Credito.RowSel, 0) = Txt_Tipo_Nota_Credito_ID.Text
    Grid_Tipos_Notas_Credito.TextMatrix(Grid_Tipos_Notas_Credito.RowSel, 1) = Trim(UCase(Txt_Descripcion_Tipos_Notas_Credito.Text))
    Grid_Tipos_Notas_Credito.TextMatrix(Grid_Tipos_Notas_Credito.RowSel, 2) = Trim(UCase(Txt_Comentarios_Tipos_Notas_Credito.Text))
    MsgBox "El tipo de nota de crédito ha sido modificado", vbInformation
    'Deshabilita botones y habilita los necesarios
    Fra_Generales_Tipos_Notas_Credito.Enabled = False
    Fra_Tipos_Notas_Credito.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Btn_Eliminar.Enabled = True
    Btn_Buscar.Enabled = True
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Txt_Horas_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Horas_Curso.Text, True)
End Sub

Private Sub Txt_Instructor_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_No_Cuenta_Banco_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_No_Nomina_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_No_Nomina.Text, True)
End Sub

Private Sub Txt_Nombre_Banco_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Nombre_Curso_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Nombre_Gap_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Nombre_Gerencia_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Nombre_Transporte_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Nombre_Usuario_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Agrega_Submenus_RH()
Dim Rs_Consulta_Apl_Cat_Accesos_RH As rdoResultset

'Consulta los permisos de los submenus de RH
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    'If UCase(Mid(Ctl.Name, 1, 4)) = UCase("MENU") Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Importacion Asistencias") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Importacion Asistencias") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing

Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Validacion Tiempo Trabajo") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Validacion Tiempo Trabajo") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Solicitud de Permisos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Solicitud de Permisos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Mantenimiento Asistencias") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Mantenimiento Asistencias") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
        
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
        
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Correo_Validacion'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Correo de Validacion") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Correo_Validacion" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Correo de Validacion") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Correo_Validacion" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Asistencia de Empleados") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Asistencia de Empleados") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Compaq'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Exportación a Compaq") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Compaq" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Exportación a Compaq") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Compaq" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Visor_Registros'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Visor de Registros") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Visor_Registros" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Visor de Registros") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Visor_Registros" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
    '**********Catalogos
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Empresas'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Empresas") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empresas" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Empresas") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empresas" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Incidencias Extraordinarias") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
        
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Departamentos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Departamentos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Motivos Baja") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Motivos Baja") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Puestos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Puestos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Puestos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Puestos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Puestos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Equipo de Identificacion") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Equipo de Identificacion") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Nivel de Estudios") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Nivel de Estudios") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Dias No Laborales") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Dias No Laborales") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Turnos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Turnos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Turnos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Turnos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Turnos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Empleados'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Empleados") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empleados" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Empleados") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Empleados" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Cat_Parametros'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Paramentros") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Parametros" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Paramentros") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Cat_Parametros" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
    '********Reportes
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Asistencias") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Asistencias") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Historico Faltas y Retardos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Historico Faltas y Retardos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Historico de Permisos") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Historico de Permisos") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Horas Trabajadas Empleado") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Horas Trabajadas Empleado") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Empleados No Validados") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Empleados No Validados") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Empleados de Baja") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Empleados de Baja") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing
    
Mi_SQL = "SELECT Rol_ID, Menu_Habilitado,Nombre_Sistema"
Mi_SQL = Mi_SQL & " Tipo, ISNULL(Habilitar,'N') as Habilitar, "
Mi_SQL = Mi_SQL & " ISNULL(Alta,'N') as Alta, "
Mi_SQL = Mi_SQL & " ISNULL(Cambio,'N') as Cambio, "
Mi_SQL = Mi_SQL & " ISNULL(Eliminar,'N') as Eliminar,"
Mi_SQL = Mi_SQL & " ISNULL(Consultar,'N') as Consultar"
Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
Mi_SQL = Mi_SQL & " WHERE Rol_ID ='" & Trim(Txt_Rol_ID.Text) & "'"
Mi_SQL = Mi_SQL & " AND Nombre_Sistema = 'SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta'"
Set Rs_Consulta_Apl_Cat_Accesos_RH = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
If Not Rs_Consulta_Apl_Cat_Accesos_RH.EOF Then
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
        "" & Chr(9) & UCase("Empleados de Alta") & _
        Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta" & Chr(9) & _
        "SubMenu" & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Habilitar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Alta") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Cambio") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Eliminar") & _
        Chr(9) & Rs_Consulta_Apl_Cat_Accesos_RH.rdoColumns("Consultar")
Else
    Grid_Accesos_Seguridad.AddItem "" & Chr(9) & _
    "" & Chr(9) & UCase("Empleados de Alta") & _
    Chr(9) & "SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta" & Chr(9) & "SubMenu" & _
    Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & "N"
End If
Set Rs_Consulta_Apl_Cat_Accesos_RH = Nothing

End Sub



Private Sub Txt_Nombre_Zona_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Seccion_Clave_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Sucursal_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

