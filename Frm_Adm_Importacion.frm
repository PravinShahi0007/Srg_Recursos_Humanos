VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Adm_Importacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   135
      Left            =   90
      OleObjectBlob   =   "Frm_Adm_Importacion.frx":0000
      TabIndex        =   47
      Top             =   90
      Width           =   210
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   690
      Left            =   2096
      Picture         =   "Frm_Adm_Importacion.frx":0024
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "A"
      Top             =   7915
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   690
      Left            =   7980
      Picture         =   "Frm_Adm_Importacion.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7915
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Guardar 
      Caption         =   "Guardar"
      Height          =   690
      Left            =   135
      Picture         =   "Frm_Adm_Importacion.frx":0FB0
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "A"
      Top             =   7915
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Limpiar 
      Caption         =   "Limpiar"
      Height          =   690
      Left            =   6018
      Picture         =   "Frm_Adm_Importacion.frx":153A
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "A"
      Top             =   7915
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Exportar 
      Caption         =   "Exportar"
      Height          =   690
      Left            =   4057
      Picture         =   "Frm_Adm_Importacion.frx":1AC4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar Prg_Guardar 
      Height          =   690
      Left            =   1350
      TabIndex        =   5
      Top             =   7915
      Visible         =   0   'False
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   1217
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Prbar_Exportacion 
      Height          =   690
      Left            =   5265
      TabIndex        =   6
      Top             =   7915
      Visible         =   0   'False
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   1217
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Cmd_Exportar 
      Left            =   3285
      Top             =   7915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_Importacion_Keri_System 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   0
      ScaleHeight     =   7800
      ScaleWidth      =   9240
      TabIndex        =   7
      Top             =   0
      Width           =   9240
      Begin VB.Frame Fra_Importacion_Keri_Automatico 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Automatica (Conexión IP)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   90
         TabIndex        =   22
         Top             =   420
         Width           =   9105
         Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Automatico 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   240
            Width           =   5700
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   690
            Left            =   7695
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   1217
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   8520
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Inicio_Automatico 
            Height          =   315
            Left            =   1350
            TabIndex        =   32
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   125894659
            CurrentDate     =   39931
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Termino_Automatico 
            Height          =   315
            Left            =   5190
            TabIndex        =   33
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   125894659
            CurrentDate     =   39931
         End
         Begin VB.CommandButton Btn_Importacion_Automatica 
            Caption         =   "Importar"
            Height          =   690
            Left            =   7785
            Picture         =   "Frm_Adm_Importacion.frx":204E
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label7 
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
            Left            =   180
            TabIndex        =   37
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Termino"
            Height          =   195
            Left            =   3645
            TabIndex        =   34
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   660
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Importacion_Keri 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manual por Archivo"
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
         Left            =   90
         TabIndex        =   8
         Top             =   1500
         Width           =   9105
         Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Manual 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   5700
         End
         Begin VB.ComboBox Cmb_Adm_Importacion_Checador 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   5700
         End
         Begin VB.CommandButton Btn_Importacion 
            Caption         =   "Importar"
            Height          =   690
            Left            =   7785
            Picture         =   "Frm_Adm_Importacion.frx":25D8
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton Btn_Ruta_Checador 
            Caption         =   "..."
            Height          =   315
            Left            =   6570
            TabIndex        =   10
            Top             =   975
            Width           =   450
         End
         Begin VB.TextBox Txt_Adm_Importacion_Ruta_Archivo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   9
            Top             =   960
            Width           =   5205
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Inicio 
            Height          =   315
            Left            =   1350
            TabIndex        =   11
            Top             =   1365
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   125894659
            CurrentDate     =   39931
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Termino 
            Height          =   315
            Left            =   5160
            TabIndex        =   12
            Top             =   1365
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   125894659
            CurrentDate     =   39931
         End
         Begin MSComctlLib.ProgressBar PrgBar_Importacion 
            Height          =   690
            Left            =   7695
            TabIndex        =   13
            Top             =   300
            Visible         =   0   'False
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   1217
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComDlg.CommonDialog cmd_Archivo_Registros 
            Left            =   7155
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label9 
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
            Left            =   180
            TabIndex        =   39
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Checador"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   660
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Termino"
            Height          =   195
            Left            =   3645
            TabIndex        =   17
            Top             =   1425
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1425
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   1035
            Width           =   540
         End
      End
      Begin TabDlg.SSTab Tab_Importacion_Checadas 
         Height          =   4515
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   7964
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Generar Asistencia"
         TabPicture(0)   =   "Frm_Adm_Importacion.frx":2B62
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Informacion_Importación_Kery_Systema"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Log"
         TabPicture(1)   =   "Frm_Adm_Importacion.frx":2B7E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   3885
            Left            =   -74955
            TabIndex        =   30
            Top             =   360
            Width           =   9015
            Begin VB.TextBox Txt_Importacion_Keri_Log 
               Height          =   3615
               Left            =   90
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   180
               Width           =   8835
            End
         End
         Begin VB.Frame Fra_Informacion_Importación_Kery_Systema 
            BackColor       =   &H00FFFFFF&
            Height          =   4005
            Left            =   60
            TabIndex        =   27
            Top             =   480
            Width           =   9015
            Begin MSFlexGridLib.MSFlexGrid Grid_Importacion_Lista_Depurada 
               Height          =   2535
               Left            =   0
               TabIndex        =   28
               Top             =   1320
               Width           =   8880
               _ExtentX        =   15663
               _ExtentY        =   4471
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.ComboBox Cmb_Empleado 
               Height          =   315
               ItemData        =   "Frm_Adm_Importacion.frx":2B9A
               Left            =   1275
               List            =   "Frm_Adm_Importacion.frx":2B9C
               TabIndex        =   49
               Top             =   540
               Width           =   7620
            End
            Begin MSComctlLib.ProgressBar Prbar_Asistencia 
               Height          =   105
               Left            =   6960
               TabIndex        =   35
               Top             =   1200
               Visible         =   0   'False
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   185
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Asistencia 
               Height          =   315
               Left            =   1275
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   180
               Width           =   7620
            End
            Begin VB.CommandButton Btn_Generar_Asistencias 
               Caption         =   "Generar Asistencia"
               Height          =   270
               Left            =   6900
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   900
               Width           =   2040
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Importacion 
               Height          =   600
               Left            =   45
               TabIndex        =   29
               Top             =   1425
               Visible         =   0   'False
               Width           =   7350
               _ExtentX        =   12965
               _ExtentY        =   1058
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin MSComCtl2.DTPicker Dtp_Asistencia_Fecha_Inicio 
               Height          =   315
               Left            =   1275
               TabIndex        =   41
               Top             =   960
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "ddd dd MMM yyyy"
               Format          =   125894659
               CurrentDate     =   39931
            End
            Begin MSComCtl2.DTPicker Dtp_Asistencia_Fecha_Termino 
               Height          =   315
               Left            =   4755
               TabIndex        =   42
               Top             =   960
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "ddd dd MMM yyyy"
               Format          =   125894659
               CurrentDate     =   39931
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Archivo"
               Height          =   135
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Lbl_Empleado 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Empleado"
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
               TabIndex        =   48
               Top             =   600
               Width           =   840
            End
            Begin VB.Label Label12 
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
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha Inicio"
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   960
               Width           =   870
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha Termino"
               Height          =   195
               Left            =   3405
               TabIndex        =   43
               Top             =   960
               Width           =   1065
            End
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTACION ASISTENCIAS"
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
         Left            =   1935
         TabIndex        =   18
         Top             =   0
         Width           =   5355
      End
   End
   Begin VB.Label Lbl_Progreso_Exportacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Archivo"
      Height          =   135
      Left            =   5445
      TabIndex        =   19
      Top             =   6945
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Frm_Adm_Importacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As String                     'Define la opcion para los procesos

Private Sub Btn_Exportar_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String
On Error GoTo HANDLER
    If Grid_Importacion_Lista_Depurada.Rows > 1 Then
        Cmd_Exportar.CancelError = True
        Cmd_Exportar.DialogTitle = "Seleccione el directorio"
        Cmd_Exportar.Flags = cdlOFNHideReadOnly
        Cmd_Exportar.Filter = "Archivos de Excel(*.xls)|*.xls"
        Cmd_Exportar.FilterIndex = 2
        Cmd_Exportar.FileName = Opcion & ".xls"
        Cmd_Exportar.ShowSave
        Ruta_Exportacion = Cmd_Exportar.FileName
        Nombre_Archivo = Cmd_Exportar.FileTitle
        If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
'            Call Exportar_Excel(Ruta_Temporal & Opcion & "xls.txt", Ruta_Exportacion, Prbar_Exportacion, Lbl_Progreso_Exportacion, Me)
             Call Exportar_Excel_Bien(Ruta_Temporal & Opcion & "xls.txt", Ruta_Exportacion)
        End If
    Else
        MsgBox "No existe información para exportar", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Generar_Asistencias_Click()
    If Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex > -1 Then
'        If (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SABADO" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "SABADO") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SÁBADO" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "SÁBADO") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SATURDAY" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "SATURDAY") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "DOMINGO" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "DOMINGO") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SUNDAY" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "SUNDAY") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SABADO" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "DOMINGO") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SÁBADO" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "DOMINGO") _
'        Or (UCase(Format(Dtp_Asistencia_Fecha_Inicio.Value, "dddd")) = "SATURDAY" And UCase(Format(Dtp_Asistencia_Fecha_Termino.Value, "dddd")) = "SUNDAY") Then
'            'Valida que seleccione sólo sábado o sólo domingo
'            If Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") = Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") Then
'                Depurar_Lista_Fin_Semana
'            Else
'                MsgBox "Para las asistencias de fin de semana, deberá cargarlos por día, es decir, sólo sábado o sólo domingo", vbExclamation
'                Dtp_Asistencia_Fecha_Termino.SetFocus
'            End If
'        Else
            Depurar_Lista
            Depurar_Lista_Turnos_Flexibles
            If Grid_Importacion_Lista_Depurada.Rows <= 0 Then
                MsgBox "No hay datos para mostrar", vbInformation
            End If
'        End If
    Else
        MsgBox "Seleccione la empresa a la que se genera la asistencia", vbExclamation
        Cmb_Adm_Importacion_Empresa_Asistencia.SetFocus
    End If
End Sub

Private Sub Btn_Importacion_Automatica_Click()
Dim Ruta_Archivo As String
Dim Mensaje As String
Dim Checador_ID As String
    If Cmb_Adm_Importacion_Empresa_Automatico.ListIndex > -1 Then
        'Valida la fecha inicial y final
        If Format(Dtp_Importacion_Fecha_Inicio_Automatico.Value, "yyyyMMdd") > Format(Dtp_Importacion_Fecha_Termino_Automatico.Value, "yyyyMMdd") Then
            MsgBox "Revise las fechas, la fecha inicial debe ser menor o igual que la final", vbExclamation
            Dtp_Importacion_Fecha_Inicio_Automatico.SetFocus
            Exit Sub
        End If
        Ruta_Archivo = ""
        Mensaje = ""
        Tab_Importacion_Checadas.Tab = 1
        Call Obtiene_Informacion_Checadas(Ruta_Archivo, Mensaje)
    End If
End Sub

Private Sub Btn_Imprimir_Click()
Dim linea As String 'Obtiene el texto a imprimir
Dim Impresora As String            'Tomna el nombre la impresora
Dim Mi_Impresora As Printer        'Toma el nombre de la impresora
Dim Ubicacion_Impresora As String  'Toma el valor de la ubicacion dela impresora
Dim Encabezado As Boolean          'Identifica si se ha encontrado el encabezado
Dim Lineas_Impresas As Integer     'Contador de lineas impresas
Dim Primer_Paso As Boolean         'Indica si es el inicio del archivop
Dim Escribir_Archivo As Boolean    'Indica si se escribe en el archivo temp
Dim Imprime_No_Pagina As Boolean   'Se imprime el no de pagina
Dim Control_Encabezado As Integer  'Control del encabezado
Dim Lineas_Encabezado As Integer

On Error GoTo HANDLER
    If Grid_Importacion_Lista_Depurada.Rows < 1 Then
        MsgBox "No existe información para imprimir", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    Primer_Paso = True
    Escribir_Archivo = True
    Lineas_Impresas = 0
    Imprime_No_Pagina = True
    Lineas_Encabezado = 9
On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    Debug.Print Printer.DeviceName
    Debug.Print Printer.Page
    Printer.FontSize = 9
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 10
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 9
    Printer.Font = "Courier New"
    Open Ruta_Temporal & Opcion & ".txt" For Input As #1
    Do While Not EOF(1)
        'Localiza el inicio del archivo y la primera fecha encontrada delimitando el encabezado
        'Lee la primera linea
        Line Input #1, linea
        If linea = "NP" Then
            GoTo 6000
        End If
        If Imprime_No_Pagina = True Then
            If Len(linea) > 0 Then
                Printer.Print linea & "                         Pag #" & Conectar_Ayudante.Alinea_Derecha(Printer.Page, 3)
                Imprime_No_Pagina = False
            Else
                Printer.Print linea
            End If
        Else
            Printer.Print linea
        End If
        'Se crea el archivo de encabezado para los reportes
        'Printer.Print Linea
        'Imprime el numero de pagina
        If Primer_Paso = True Then
            Open Ruta_Temporal & Opcion & "_tmp.txt" For Output As #2
            Primer_Paso = False
        End If
        Control_Encabezado = Control_Encabezado + 1
        If Escribir_Archivo = True And Lineas_Encabezado >= Control_Encabezado Then Print #2, linea
        Debug.Print linea
        If Lineas_Encabezado < Control_Encabezado _
          And Escribir_Archivo = True Then
            'Print #2, " ";
            Close #2
            Escribir_Archivo = False
        End If
        'se incrementa el contador de lineas
        Lineas_Impresas = Lineas_Impresas + 1
        If Lineas_Impresas >= 57 Then
6000:       Imprime_No_Pagina = True
            If Lineas_Impresas < 57 Then
                Lineas_Impresas = Lineas_Impresas + 2
                For Lineas_Impresas = Lineas_Impresas To 60
                    Printer.Print
                Next Lineas_Impresas
            End If
            Printer.FontSize = 9
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 10
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 9
            Printer.Font = "Courier New"
            'Imprime el encabezado
            Open Ruta_Temporal & Opcion & "_tmp.txt" For Input As #2
            Do While Not EOF(2)
                Line Input #2, linea
                If Imprime_No_Pagina = True Then
                    If Len(linea) > 0 Then
                        Printer.Print linea & " Pag #" & Conectar_Ayudante.Alinea_Derecha(Printer.Page, 3)
                        Imprime_No_Pagina = False
                    Else
                        Printer.Print linea
                    End If
                Else
                    Printer.Print linea
                End If
                Debug.Print linea
            Loop
            Lineas_Impresas = Lineas_Encabezado
            Close #2
        End If
    Loop
    Printer.EndDoc
    Close #1
    MsgBox "Reporte enviado a impresora", vbInformation + vbOKOnly, Me.Caption
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Er In rdoErrors
        If Mid(Er, 1, 5) = "01S02" Then MsgBox "No se encontro la impresora", vbCritical + vbOKOnly, Me.Caption
    Next Er
    Debug.Print Err.Description
End Sub

Private Sub Btn_Ruta_Checador_Click()
    
On Error GoTo HANDLER
  cmd_Archivo_Registros.CancelError = True
  
  cmd_Archivo_Registros.Flags = cdlOFNHideReadOnly
  cmd_Archivo_Registros.Filter = "Archivo de Asistencias (*.dat)|*.dat"
  cmd_Archivo_Registros.FilterIndex = 2
  cmd_Archivo_Registros.ShowOpen
  Txt_Adm_Importacion_Ruta_Archivo = cmd_Archivo_Registros.FileName
  Exit Sub
  
HANDLER:
  Exit Sub

End Sub

Private Sub Btn_Guardar_Click()
    Select Case Opcion
        Case "Importacion_Asistencias":
            Dim Respuesta As Integer        'Respuesta a la pregunta de almacenar datos
            Dim No_Operacion As String      'No de operacion que se encontro para modificar
            Dim Cont_Fila As Integer        'Recorre la lista para ver si hay trabajadores no registrados
            
                If Grid_Importacion_Lista_Depurada.Rows > 1 Then
                    'Valida que los empleados esten registrados
                    For Cont_Fila = 1 To Grid_Importacion_Lista_Depurada.Rows - 1
                        If Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 9) = "N" Then
                            If MsgBox("Hay empleados no registrados en la Interfaz," + vbCrLf + _
                                   "se guardara solamente la información exisente en el sistema, ¿Continuar?", vbInformation + vbYesNo, Me.Caption) = vbYes Then
                                Exit For
                            Else
                                Exit Sub
                            End If
                        End If
                    Next
                    'Valida que no se ha ingresado la información previamente
                    Mi_SQL = "SELECT COUNT(*) AS Detalles FROM Adm_Asistencias_Detalles"
                    Mi_SQL = Mi_SQL & " WHERE Fecha BETWEEN '" & Format(Dtp_Importacion_Fecha_Inicio_Automatico, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Importacion_Fecha_Termino_Automatico, "MM/dd/yyyy") & "'"
                    If Val(Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Detalles")) > 0 Then
                        If MsgBox("Ya existe información en el rango de fechas seleccionado" + vbCrLf + _
                            "¿Desea Sobreescribirla?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                            Alta_Importacion_Asistencias
                        End If
                    Else
                        Alta_Importacion_Asistencias
                    End If
                Else
                    MsgBox "No existe información a guardar", vbInformation + vbOKOnly, Me.Caption
                End If
        Case "Importacion_Archivo":
        
    End Select
End Sub

Private Sub Btn_Importacion_Click()
    'valida el rango de fechas
    Dim Checador_ID As String
    If DateDiff("d", Dtp_Importacion_Fecha_Inicio.Value, Dtp_Importacion_Fecha_Termino.Value) >= 0 Then
        If Cmb_Adm_Importacion_Empresa_Manual.ListIndex > -1 Then
            If Cmb_Adm_Importacion_Checador.ListIndex > -1 Then
                Tab_Importacion_Checadas.Tab = 1
                Call Obtiene_Informacion_Checadas_Archivo(Trim(Txt_Adm_Importacion_Ruta_Archivo.Text), _
                    Format(Cmb_Adm_Importacion_Checador.ItemData(Cmb_Adm_Importacion_Checador.ListIndex), "00000"), _
                    Format(Cmb_Adm_Importacion_Empresa_Manual.ItemData(Cmb_Adm_Importacion_Empresa_Manual.ListIndex), "00000"))

            Else
                MsgBox "Seleccione un checador para asignar las checadas"
            End If
        End If
    Else
        MsgBox "Rango de fechas no valido", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
End Sub

Private Sub Btn_Limpiar_Click()
    Select Case Opcion
        Case "Importacion_Asistencias":
            Grid_Importacion.Rows = 0
            Grid_Importacion_Lista_Depurada.Rows = 0
            Dtp_Importacion_Fecha_Inicio.Value = Now
            Dtp_Importacion_Fecha_Termino.Value = Now
            Txt_Adm_Importacion_Ruta_Archivo.Text = ""
            Txt_Importacion_Keri_Log.Text = ""
    End Select
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN:  Generar_Lista
'DESCRIPCIÓN:           Genera la lista de informacion del systema
'PARÁMETROS :           Nomre_Archivo: Ruta del archivo a leer
'                       Checador_ID: Identificador del Checador
'                       Origen: AUTOMATICO/MANUAL
'CREO       :           Yañez Rodriguez Diego Neftali
'FECHA_CREO :           19 Mayo 2009
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Generar_Lista(Nombre_Archivo As String, Checador_ID As String, Origen As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Ejecuta_Export As String
Dim Ruta_Archivo As String
Dim Archivo_Temporal As String
Dim linea As String
Dim Datos() As String
Dim I As Integer
Dim num As Double
Dim Contador As Long
Dim hProcess As Long
Dim Primero_Local As Boolean
Dim Cadena As String
Dim Nombre_Empleado As String
Dim Empleado_ID_Consulta As String
Dim nomruta As String

On Error GoTo HANDLER:
    Me.MousePointer = 11
    Nombre_Empleado = ""
    Empleado_ID_Consulta = ""
    Grid_Importacion.Rows = 0
    Grid_Importacion.Cols = 9
    Grid_Importacion_Lista_Depurada.Rows = 0
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Generando Lista de empleados para la fecha seleccionada"
    If Nombre_Archivo = "" Then
        MsgBox "Seleccione un archivo para importar", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe archivo para importar información"
        Me.MousePointer = 0
        Exit Sub
    End If
    Ruta_Archivo = Nombre_Archivo
    'valida que el archivo de exportacion exista en la ruta proporcionada
    If Len(Dir$(Ruta_Archivo)) <= 0 Then
        MsgBox "El archivo no contiene información o no existe, favor de verificar", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "El archivo no contiene información"
        Me.MousePointer = 0
        PrgBar_Importacion.Visible = False
        Exit Sub
    End If
    nomruta = Ruta_Temporal & Checador_ID & "_" & Format(Now, "MMddyyyy_HHmmss") & ".dat"
    'Forma la cadena No_Empleado(ya), Checador_ID(ya), Modo de Verificacion(ya), Modo_ES(ya), Fecha MM/dd/yyyy(ya) hh:mm:ss(ya)
    Grid_Importacion.AddItem "No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S"
    Open Ruta_Archivo For Input As #1   'Abre el archivo del sistema de checadas
    Open nomruta For Output As #2   'Abre el archivo temporal
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
        "Incicio de depuracion de información para generar la lista"
    Do While Not EOF(1)
        Line Input #1, linea
        Datos() = Split(linea, Chr(9))
        Debug.Print linea
        If Datos(0) = "Fin_Archivo" Then
            Exit Do
        End If
        If Origen = "AUTOMATICO" Then
            If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio_Automatico.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
                DateDiff("d", Format(Dtp_Importacion_Fecha_Termino_Automatico.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
                Print #2, linea
                Contador = Contador + 1
                Me.Refresh
            End If
        End If
        If Origen = "MANUAL" Then
            If Mid(Datos(0), 1, 3) <> "SN=" And Mid(Datos(0), 1, 8) <> "CHECKSUM" Then
                If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
                    DateDiff("d", Format(Dtp_Importacion_Fecha_Termino.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
                    Print #2, linea & Chr(9) & Checador_ID
                    Contador = Contador + 1
                    Me.Refresh
                End If
            End If
        End If
    Loop
    Close #1, #2
    PrgBar_Importacion.Visible = True
    PrgBar_Importacion.Value = 0
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
        "No de registros localizados: " & Contador
    If Contador > 0 Then
        PrgBar_Importacion.Max = Contador
    Else
        PrgBar_Importacion.Visible = False
        MsgBox "No existe información para la fecha seleccionada", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe información para la fecha seleccionada"
        Me.MousePointer = 0
        Exit Sub
    End If
    If Len(Dir$(nomruta)) <= 0 Then
        MsgBox "No existe información en el rango de fechas seleccionado", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe información en el rango de fechas seleccionado"
        Me.MousePointer = 0
        Exit Sub
    End If
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Generando Lista..."
    Open nomruta For Input As #1
    Do While Not EOF(1)
        Me.MousePointer = 11
        Line Input #1, linea
        Cadena = ""
        num = 0
'        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Generando Lista..."
        Datos() = Split(linea, Chr(9))
        If Origen = "AUTOMATICO" Then
            If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio_Automatico.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
                DateDiff("d", Format(Dtp_Importacion_Fecha_Termino_Automatico.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
                'Grid_Importacion.AddItem "No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S"
                For I = 0 To UBound(Datos())
                    If I = 0 Then
                        Cadena = Trim(Datos(I))
                        num = Val(Trim(Datos(I)))
                        'cadena = cadena & Chr(9) & Val(Trim(Datos(I)))
                        If num <> 0 Then
                            Mi_SQL = "SELECT Empleado_ID, No_Tarjeta, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre FROM Cat_Empleados WHERE No_Tarjeta = " & Val(num) & ""
                            Nombre_Empleado = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Nombre")
                            Empleado_ID_Consulta = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Empleado_ID")
                            If Nombre_Empleado <> "" Then
                                Cadena = Cadena & Chr(9) & Nombre_Empleado & Chr(9) & Empleado_ID_Consulta & Chr(9) & "S"
                            Else
                                Cadena = Cadena & Chr(9) & "Gafete Desconocido" & Chr(9) & "Desconocido" & Chr(9) & "N"
                            End If
                        End If
                    Else
                        If I = 1 Then
                            Cadena = Cadena & Chr(9) & Format(Trim(Datos(I)), "MM/dd/yyyy")
                            Cadena = Cadena & Chr(9) & Format(Trim(Datos(I)), "HH:mm:ss")
                            'Se agrega la informacion restante
                            '"No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S
                            Cadena = Cadena & Chr(9) & Datos(6) & Chr(9) & Datos(2) & Chr(9) & Datos(3)
                            Exit For
                        End If
                    End If
                Next I
                Debug.Print Cadena
                If num <> 0 Then Grid_Importacion.AddItem Cadena
                PrgBar_Importacion.Value = PrgBar_Importacion.Value + 1
                Me.Refresh
            End If
        End If
        If Origen = "MANUAL" Then
            If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
                DateDiff("d", Format(Dtp_Importacion_Fecha_Termino.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
                For I = 0 To UBound(Datos())
                    If I = 0 Then
                        Cadena = Trim(Datos(I))
                        num = Val(Trim(Datos(I)))
                        'cadena = cadena & Chr(9) & Val(Trim(Datos(I)))
                        If num <> 0 Then
                            Mi_SQL = "SELECT Empleado_ID, No_Tarjeta, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre FROM Cat_Empleados WHERE No_Tarjeta = " & Val(num) & ""
                            Nombre_Empleado = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Nombre")
                            Empleado_ID_Consulta = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Empleado_ID")
                            If Nombre_Empleado <> "" Then
                                Cadena = Cadena & Chr(9) & Nombre_Empleado & Chr(9) & Empleado_ID_Consulta & Chr(9) & "S"
                            Else
                                Cadena = Cadena & Chr(9) & "Gafete Desconocido" & Chr(9) & "Desconocido" & Chr(9) & "N"
                            End If
                        End If
                    Else
                        If I = 1 Then
                            Cadena = Cadena & Chr(9) & Format(Trim(Datos(I)), "MM/dd/yyyy")
                            Cadena = Cadena & Chr(9) & Format(Trim(Datos(I)), "HH:mm:ss")
                            'Se agrega la informacion restante
                            '"No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S
                            Cadena = Cadena & Chr(9) & Datos(6) & Chr(9) & Datos(2) & Chr(9) & Datos(3)
                            Exit For
                        End If
                    End If
                Next I
                Debug.Print Cadena
                If num <> 0 Then Grid_Importacion.AddItem Cadena
                PrgBar_Importacion.Value = PrgBar_Importacion.Value + 1
                Me.Refresh
            End If
        End If
    Loop
    Close #1
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Generación de Lista Terminado..."
    If Grid_Importacion.Rows > 1 Then
'        "No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S"
        Grid_Importacion.FixedRows = 1
        Grid_Importacion.ColWidth(0) = 1100         'No Tarjeta
        Grid_Importacion.ColAlignment(0) = 3
        Grid_Importacion.ColWidth(1) = 1100         'Empleado
        Grid_Importacion.ColAlignment(1) = 3
        Grid_Importacion.ColWidth(2) = 1100         'Empleado_ID
        Grid_Importacion.ColAlignment(2) = 3
        Grid_Importacion.ColWidth(3) = 3400         'Registrado
        Grid_Importacion.ColWidth(4) = 400          'Fecha
        Grid_Importacion.ColWidth(5) = 400          'Hora
        Grid_Importacion.ColWidth(6) = 0           'Checador_ID
        Grid_Importacion.ColWidth(7) = 0           'Modo_Verficacion
        Grid_Importacion.ColWidth(8) = 0           'E/S
    End If
    PrgBar_Importacion.Visible = False
Exit Sub
HANDLER:
    Me.MousePointer = 0
    Close #1
    Grid_Importacion.Rows = 0
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Alta_Importacion_Asistencias
'DESCRIPCION: Genera la lista de informacion del systema Keri-System
'PARAMETROS :
'CREO       : Yañez Rodriguez Diego Neftali
'FECHA_CREO : 19-Mayo-2009
'MODIFICO   : Sergio Ulises Durán Hernández
'FECHA_MODIFICO: 06-Marzo-2012
'CAUSA_MODIFICO: Se cambiaron por UPDATES e INSERTS la carga de registros para asistencias
'*******************************************************************************
Private Sub Alta_Importacion_Asistencias()
Dim Rs_Alta_Adm_Asistencias_Detalles As rdoResultset     'Información de las asistencias
Dim Rs_Modifica_Adm_Asistencias_Detalles As rdoResultset     'Información de las asistencias
Dim Rs_Consulta_Informacion_Turnos As rdoResultset              'Informacion de los turnos
Dim Operacion As String                                         'Consecutivo del maximo del catalogo
Dim Cont_Fila As Integer                                        'Recorre el grid
Dim Turno_Empleado As String                                    'Guarda el Turno del empleado
Dim Hora_Inicio_Turno As Date                                   'Guarda la hora de inicio del turno
Dim Hora_Termino_Turno As Date                                   'Guarda la hora de inicio del turno

On Error GoTo HANDLER:
    Me.MousePointer = 11
    If Grid_Importacion_Lista_Depurada.Rows > 0 Then
        Prg_Guardar.Value = 0
        Prg_Guardar.Max = Grid_Importacion_Lista_Depurada.Rows
        Prg_Guardar.Visible = True
    End If
    Conexion_Base.BeginTrans
    For Cont_Fila = 1 To Grid_Importacion_Lista_Depurada.Rows - 1
        If Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 9) = "S" Then
            'Verifica si el registro ya se ha generado para actualizarlo, si no lo da de alta
            Mi_SQL = "SELECT * FROM Adm_Asistencias_Detalles "
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 10)) & "'"
            Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
            Set Rs_Modifica_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Modifica_Adm_Asistencias_Detalles.EOF Then
'                With Rs_Modifica_Adm_Asistencias_Detalles
'                    .Edit
'                        '.rdoColumns("Empleado_ID") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 9))
'                        '.rdoColumns("No_Tarjeta") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 1))
'                        '.rdoColumns("Fecha") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy")
'                        .rdoColumns("Hora_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss")
'                        .rdoColumns("Hora_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 6), "HH:mm:ss")
'                        .rdoColumns("Hora_Comida_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4), "HH:mm:ss")
'                        .rdoColumns("Hora_Comida_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 5), "HH:mm:ss")
'                        .rdoColumns("Horas_Laboradas") = Val(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 7))
'                        '.rdoColumns("Validada") = "N"
'                    .Update
'                End With
                'Cambia sólo si no está validada la asistencia
                'If Rs_Modifica_Adm_Asistencias_Detalles.rdoColumns("Validada") = "N" Then
                    Mi_SQL = "UPDATE Adm_Asistencias_Detalles"
                    Mi_SQL = Mi_SQL & " SET Hora_Entrada='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Salida='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 6), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Comida_Entrada='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Comida_Salida='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 5), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Horas_Laboradas='" & Val(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 7)) & "'"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 10)) & "'"
                    Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
                    Conexion_Base.Execute Mi_SQL
                'End If
            Else
'                Set Rs_Alta_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias_Detalles")
'                With Rs_Alta_Adm_Asistencias_Detalles
'                    .AddNew
'                        .rdoColumns("Empleado_ID") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 10))
'                        .rdoColumns("No_Tarjeta") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 1))
'                        .rdoColumns("Fecha") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy")
'                        .rdoColumns("Hora_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss")
'                        .rdoColumns("Hora_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 6), "HH:mm:ss")
'                        .rdoColumns("Hora_Comida_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4), "HH:mm:ss")
'                        .rdoColumns("Hora_Comida_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 5), "HH:mm:ss")
'                        .rdoColumns("Horas_Laboradas") = Val(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 7))
'                        '.rdoColumns("Checador_ID") = Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 11)
'                        .rdoColumns("Validada") = "N"
'                    .Update
'                End With
                Mi_SQL = "INSERT INTO Adm_Asistencias_Detalles(Empleado_ID,No_Tarjeta,Fecha,Hora_Entrada,Hora_Salida,Hora_Comida_Entrada,Hora_Comida_Salida,Horas_Laboradas,Validada)"
                Mi_SQL = Mi_SQL & " VALUES('" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 10)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 1)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 6), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 5), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Val(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 7)) & "'"
                Mi_SQL = Mi_SQL & " , 'N')"
                Conexion_Base.Execute Mi_SQL
            End If
            Rs_Modifica_Adm_Asistencias_Detalles.Close
        End If
        Prg_Guardar.Value = Prg_Guardar.Value + 1
        Me.Refresh
    Next
    Conexion_Base.CommitTrans
    'Grid_Importacion_Lista_Depurada.Rows = 0
    Me.MousePointer = 0
    Prg_Guardar.Visible = False
    Grid_Importacion.Rows = 0
    'Dtp_Importacion_Fecha_Inicio.Value = Now
    'Dtp_Importacion_Fecha_Termino.Value = Now
    Txt_Adm_Importacion_Ruta_Archivo.Text = ""
    MsgBox "Lista Registrada", vbInformation + vbOKOnly, Me.Caption
'    If Grid_Importacion_Lista_Depurada.Rows > 0 Then
'        MsgBox "Hay información que no se registro en el sistema, favor de revisar", vbInformation + vbOKOnly, Me.Caption
'    End If
Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Me.MousePointer = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modifica_Importacion_Asistencias
    'DESCRIPCIÓN:           Genera la lista de informacion del systema Keri-System
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           19 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Importacion_Asistencias()
'Dim Rs_Modifica_Adm_Asistencias As rdoResultset     'Información de las asistencias
'Dim Rs_Consulta_Informacion_Turnos As rdoResultset              'Informacion de los turnos
'Dim Cont_Fila As Integer                                        'Recorre el grid
'Dim Turno_Empleado As String                                    'Guarda el Turno del empleado
'Dim Hora_Inicio_Turno As Date                                   'Guarda la hora de inicio del turno
'Dim Hora_Termino_Turno As Date                                   'Guarda la hora de inicio del turno
'On Error GoTo HANDLER
'Conexion_Base.BeginTrans
'
''Mi_SQL = "DELETE FROM Adm_Asistencias_Detalles "
''Mi_SQL = Mi_SQL & " WHERE Fecha >= " & Par_Fecha & Format(Dtp_Importacion_Fecha_Inicio, "MM/dd/yyyy") & Par_Fecha
''Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(Dtp_Importacion_Fecha_Termino, "MM/dd/yyyy") & Par_Fecha
'Conexion_Base.Execute Mi_SQL
'Me.MousePointer = 11
'If Grid_Importacion_Lista_Depurada.Rows > 0 Then
'    Prg_Guardar.Value = 0
'    Prg_Guardar.Max = Grid_Importacion_Lista_Depurada.Rows
'    Prg_Guardar.Visible = True
'End If
'For Cont_Fila = 1 To Grid_Importacion_Lista_Depurada.Rows - 1
'    If Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 8) = "S" Then
'        'Actualiza la informacion
'
'        Set Rs_Alta_Adm_Asistencias = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias_Detalles")
'        With Rs_Alta_Adm_Asistencias
'            .AddNew
'                .rdoColumns("Empleado_ID") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 9))
'                .rdoColumns("No_Tarjeta") = Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 1))
'                .rdoColumns("Fecha") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy")
'                .rdoColumns("Hora_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss")
'                .rdoColumns("Hora_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 6), "HH:mm:ss")
'                .rdoColumns("Comida_Entrada") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4), "HH:mm:ss")
'                .rdoColumns("Comida_Salida") = Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 5), "HH:mm:ss")
'                .rdoColumns("Horas_Laboradas") = Val(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 7))
'                .rdoColumns("Validada") = "N"
'            .Update
'            .Close
'        End With
'        Set Rs_Alta_Adm_Asistencias = Nothing
'    End If
'    Prg_Guardar.Value = Prg_Guardar.Value + 1
'    Me.Refresh
'Next
'Conexion_Base.CommitTrans
''Grid_Importacion_Lista_Depurada.Rows = 0
'Me.MousePointer = 0
'Prg_Guardar.Visible = False
'Grid_Importacion.Rows = 0
''Dtp_Importacion_Fecha_Inicio.Value = Now
''Dtp_Importacion_Fecha_Termino.Value = Now
'Txt_Adm_Importacion_Ruta_Archivo.Text = ""
'MsgBox "Lista Registrada", vbInformation + vbOKOnly, Me.Caption
'If Grid_Importacion_Lista_Depurada.Rows > 0 Then
'    MsgBox "Hay información que no se registro en el sistema, favor de revisar", vbInformation + vbOKOnly, Me.Caption
'End If
'
'Exit Sub
''Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
'HANDLER:
'Me.MousePointer = 0
'    Conexion_Base.RollbackTrans
'    For Each Er In rdoErrors
'        MsgBox Er.Description
'    Next Er
'
End Sub

Public Sub Inicializa()
    Select Case Opcion
        Case "Importacion_Asistencias":
            Dtp_Importacion_Fecha_Inicio.Value = DateAdd("d", -1, Now)
            Dtp_Importacion_Fecha_Termino.Value = DateAdd("d", 0, Now)
            Dtp_Importacion_Fecha_Inicio_Automatico.Value = DateAdd("d", -1, Now)
            Dtp_Importacion_Fecha_Termino_Automatico.Value = DateAdd("d", 0, Now)
            Dtp_Asistencia_Fecha_Inicio.Value = DateAdd("d", -1, Now)
            Dtp_Asistencia_Fecha_Termino.Value = DateAdd("d", 0, Now)
            If Cmb_Adm_Importacion_Checador.ListCount > 0 Then
                Cmb_Adm_Importacion_Checador.ListIndex = 0
            End If
            'Carga las empresas localidades
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Importacion_Empresa_Automatico, 0, "Nombre", , True, "TODAS")
            If Cmb_Adm_Importacion_Empresa_Automatico.ListCount > 0 Then
                Cmb_Adm_Importacion_Empresa_Automatico.ListIndex = 0
            End If
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Importacion_Empresa_Manual, 0, "Nombre", , False, "TODAS")
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Importacion_Empresa_Asistencia, 0, "Nombre", , False, "TODAS")
            
    End Select
End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date)
    Open Ruta_Temporal & Opcion & ".txt" For Output As #1
    Open Ruta_Temporal & Opcion & "xls.txt" For Output As #2 'Reporte a xls
    Print #1,
    Print #2,
    Print #1, Conectar_Ayudante.Centrar_Texto(Empresa, 120)
    Print #2, "||"; Empresa
    Print #1,
    Print #2,
    Print #1, Titulo; Conectar_Ayudante.Alinea_Derecha(Format(Now, "dd MMM yyyy"), 110 - Len(Titulo))
    Print #2, "||" & Titulo; "|||||"; Format(Now, "dd MMM yyyy")
    Print #1,
    Print #2,
    If DateDiff("s", Format(Fecha_Inicial, "HH:mm:ss"), "00:00:00") <> 0 And DateDiff("s", Format(Fecha_Termino, "HH:mm:ss"), "00:00:00") <> 0 Then
        Print #1, "DE "; Format(Fecha_Inicial, "dd MMMM yyyy") & " A "; Format(Fecha_Termino, "dd MMMM yyyy")
        Print #2, "|DE|"; Format(Fecha_Inicial, "dd MMMM yyyy") & "|A|"; Format(Fecha_Termino, "dd MMMM yyyy")
    End If
    Print #1,
    Print #2,
    Print #1, "__________________________________________________________________________________________________________________________"
    Print #2, "__________________________________________________________________________________________________________________________"
    Print #1,
    Print #2,
End Sub

Private Sub Finalizar_Reporte()
    Close #1, #2
End Sub

Private Sub Obtiene_Informacion_Checadas(ByRef Ruta_Archivo As String, ByRef Mensaje As String)
'Genera los archivos de checadas
'Dim dwEnrollNumber As Long
Dim dwEnrollNumber As String
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim I As Long
Dim dwYear As Long
Dim dwMonth As Long
Dim dwDay As Long
Dim dwHour As Long
Dim dwMinute As Long
Dim dwSecond As Long
Dim dwWorkcode As Long
Dim dwReserved As Long
Dim nomruta As String
Dim gid As String
Dim IP As String
Dim Puerto As String
Dim Maquina As String
Dim Descripcion_Equipo As String
Dim Checador_ID As String
Dim Rs_Alta_Adm_Asistencias_Registro_Checadores As rdoResultset
Dim Rs_Consulta_Adm_Asistencias_Registro_Checadores As rdoResultset
    
Dim fso, txtfile
Dim LogStr As String
Dim dato As String
Dim dwvalue As Long
Dim res As Boolean
Dim Cuenta As Long
Dim cad As String
Dim aux As String
Dim Rs_Consulta_Dispositivos_Empresa As rdoResultset     'Informacion dek dispositivo
Dim Rs_Consulta_Informacion_Dispositivo As rdoResultset     'Informacion dek dispositivo
Dim Rs_Alta_ As rdoResultset     'Informacion dek dispositivo
Dim Fecha_Checada As String

On Error GoTo HANDLER
    Ruta_Archivo = ""
    'Configura el dispositivo
    'Llena informacion del dispositivo
    IP = ""
    Puerto = 0
    Maquina = 0
    CZKEM1.Disconnect
    Me.Refresh
    Tab_Importacion_Checadas.Tab = 1
    Me.Refresh
    Me.MousePointer = 11
    'Consulta los checadores de la empresa seleccionada
    Mi_SQL = "SELECT Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Empresas_Equipos_Identificacion.Equipo_ID,Cat_Equipos_Identificadores.No_Equipo"
    Mi_SQL = Mi_SQL & " FROM Cat_Empresas_Equipos_Identificacion,Cat_Equipos_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empresas_Equipos_Identificacion.Equipo_ID=Cat_Equipos_Identificadores.Equipo_ID"
    If Cmb_Adm_Importacion_Empresa_Automatico.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empresas_Equipos_Identificacion.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Equipos_Identificadores.No_Equipo"
    Set Rs_Consulta_Dispositivos_Empresa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Dispositivos_Empresa.EOF Then
        With Rs_Consulta_Dispositivos_Empresa
            Set Rs_Alta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias_Registro_Checadores")
            While Not .EOF
                Txt_Importacion_Keri_Log.Text = ""
                'De acuerdo a los checadores de las empresas inicia la extraccion de informacion
                'Consulta la informacion para conectarse
                Mi_SQL = "SELECT Direccion_IP, Puerto_IP, No_Equipo, Descripcion "
                Mi_SQL = Mi_SQL & " FROM Cat_Equipos_Identificadores"
                Mi_SQL = Mi_SQL & " WHERE Equipo_ID = '" & .rdoColumns("Equipo_ID") & "'"
                Set Rs_Consulta_Informacion_Dispositivo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                With Rs_Consulta_Informacion_Dispositivo
                    If Not .EOF Then
                        IP = .rdoColumns("Direccion_IP")
                        Puerto = .rdoColumns("Puerto_IP")
                        Maquina = .rdoColumns("No_Equipo")
                        Descripcion_Equipo = .rdoColumns("Descripcion")
                        Checador_ID = Rs_Consulta_Dispositivos_Empresa.rdoColumns("Equipo_ID")
                        If CZKEM1.Connect_Net(IP, Puerto) Then
                            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                Descripcion_Equipo & ", conectado"
                            Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                        Else
                            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                Descripcion_Equipo & ", no se pudo conectar"
                            Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                            GoTo SIGUIENTE:
                        End If
                        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                            "Obteniendo Información de dispositivo: " & Descripcion_Equipo
                        Me.MousePointer = 0
                        'Inicia la recoleccion de datos
                        'Abre los logs en los checadores para extraer la informacion
                        res = CZKEM1.GetDeviceStatus(Maquina, 6, dwvalue)
                        Me.Refresh
                        If res Then
                            If dwvalue = 0 Then
                                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                    "No existen checadas"
                                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                                GoTo SIGUIENTE1:
                            End If
                        End If
                        Me.Refresh
                        If CZKEM1.ReadGeneralLogData(Maquina) Then
                            Me.Refresh
                                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                    "Equipo listo, inicia generación de información ..."
                                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                            CZKEM1.ReadAllUserID Maquina
                            Me.Refresh
                            'While CZKEM1.GetGeneralLogDataStr(1, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
                            While CZKEM1.SSR_GetGeneralLogData(Maquina, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode)
                                If CStr(dwInOutMode) > 1 Then
                                    dwInOutMode = 0
                                End If
                                Me.Refresh
                                If IsNumeric(dwEnrollNumber) Then
                                    cad = Trim(Str(dwEnrollNumber))
                                    Fecha_Checada = Format(dwYear, "0000") & "-" + Format(dwMonth, "00") & "-" & Format(dwDay, "00") & " " & Format(dwHour, "00") & ":" & Format(dwMinute, "00") & ":" & Format(dwSecond, "00")
                                    
                                    If dwEnrollNumber > 0 Then
                                        If DateDiff("d", Dtp_Importacion_Fecha_Inicio_Automatico.Value, CDate(Fecha_Checada)) >= 0 And DateDiff("d", Dtp_Importacion_Fecha_Termino_Automatico.Value, CDate(Fecha_Checada)) <= 0 Then
                                            aux = IIf(IsNull(Fecha_Checada), "", Fecha_Checada)
                                            'Checador_ID = Rs_Consulta_Dispositivos_Empresa.rdoColumns("Equipo_ID")
                                            If Len(aux) = 19 Then
                                                'Forma la cadena No_Empleado, Checador_ID,
                                                'Modo de Verificacion(1=Huella, 0=Contraseña), Modo_ES, Fecha MM/dd/yyyy hh:mm:ss
                                                aux = cad & Chr(9) & Fecha_Checada
                                                aux = aux & Chr(9) & CStr(dwVerifyMode)
                                                aux = aux & Chr(9) & CStr(dwInOutMode)
                                                aux = aux & Chr(9) & CStr(dwVerifyMode)
                                                aux = aux & Chr(9) & CStr(dwInOutMode)
                                                aux = aux & Chr(9) & Checador_ID
                                                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                                    "Registro importado:" & aux
                                                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                                                aux = ""
                                                Cuenta = Cuenta + 1
                                                'Verifica si el registro existe, para no duplicar información
                                                Mi_SQL = "SELECT * FROM Adm_Asistencias_Registro_Checadores AARC"
                                                Mi_SQL = Mi_SQL & " WHERE AARC.No_Tarjeta = '" & Trim(Str(dwEnrollNumber)) & "'"
                                                Mi_SQL = Mi_SQL & " AND AARC.Fecha = '" & Format(CDate(Fecha_Checada), "MM/dd/yyyy") & "'"
                                                Mi_SQL = Mi_SQL & " AND AARC.Hora = '" & "12/30/1899 " & Format(CDate(Fecha_Checada), "HH:mm") & "'"
                                                Mi_SQL = Mi_SQL & " AND AARC.No_Equipo = '" & Maquina & "'"
                                                If Cmb_Adm_Importacion_Empresa_Automatico.ListIndex > 0 Then
                                                    Mi_SQL = Mi_SQL & " AND AARC.Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000") & "'"
                                                End If
                                                'Mi_SQL = Mi_SQL & " AND AARC.E_S = '" & CStr(dwInOutMode) & "'"
                                                'Mi_SQL = Mi_SQL & " AND AARC.Verificacion = '" & CStr(dwVerifyMode) & "'"
                                                Set Rs_Consulta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                                If Rs_Consulta_Adm_Asistencias_Registro_Checadores.EOF Then
                                                    'Guarda el registro de las checadas en la base de datos
                                                    Mi_SQL = "INSERT INTO Adm_Asistencias_Registro_Checadores(No_Tarjeta,Fecha,Hora,Fecha_Importacion"
                                                    Mi_SQL = Mi_SQL & " ,No_Equipo,Equipo_ID,Empresa_ID,E_S,IP,Verificacion)"
                                                    Mi_SQL = Mi_SQL & " VALUES('" & Trim(Str(dwEnrollNumber)) & "'"
                                                    Mi_SQL = Mi_SQL & " , '" & Format(CDate(Fecha_Checada), "MM/dd/yyyy") & "'"
                                                    Mi_SQL = Mi_SQL & " , '12/30/1899 " & Format(CDate(Fecha_Checada), "HH:mm") & "'"
                                                    Mi_SQL = Mi_SQL & " , '" & Format(Now, "MM/dd/yyyy") & "'"
                                                    Mi_SQL = Mi_SQL & " , " & Maquina & ""
                                                    Mi_SQL = Mi_SQL & " , '" & Checador_ID & "'"
                                                    If Cmb_Adm_Importacion_Empresa_Automatico.ListIndex > 0 Then
                                                        Mi_SQL = Mi_SQL & " , '" & Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000") & "'"
                                                    Else
                                                        Mi_SQL = Mi_SQL & " , '00001'"
                                                    End If
                                                    Mi_SQL = Mi_SQL & " , '" & CStr(dwInOutMode) & "'"
                                                    Mi_SQL = Mi_SQL & " , '" & IP & "'"
                                                    Mi_SQL = Mi_SQL & " , '" & CStr(dwVerifyMode) & "')"
                                                    Conexion_Base.Execute (Mi_SQL)
'                                                    With Rs_Alta_Adm_Asistencias_Registro_Checadores
'                                                        .AddNew
'                                                            '.rdoColumns ("No_Movimiento") = Se genera automaticamente
'                                                            .rdoColumns("No_Tarjeta") = Trim(Str(dwEnrollNumber))
'                                                            .rdoColumns("Fecha") = Format(CDate(Fecha_Checada), "MM/dd/yyyy")
'                                                            .rdoColumns("Hora") = "12/30/1899 " & Format(CDate(Fecha_Checada), "HH:mm")
'                                                            .rdoColumns("Fecha_Importacion") = Format(Now, "MM/dd/yyyy")
'                                                            .rdoColumns("No_Equipo") = Maquina
'                                                            .rdoColumns("Equipo_ID") = Checador_ID
'                                                            '.rdoColumns("Empresa_ID") = Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000")
'                                                            .rdoColumns("Empresa_ID") = "00001"
'                                                            .rdoColumns("E_S") = CStr(dwInOutMode)
'                                                            .rdoColumns("IP") = IP
'                                                            .rdoColumns("Verificacion") = CStr(dwVerifyMode)
'                                                        .Update
'                                                    End With
                                                Else
                                                    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & aux & ", Importado anteriormente"
                                                    Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                                                End If
                                                Rs_Consulta_Adm_Asistencias_Registro_Checadores.Close
                                            End If
                                        End If
                                    End If
                                End If
                                DoEvents
                            Wend
                                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                    "Termino de importación equipo: " & Descripcion_Equipo
                                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                            End If
                        Else
                            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                "No hay informacion para el dispositivo " & Descripcion_Equipo & ", favor de verificar"
                            Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                            GoTo SIGUIENTE2:
                        End If
                    End With
                    Set Rs_Consulta_Informacion_Dispositivo = Nothing
SIGUIENTE:
SIGUIENTE1:
SIGUIENTE2:
                .MoveNext
            Wend
        End With
    Else
        MsgBox "La empresa seleccionada no tiene checadores asignados", vbInformation + vbOKOnly, Me.Caption
    End If
    Rs_Consulta_Dispositivos_Empresa.Close
    Me.Refresh
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Close #1
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            ": " & Err.Description
    Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
End Sub

Private Sub Cmb_Adm_Importacion_Empresa_Manual_Click()
    If Cmb_Adm_Importacion_Empresa_Manual.ListIndex > -1 Then
        'Llena los checadores de la empresa seleccionada
        Call Conectar_Ayudante.Llena_Combo_Item("CEI.Equipo_ID, CAST(CEI.No_Equipo as varchar)+' '+CEI.Descripcion", _
            "Cat_Equipos_Identificadores CEI, Cat_Empresas_Equipos_Identificacion CEEI WHERE CEI.Equipo_ID = CEEI.Equipo_ID AND CEEI.Empresa_Id = '" & Format(Cmb_Adm_Importacion_Empresa_Manual.ItemData(Cmb_Adm_Importacion_Empresa_Manual.ListIndex), "00000") & "'", Cmb_Adm_Importacion_Checador, 0, "No_Equipo", , False, "TODAS")
        If Cmb_Adm_Importacion_Checador.ListCount > 0 Then
            Cmb_Adm_Importacion_Checador.ListIndex = 0
        End If
    End If
End Sub

Private Sub Obtiene_Informacion_Checadas_Archivo(Nombre_Archivo As String, Checador_ID As String, Empresa_ID As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim IP As String
Dim Maquina As String
Dim Ruta_Archivo As String
Dim nomruta As String
Dim linea
Dim Datos() As String
Dim Contador As Integer
Dim Cadena  As String
Dim num As Integer
Dim Rs_Consulta_Adm_Asistencias_Registro_Checadores As rdoResultset
Dim Rs_Alta_Adm_Asistencias_Registro_Checadores As rdoResultset

On Error GoTo HANDLER:
    Me.MousePointer = 11
    Me.Refresh
    Txt_Importacion_Keri_Log.Text = "Generando Lista de empleados para la fecha seleccionada"
    If Nombre_Archivo = "" Then
        MsgBox "Seleccione un archivo para importar", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe archivo para importar información"
        Me.MousePointer = 0
        Exit Sub
    End If
    Ruta_Archivo = Nombre_Archivo
    'valida que el archivo de exportacion exista en la ruta proporcionada
    If Len(Dir$(Ruta_Archivo)) <= 0 Then
        MsgBox "El archivo no contiene información o no existe, favor de verificar", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "El archivo no contiene información"
        Me.MousePointer = 0
        PrgBar_Importacion.Visible = False
        Exit Sub
    End If
    nomruta = Ruta_Temporal & Checador_ID & "_" & Format(Now, "MMddyyyy_HHmmss") & ".dat"
    Me.Refresh
    'Forma la cadena No_Empleado(ya), Checador_ID(ya), Modo de Verificacion(ya), Modo_ES(ya), Fecha MM/dd/yyyy(ya) hh:mm:ss(ya)
    'Grid_Importacion.AddItem "No. Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Empleado_ID" & Chr(9) & "Registrado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Checador_ID" & Chr(9) & "Modo_Verificacion" & Chr(9) & "E/S"
    Open Ruta_Archivo For Input As #1   'Abre el archivo del sistema de checadas
    Open nomruta For Output As #2   'Abre el archivo temporal
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
        "Incicio de depuracion de información para generar la lista"
    Do While Not EOF(1)
        Me.Refresh
        Line Input #1, linea
        Datos() = Split(linea, Chr(9))
        Debug.Print linea
        If Datos(0) = "Fin_Archivo" Then
            Exit Do
        End If
        If Mid(Datos(0), 1, 3) <> "SN=" And Mid(Datos(0), 1, 8) <> "CHECKSUM" Then
            If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
                DateDiff("d", Format(Dtp_Importacion_Fecha_Termino.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
                Print #2, linea & Chr(9) & Checador_ID
                Contador = Contador + 1
                Me.Refresh
            End If
        End If
    Loop
    Close #1, #2
    Me.Refresh
    PrgBar_Importacion.Visible = True
    PrgBar_Importacion.Value = 0
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
        "No de registros localizados: " & Contador
    Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
    If Contador > 0 Then
        PrgBar_Importacion.Max = Contador
    Else
        PrgBar_Importacion.Visible = False
        MsgBox "No existe información para la fecha seleccionada", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe información para la fecha seleccionada"
        Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
        Me.MousePointer = 0
        Exit Sub
    End If
    If Len(Dir$(nomruta)) <= 0 Then
        MsgBox "No existe información en el rango de fechas seleccionado", vbInformation + vbOKOnly, Me.Caption
        Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
            "No existe información en el rango de fechas seleccionado"
        Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
        Me.MousePointer = 0
        Exit Sub
    End If
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Generando Lista..."
    Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
    'Consulta la maquina
    Mi_SQL = "SELECT No_Equipo, Direccion_IP FROM Cat_Equipos_Identificadores WHERE Equipo_ID = '" & Format(Cmb_Adm_Importacion_Checador.ItemData(Cmb_Adm_Importacion_Checador.ListIndex), "00000") & "' "
    Maquina = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "No_Equipo")
    IP = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Direccion_IP")
    Open nomruta For Input As #1
    Do While Not EOF(1)
        Me.Refresh
        Me.MousePointer = 11
        Line Input #1, linea
        Cadena = ""
        num = 0
        Datos() = Split(linea, Chr(9))
        Maquina = Trim(Datos(2))
        Debug.Print linea
        If DateDiff("d", Format(Dtp_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) >= 0 And _
            DateDiff("d", Format(Dtp_Importacion_Fecha_Termino.Value, "MM/dd/yyyy"), Format(Datos(1), "MM/dd/yyyy")) <= 0 Then
            'Verifica si el registro existe, para no duplicar información
            Mi_SQL = "SELECT No_Movimiento,No_Tarjeta FROM Adm_Asistencias_Registro_Checadores"
            Mi_SQL = Mi_SQL & " WHERE No_Tarjeta='" & CStr(Val(Trim(Datos(0)))) & "'"
            Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Trim(Datos(1)), "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Hora='" & "12/30/1899 " & Format(Trim(Datos(1)), "HH:mm") & "'"
            Mi_SQL = Mi_SQL & " AND No_Equipo='" & Maquina & "'"
            Mi_SQL = Mi_SQL & " AND Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Manual.ItemData(Cmb_Adm_Importacion_Empresa_Manual.ListIndex), "00000") & "'"
            'Mi_SQL = Mi_SQL & " AND E_S='" & CStr(Datos(3)) & "'"
            'Mi_SQL = Mi_SQL & " AND Verificacion='" & CStr(Datos(4)) & "'"
            Set Rs_Consulta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Rs_Consulta_Adm_Asistencias_Registro_Checadores.EOF Then
                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & linea
                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                'Cambio a insert
                Mi_SQL = "INSERT INTO Adm_Asistencias_Registro_Checadores "
                Mi_SQL = Mi_SQL & " (No_Tarjeta, Fecha, Hora, "
                Mi_SQL = Mi_SQL & " Fecha_Importacion, No_Equipo, Equipo_ID, "
                Mi_SQL = Mi_SQL & " Empresa_ID, E_S, IP, "
                Mi_SQL = Mi_SQL & " Verificacion)"
                Mi_SQL = Mi_SQL & " VALUES ("
                Mi_SQL = Mi_SQL & " '" & CStr(Val(Trim(Datos(0)))) & "', "
                Mi_SQL = Mi_SQL & " '" & Format(Trim(Datos(1)), "MM/dd/yyyy") & "',"
                Mi_SQL = Mi_SQL & " '12/30/1899 " & Format(Trim(Datos(1)), "HH:mm") & "',"
                Mi_SQL = Mi_SQL & " '" & Format(Now, "MM/dd/yyyy") & "',"
                Mi_SQL = Mi_SQL & " " & Maquina & ","
                Mi_SQL = Mi_SQL & " '" & Format(Cmb_Adm_Importacion_Checador.ItemData(Cmb_Adm_Importacion_Checador.ListIndex), "00000") & "',"
                Mi_SQL = Mi_SQL & " '" & Format(Cmb_Adm_Importacion_Empresa_Manual.ItemData(Cmb_Adm_Importacion_Empresa_Manual.ListIndex), "00000") & "',"
                Mi_SQL = Mi_SQL & " '" & CStr(Datos(3)) & "',"
                Mi_SQL = Mi_SQL & " '" & IP & "',"
                Mi_SQL = Mi_SQL & " '" & CStr(Datos(4)) & "')"
                Conexion_Base.Execute (Mi_SQL)
                Me.Refresh
'                'Guarda el registro de las checadas en la base de datos
'                Set Rs_Alta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias_Registro_Checadores")
'                    With Rs_Alta_Adm_Asistencias_Registro_Checadores
'                        .AddNew
'                            '.rdoColumns ("No_Movimiento") = Se genera automaticamente
'                            .rdoColumns("No_Tarjeta") = CStr(Val(Trim(Datos(0))))
'                            .rdoColumns("Fecha") = Format(Trim(Datos(1)), "MM/dd/yyyy")
'                            .rdoColumns("Hora") = "12/30/1899 " & Format(Trim(Datos(1)), "HH:mm")
'                            .rdoColumns("Fecha_Importacion") = Format(Now, "MM/dd/yyyy")
'                            .rdoColumns("No_Equipo") = Maquina
'                            .rdoColumns("Equipo_ID") = Format(Cmb_Adm_Importacion_Checador.ItemData(Cmb_Adm_Importacion_Checador.ListIndex), "00000")
'                            .rdoColumns("Empresa_ID") = Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000")
'                            .rdoColumns("E_S") = CStr(Datos(3))
'                            .rdoColumns("IP") = IP
'                            .rdoColumns("Verificacion") = CStr(Datos(4))
'                        .Update
'                        .Close
'                    End With
'                Set Rs_Alta_Adm_Asistencias_Registro_Checadores = Nothing
            Else
                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & linea & ", Importado anteriormente"
                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
            End If
            Set Rs_Consulta_Adm_Asistencias_Registro_Checadores = Nothing
            Debug.Print Cadena
            If num <> 0 Then Grid_Importacion.AddItem Cadena
            PrgBar_Importacion.Value = PrgBar_Importacion.Value + 1
            Me.Refresh
        End If
    Loop
    Close #1
    Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
        "Importación de Información Terminada..."
    Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
    PrgBar_Importacion.Visible = False
    Me.Refresh
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    Close #1, #2
    Grid_Importacion.Rows = 0
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Depurar_Lista
'DESCRIPCION: Depura la lista de información para obtener la hora de entrada y salida
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 10-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Depurar_Lista()
Dim Rs_Consulta_Cat_Empleados As rdoResultset   'informacion del turno del empleado
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Checada As rdoResultset
Dim Cont_Fila As Integer                'Recorre el grid de Grid_Importacion_Keri_System
Dim Cont_Fila_2 As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada
Dim Cont_Fila_3 As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada para buscar las horas intermedias
Dim Encontrado As Boolean               'Indica si se ha encontrado el registro en el grid
Dim Cont_Columna As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada
Dim Fila_Encontrado As Integer          'Indica la fila donde se encontro el registro
Dim Columna_Grid As Integer             'Hace referencia a la columna en que se colocara la información de la hora
Dim Registrado As String
Dim Nombre_Empleado As String
Dim Empleado_ID As String
Dim Empleado_ID_Agregar As String
Dim Dias As Integer
Dim Inicio As Integer
Dim Hora_Entrada As String
Dim Hora_Salida As String
Dim Hora_Comida As String
Dim Hora_Comida2 As String
Dim No_Checadas As Integer
Dim No_Tarjeta As String
Dim Nombre As String
Dim Checador As String
Dim empledo_id As String
Dim Horas As Integer
Dim Fecha As String

    'Llena el grid de acuerdo a la consulta
    Grid_Importacion.Rows = 0
    Grid_Importacion.Cols = 12
    Grid_Importacion_Lista_Depurada.Rows = 0
    Me.MousePointer = 11
    Me.Refresh
    Prbar_Asistencia.Visible = True
    Prbar_Asistencia.Value = 0
    Prbar_Asistencia.Min = 0
    Hora_Entrada = ""
    Hora_Salida = ""
    Hora_Comida = ""
    Hora_Comida2 = ""
    No_Tarjeta = ""
    Nombre = ""
    Checador = ""
    'Consulta los registros de checadas
            
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC,Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Turno_ID=Cat_Turnos.Turno_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno>=0"       'Horario dentro de la misma jornada laboral
    If Cmb_Empleado.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID = '" & Format(Cmb_Empleado.ItemData(Cmb_Empleado.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND NOT EXISTS ("
    Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     FROM ("
    Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
    Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
    Mi_SQL = Mi_SQL & "         From Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
    Mi_SQL = Mi_SQL & "         Where Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
    Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "                 AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
    Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = AARC.Fecha"
    Mi_SQL = Mi_SQL & " )"
    Mi_SQL = Mi_SQL & " ORDER BY CE.No_Tarjeta,AARC.Fecha,AARC.Hora"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Call Encabezado_Reporte("IMPORTACION ASISTENCIAS", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        With Rs_Consulta_Cat_Clientes
            Prbar_Asistencia.Max = Rs_Consulta_Cat_Clientes.RowCount
            Me.Refresh
            Grid_Importacion_Lista_Depurada.Rows = 0
            Grid_Importacion_Lista_Depurada.Cols = 12
            Encontrado = False
            Fila_Encontrado = 0
            Me.MousePointer = 11
            Me.Refresh
'            Call Encabezado_Reporte("IMPORTACION ASISTENCIAS", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
            Grid_Importacion_Lista_Depurada.AddItem "Fecha" & Chr(9) & "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Entrada" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
                      '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
            Print #1, "No Nomina   Empleado                              Entrada   S.Comida   E.Comida   Salida   Horas"
            Print #2, "Fecha|No Nomina |Empleado|||Entrada|Salida|Horas|Checador"
            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Iniciando depuración de información... (asignando hora de entrada, comida y salida)"
            'LLenado de la informacion
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'valida el empleado
                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Empleado_ID_Agregar = Empleado_ID
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
                    No_Checadas = 0
                    Checador = .rdoColumns("Equipo_ID")
                    'Limpia las variables
                    Hora_Entrada = Format(.rdoColumns("Hora"), "HH:mm:ss")
                    Hora_Salida = "0"
                    Hora_Comida = "0"
                    Hora_Comida2 = "0"
                End If
                No_Checadas = No_Checadas + 1
                'Prepara las fechas del registro
                Select Case No_Checadas
                    Case 2
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                    Case 3
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                    Case 4
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
    
                    Case Else
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                End Select
                .MoveNext
                If Not .EOF Then
                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Realiza el calculo de horas trabajadas
                        Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                        If Horas < 0 Then Horas = Horas + 24
                        'Valida las horas de comida
                        If (Hora_Entrada = Hora_Comida) Then
                            Hora_Comida = 0
                        End If
                        If (Hora_Salida = Hora_Comida2) Then
                            Hora_Comida2 = 0
                        End If
                        If (Hora_Entrada = Hora_Salida) Then
                            Hora_Salida = 0
                        End If
                        'Agrega el registro
                        Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & "S" _
                            & Chr(9) & Empleado_ID_Agregar _
                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                            Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                            Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                        Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
                            Val(Horas); "|"; Checador
                        Me.Refresh
                        Empleado_ID = ""
                    End If
                Else
                    'Realiza el calculo de horas trabajadas
                    Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                    If Horas < 0 Then Horas = Horas + 24
                    'Valida las horas de comida
                    If (Hora_Entrada = Hora_Comida) Then
                        Hora_Comida = 0
                    End If
                    If (Hora_Salida = Hora_Comida2) Then
                        Hora_Comida2 = 0
                    End If
                    If (Hora_Entrada = Hora_Salida) Then
                        Hora_Salida = 0
                    End If
                    If Horas < 0 Then
                        Horas = 0
                    End If
                    'Agrega el registro
                    Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                        & Chr(9) & No_Tarjeta _
                        & Chr(9) & Nombre _
                        & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                        & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                        & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                        & Chr(9) & "S" _
                        & Chr(9) & Empleado_ID_Agregar _
                        & Chr(9) & Checador
                    Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
                        Spc(2); Mid(Nombre, 40); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Round(Horas, 2), "#0.00"), 8)
                    Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                        Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                        Format(Hora_Comida, "HH:mm:ss"); "|"; _
                        Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                        Format(Hora_Salida, "HH:mm:ss"); "|"; _
                        Val(Horas); "|"; Checador
                     Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    
    'Turnos de 2 días
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC,Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Turno_ID=Cat_Turnos.Turno_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
    If Cmb_Empleado.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID = '" & Format(Cmb_Empleado.ItemData(Cmb_Empleado.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND NOT EXISTS ("
    Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     FROM ("
    Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
    Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
    Mi_SQL = Mi_SQL & "         From Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
    Mi_SQL = Mi_SQL & "         Where Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
    Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "                 AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
    Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = AARC.Fecha"
    Mi_SQL = Mi_SQL & " )"
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AARC.Hora DESC"

    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        If Grid_Importacion_Lista_Depurada.Rows <= 0 Then
        Grid_Importacion_Lista_Depurada.Cols = 12
        Grid_Importacion_Lista_Depurada.AddItem "Fecha" & Chr(9) & "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Entrada" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
        End If
        With Rs_Consulta_Cat_Clientes
            Prbar_Asistencia.Value = 0
            Prbar_Asistencia.Max = Rs_Consulta_Cat_Clientes.RowCount
            Me.Refresh
            Encontrado = False
            Fila_Encontrado = 0
            Me.MousePointer = 11
            Me.Refresh
            'LLenado de la informacion
            'Se recorre el recorset
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'Valida el empleado
                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Empleado_ID_Agregar = Empleado_ID
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
                    No_Checadas = 0
                    Checador = .rdoColumns("Equipo_ID")
                    'Limpia las variables
                    Hora_Entrada = "0"
                    Hora_Salida = "0"
                    Hora_Comida = "0"
                    Hora_Comida2 = "0"
                End If
                No_Checadas = No_Checadas + 1
                'Prepara las fechas del registro
                Select Case No_Checadas
                    Case 2
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                    Case 3
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                    Case 4
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
                    Case Else
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                End Select
                'Consulta la hora de entrada del día
                Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Fecha & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora>'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Checada.EOF Then
                    Hora_Entrada = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                    Hora_Salida = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                End If
                Rs_Consulta_Checada.Close
                'Consulta la hora de salida del día siguiente
                Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Format(DateAdd("d", 1, Fecha), "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora<'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Checada.EOF Then
                    If Hora_Entrada = "0" Then 'Si no tuvo hora de entrada le asigna su entrada coo su salida y no le calcula tiempo extra
                        Hora_Entrada = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                    End If
                    Hora_Salida = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                End If
                Rs_Consulta_Checada.Close
                .MoveNext
                If Not .EOF Then
                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Realiza el calculo de horas trabajadas
                        Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                        If Horas < 0 Then Horas = Horas + 24
                        'Valida las horas de comida
                        If (Hora_Entrada = Hora_Comida) Then
                            Hora_Comida = 0
                        End If
                        If (Hora_Salida = Hora_Comida2) Then
                            Hora_Comida2 = 0
                        End If
                        If (Hora_Entrada = Hora_Salida) Then
                            Hora_Salida = 0
                        End If
                        'Agrega el registro si tiene hora de entrada
                        If Hora_Entrada <> "0" Then
                            Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                                & Chr(9) & No_Tarjeta _
                                & Chr(9) & Nombre _
                                & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & "S" _
                                & Chr(9) & Empleado_ID_Agregar _
                                & Chr(9) & Checador
                            Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                                Mid(Nombre, 40); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                            Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                                Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                                Format(Hora_Salida, "HH:mm:ss"); "|"; _
                                Val(Horas)
                        End If
                    End If
                Else
                    Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                    If Horas < 0 Then Horas = Horas + 24
                    'Valida las horas de comida
                    If (Hora_Entrada = Hora_Comida) Then
                        Hora_Comida = 0
                    End If
                    If (Hora_Salida = Hora_Comida2) Then
                        Hora_Comida2 = 0
                    End If
                    If (Hora_Entrada = Hora_Salida) Then
                        Hora_Salida = 0
                    End If
                    'Agrega el registro si tiene hora de entrada
                    If Hora_Entrada <> "0" Then
                        Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & "S" _
                            & Chr(9) & Empleado_ID_Agregar _
                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
                            Spc(2); Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Round(Horas, 2), "#0.00"), 8)
                        Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
                            Val(Horas); "|"; Checador
                    End If
                    Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    
    'PrgBar_Importacion.Value = PrgBar_Importacion.Value + 1
    Me.Refresh
    Prbar_Asistencia.Visible = False
    With Grid_Importacion_Lista_Depurada
        If Grid_Importacion_Lista_Depurada.Rows > 1 Then
            .FixedRows = 1
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(0) = 1200    'Fecha
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(1) = 700     'No Tarjeta
            .ColAlignment(1) = flexAlignRightCenter
            .ColWidth(2) = 4200    'Empleado
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(3) = 800     'entrada
            .ColAlignment(3) = flexAlignCenterCenter
            .ColWidth(4) = 0       'comida
            .ColWidth(5) = 0       'comida
            .ColWidth(6) = 800     'salida
            .ColAlignment(6) = flexAlignCenterCenter
            .ColWidth(7) = 700     'horas
            .ColWidth(8) = 0       'horas
            .ColWidth(9) = 0       'Registrado
            .ColWidth(10) = 0      'Empleado_ID
            .ColWidth(11) = 0      'Checador
        End If
        If Grid_Importacion_Lista_Depurada.Col > 0 Then
            Grid_Importacion_Lista_Depurada.Col = 1
            Grid_Importacion_Lista_Depurada.Sort = flexSortGenericAscending
        End If
        Call Finalizar_Reporte
    End With
    Me.MousePointer = 0
    Me.Refresh
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Depurar_Lista_Turnos_Flexibles
'DESCRIPCION: Depura la lista de información para obtener la hora de entrada y salida
'PARAMETROS :
'CREO       : Antonio Salvador Benavides Guardado
'FECHA_CREO : 26/Marzo/2017
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Depurar_Lista_Turnos_Flexibles()
Dim Rs_Consulta_Cat_Empleados As rdoResultset   'informacion del turno del empleado
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Checada As rdoResultset
Dim Cont_Fila As Integer                'Recorre el grid de Grid_Importacion_Keri_System
Dim Cont_Fila_2 As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada
Dim Cont_Fila_3 As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada para buscar las horas intermedias
Dim Encontrado As Boolean               'Indica si se ha encontrado el registro en el grid
Dim Cont_Columna As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada
Dim Fila_Encontrado As Integer          'Indica la fila donde se encontro el registro
Dim Columna_Grid As Integer             'Hace referencia a la columna en que se colocara la información de la hora
Dim Registrado As String
Dim Nombre_Empleado As String
Dim Empleado_ID As String
Dim Empleado_ID_Agregar As String
Dim Dias As Integer
Dim Inicio As Integer
Dim Hora_Entrada As String
Dim Hora_Salida As String
Dim Hora_Comida As String
Dim Hora_Comida2 As String
Dim No_Checadas As Integer
Dim No_Tarjeta As String
Dim Nombre As String
Dim Checador As String
Dim empledo_id As String
Dim Horas As Integer
Dim Fecha As String

    'Llena el grid de acuerdo a la consulta
    If Grid_Importacion.Rows < 2 Then
        Grid_Importacion.Rows = 0
        Grid_Importacion.Cols = 12
        If Grid_Importacion_Lista_Depurada.Rows < 2 Then
            Grid_Importacion_Lista_Depurada.Rows = 0
        End If
        Hora_Entrada = ""
        Hora_Salida = ""
        Hora_Comida = ""
        Hora_Comida2 = ""
        No_Tarjeta = ""
        Nombre = ""
        Checador = ""
    End If
    Me.MousePointer = 11
    Me.Refresh
    Prbar_Asistencia.Visible = True
    Prbar_Asistencia.Value = 0
    Prbar_Asistencia.Min = 0
    'Consulta los registros de checadas
            
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    If Cmb_Empleado.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID = '" & Format(Cmb_Empleado.ItemData(Cmb_Empleado.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND EXISTS ("
    Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     FROM ("
    Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
    Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
    Mi_SQL = Mi_SQL & "         From Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
    Mi_SQL = Mi_SQL & "         Where Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
    Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "                 AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "             AND (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) >= 0"
    Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
    Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = AARC.Fecha"
    Mi_SQL = Mi_SQL & " )"
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AARC.Hora DESC"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        With Rs_Consulta_Cat_Clientes
            Prbar_Asistencia.Max = Rs_Consulta_Cat_Clientes.RowCount
            Me.Refresh
            If Grid_Importacion_Lista_Depurada.Rows < 2 Then
                Call Encabezado_Reporte("IMPORTACION ASISTENCIAS", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
                Grid_Importacion_Lista_Depurada.Rows = 0
                Grid_Importacion_Lista_Depurada.Cols = 12
                Encontrado = False
                Fila_Encontrado = 0
                Me.MousePointer = 11
                Me.Refresh
    '            Call Encabezado_Reporte("IMPORTACION ASISTENCIAS", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
                Grid_Importacion_Lista_Depurada.AddItem "Fecha" & Chr(9) & "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Entrada" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
                          '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                Print #1, "No Nomina   Empleado                              Entrada   S.Comida   E.Comida   Salida   Horas"
                Print #2, "Fecha|No Nomina |Empleado|||Entrada|Salida|Horas|Checador"
                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Iniciando depuración de información... (asignando hora de entrada, comida y salida)"
                'LLenado de la informacion
            Else
                Open Ruta_Temporal & Opcion & ".txt" For Output As #1
                Open Ruta_Temporal & Opcion & "xls.txt" For Output As #2 'Reporte a xls
            End If
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'valida el empleado
                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Empleado_ID_Agregar = Empleado_ID
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
                    No_Checadas = 0
                    Checador = .rdoColumns("Equipo_ID")
                    'Limpia las variables
                    Hora_Entrada = Format(.rdoColumns("Hora"), "HH:mm:ss")
                    Hora_Salida = "0"
                    Hora_Comida = "0"
                    Hora_Comida2 = "0"
                End If
                No_Checadas = No_Checadas + 1
                'Prepara las fechas del registro
                Select Case No_Checadas
                    Case 2
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                    Case 3
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                    Case 4
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
    
                    Case Else
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                End Select
                .MoveNext
                If Not .EOF Then
                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Realiza el calculo de horas trabajadas
                        Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                        If Horas < 0 Then Horas = Horas + 24
                        'Valida las horas de comida
                        If (Hora_Entrada = Hora_Comida) Then
                            Hora_Comida = 0
                        End If
                        If (Hora_Salida = Hora_Comida2) Then
                            Hora_Comida2 = 0
                        End If
                        If (Hora_Entrada = Hora_Salida) Then
                            Hora_Salida = 0
                        End If
                        'Agrega el registro
                        Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & "S" _
                            & Chr(9) & Empleado_ID_Agregar _
                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                            Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                            Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                        Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
                            Val(Horas); "|"; Checador
                        Me.Refresh
                        Empleado_ID = ""
                    End If
                Else
                    'Realiza el calculo de horas trabajadas
                    Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                    If Horas < 0 Then Horas = Horas + 24
                    'Valida las horas de comida
                    If (Hora_Entrada = Hora_Comida) Then
                        Hora_Comida = 0
                    End If
                    If (Hora_Salida = Hora_Comida2) Then
                        Hora_Comida2 = 0
                    End If
                    If (Hora_Entrada = Hora_Salida) Then
                        Hora_Salida = 0
                    End If
                    If Horas < 0 Then
                        Horas = 0
                    End If
                    'Agrega el registro
                    Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                        & Chr(9) & No_Tarjeta _
                        & Chr(9) & Nombre _
                        & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                        & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                        & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                        & Chr(9) & "S" _
                        & Chr(9) & Empleado_ID_Agregar _
                        & Chr(9) & Checador
                    Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
                        Spc(2); Mid(Nombre, 40); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                        Conectar_Ayudante.Alinea_Derecha(Format(Round(Horas, 2), "#0.00"), 8)
                    Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                        Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                        Format(Hora_Comida, "HH:mm:ss"); "|"; _
                        Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                        Format(Hora_Salida, "HH:mm:ss"); "|"; _
                        Val(Horas); "|"; Checador
                     Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    
    'Turnos de 2 días
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    If Cmb_Empleado.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID = '" & Format(Cmb_Empleado.ItemData(Cmb_Empleado.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND EXISTS ("
    Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     FROM ("
    Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
    Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
    Mi_SQL = Mi_SQL & "         From Cat_Calendarios_Turnos"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
    Mi_SQL = Mi_SQL & "         Where Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
    Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "                 AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & "             AND (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) < 0"
    Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
    Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
    Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = AARC.Fecha"
    Mi_SQL = Mi_SQL & " )"
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AARC.Hora DESC"

    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        If Grid_Importacion_Lista_Depurada.Rows <= 0 Then
        Grid_Importacion_Lista_Depurada.Cols = 12
        Grid_Importacion_Lista_Depurada.AddItem "Fecha" & Chr(9) & "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Entrada" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
        End If
        With Rs_Consulta_Cat_Clientes
            Prbar_Asistencia.Value = 0
            Prbar_Asistencia.Max = Rs_Consulta_Cat_Clientes.RowCount
            Me.Refresh
            Encontrado = False
            Fila_Encontrado = 0
            Me.MousePointer = 11
            Me.Refresh
            'LLenado de la informacion
            'Se recorre el recorset
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'Valida el empleado
                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Empleado_ID_Agregar = Empleado_ID
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
                    No_Checadas = 0
                    Checador = .rdoColumns("Equipo_ID")
                    'Limpia las variables
                    Hora_Entrada = "0"
                    Hora_Salida = "0"
                    Hora_Comida = "0"
                    Hora_Comida2 = "0"
                End If
                No_Checadas = No_Checadas + 1
                'Prepara las fechas del registro
                Select Case No_Checadas
                    Case 2
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                    Case 3
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                    Case 4
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
                    Case Else
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                End Select
                'Consulta la hora de entrada del día
                Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Fecha & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora>'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Checada.EOF Then
                    Hora_Entrada = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                    Hora_Salida = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                End If
                Rs_Consulta_Checada.Close
                'Consulta la hora de salida del día siguiente
                Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Format(DateAdd("d", 1, Fecha), "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora<'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Checada.EOF Then
                    If Hora_Entrada = "0" Then 'Si no tuvo hora de entrada le asigna su entrada coo su salida y no le calcula tiempo extra
                        Hora_Entrada = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                    End If
                    Hora_Salida = Format(Rs_Consulta_Checada.rdoColumns("Hora"), "HH:mm:ss")
                End If
                Rs_Consulta_Checada.Close
                .MoveNext
                If Not .EOF Then
                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Realiza el calculo de horas trabajadas
                        Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                        If Horas < 0 Then Horas = Horas + 24
                        'Valida las horas de comida
                        If (Hora_Entrada = Hora_Comida) Then
                            Hora_Comida = 0
                        End If
                        If (Hora_Salida = Hora_Comida2) Then
                            Hora_Comida2 = 0
                        End If
                        If (Hora_Entrada = Hora_Salida) Then
                            Hora_Salida = 0
                        End If
                        'Agrega el registro si tiene hora de entrada
                        If Hora_Entrada <> "0" Then
                            Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                                & Chr(9) & No_Tarjeta _
                                & Chr(9) & Nombre _
                                & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & "S" _
                                & Chr(9) & Empleado_ID_Agregar _
                                & Chr(9) & Checador
                            Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                                Mid(Nombre, 40); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                            Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                                Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                                Format(Hora_Salida, "HH:mm:ss"); "|"; _
                                Val(Horas)
                        End If
                    End If
                Else
                    Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                    If Horas < 0 Then Horas = Horas + 24
                    'Valida las horas de comida
                    If (Hora_Entrada = Hora_Comida) Then
                        Hora_Comida = 0
                    End If
                    If (Hora_Salida = Hora_Comida2) Then
                        Hora_Comida2 = 0
                    End If
                    If (Hora_Entrada = Hora_Salida) Then
                        Hora_Salida = 0
                    End If
                    'Agrega el registro si tiene hora de entrada
                    If Hora_Entrada <> "0" Then
                        Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & "S" _
                            & Chr(9) & Empleado_ID_Agregar _
                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
                            Spc(2); Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Round(Horas, 2), "#0.00"), 8)
                        Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
                            Val(Horas); "|"; Checador
                    End If
                    Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    
    'PrgBar_Importacion.Value = PrgBar_Importacion.Value + 1
    Me.Refresh
    Prbar_Asistencia.Visible = False
    With Grid_Importacion_Lista_Depurada
        If Grid_Importacion_Lista_Depurada.Rows > 1 Then
            .FixedRows = 1
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(0) = 1200    'Fecha
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(1) = 700     'No Tarjeta
            .ColAlignment(1) = flexAlignRightCenter
            .ColWidth(2) = 4200    'Empleado
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(3) = 800     'entrada
            .ColAlignment(3) = flexAlignCenterCenter
            .ColWidth(4) = 0       'comida
            .ColWidth(5) = 0       'comida
            .ColWidth(6) = 800     'salida
            .ColAlignment(6) = flexAlignCenterCenter
            .ColWidth(7) = 700     'horas
            .ColWidth(8) = 0       'horas
            .ColWidth(9) = 0       'Registrado
            .ColWidth(10) = 0      'Empleado_ID
            .ColWidth(11) = 0      'Checador
        End If
        If Grid_Importacion_Lista_Depurada.Col > 0 Then
            Grid_Importacion_Lista_Depurada.Col = 1
            Grid_Importacion_Lista_Depurada.Sort = flexSortGenericAscending
        End If
        Call Finalizar_Reporte
    End With
    Me.MousePointer = 0
    Me.Refresh
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Depurar_Lista_Fin_Semana
'DESCRIPCION: Depura la lista de información para obtener la hora de entrada y salida
'             cuando es fin de semana
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Febrero-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Depurar_Lista_Fin_Semana()
Dim Rs_Consulta_Cat_Empleados As rdoResultset   'informacion del turno del empleado
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Checada As rdoResultset
Dim Encontrado As Boolean               'Indica si se ha encontrado el registro en el grid
Dim Cont_Columna As Integer              'Recorre el grid de Grid_Importacion_Lista_Depurada
Dim Fila_Encontrado As Integer          'Indica la fila donde se encontro el registro
Dim Columna_Grid As Integer             'Hace referencia a la columna en que se colocara la información de la hora
Dim Registrado As String
Dim Nombre_Empleado As String
Dim Empleado_ID As String
Dim Dias As Integer
Dim Inicio As Integer
Dim Hora_Entrada As String
Dim Hora_Salida As String
Dim Hora_Comida As String
Dim Hora_Comida2 As String
Dim No_Checadas As Integer
Dim No_Tarjeta As String
Dim Nombre As String
Dim Checador As String
Dim empledo_id As String
Dim Horas As Integer
Dim Fecha As String
Dim Turno_ID As String

    'Llena el grid de acuerdo a la consulta
    Grid_Importacion.Rows = 0
    Grid_Importacion.Cols = 10
    Grid_Importacion_Lista_Depurada.Rows = 0
    Me.MousePointer = 11
    Me.Refresh
    Prbar_Asistencia.Visible = True
    Prbar_Asistencia.Value = 0
    Prbar_Asistencia.Min = 0
    Hora_Entrada = ""
    Hora_Salida = ""
    Hora_Comida = ""
    Hora_Comida2 = ""
    No_Tarjeta = ""
    Nombre = ""
    Checador = ""
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC" ',Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY CE.No_Tarjeta,AARC.Fecha,AARC.Hora"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Clientes
        If Not .EOF Then
            Prbar_Asistencia.Max = Rs_Consulta_Cat_Clientes.RowCount
            Me.Refresh
            Grid_Importacion_Lista_Depurada.Rows = 0
            Grid_Importacion_Lista_Depurada.Cols = 12
            Encontrado = False
            Fila_Encontrado = 0
            Me.MousePointer = 11
            Me.Refresh
            Call Encabezado_Reporte("IMPORTACION ASISTENCIAS", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
            Grid_Importacion_Lista_Depurada.AddItem "Fecha" & Chr(9) & "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Entrada" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
                      '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
            Print #1, "No Nomina   Empleado                              Entrada   S.Comida   E.Comida   Salida   Horas"
            Print #2, "Fecha|No Nomina |Empleado|||Entrada|S.Comida|E.Comida|Salida|Horas|Checador"
            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Iniciando depuración de información... (asignando hora de entrada, comida y salida)"
            'LLenado de la informacion
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'Valida el empleado
                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
                    No_Checadas = 0
                    Checador = .rdoColumns("Equipo_ID")
                    'Si es de turno nocturno y sólo tiene una checada anterior a las 12 hrs. pone todo en 0
                    If (.rdoColumns("Turno_ID") = "00002" Or .rdoColumns("Turno_ID") = "00005" Or .rdoColumns("Turno_ID") = "00007") And DateDiff("n", Format(.rdoColumns("Hora"), "HH:mm"), "12:00") > 0 Then
                        Hora_Entrada = "0"
                    Else
                        Hora_Entrada = Format(.rdoColumns("Hora"), "HH:mm:ss")
                    End If
                    Hora_Salida = "0"
                    Hora_Comida = "0"
                    Hora_Comida2 = "0"
                    Turno_ID = .rdoColumns("Turno_ID")
                End If
                No_Checadas = No_Checadas + 1
                'Prepara las fechas del registro
                Select Case No_Checadas
                    Case 2
                        If Hora_Entrada = "0" Then
                            Hora_Entrada = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        End If
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                    Case 3
                        Hora_Entrada = Hora_Salida
                        Hora_Salida = .rdoColumns("Hora")
                    Case 4
                        Hora_Entrada = Hora_Salida
                        Hora_Salida = .rdoColumns("Hora")
                    Case Else
                        If Hora_Salida <> "0" Then
                            Hora_Entrada = Hora_Salida
                        End If
                        'Si es de turno nocturno y sólo tiene una checada anterior a las 12 hrs. pone todo en 0
                        If (Turno_ID = "00002" Or Turno_ID = "00005" Or Turno_ID = "00007") And No_Checadas = 1 And DateDiff("n", Format(Hora_Entrada, "HH:mm"), "12:00") > 0 Then
                            Hora_Salida = Hora_Entrada
                        Else
                            Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        End If
                End Select
                .MoveNext
                If Not .EOF Then
                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Realiza el calculo de horas trabajadas
                        Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                        If Horas < 0 Then Horas = Horas + 24
                        
                        'Valida las horas de comida
                        If (Hora_Entrada = Hora_Comida) Then
                            Hora_Comida = 0
                        End If
                        If (Hora_Salida = Hora_Comida2) Then
                            Hora_Comida2 = 0
                        End If
                        If (Hora_Entrada = Hora_Salida) Then
                            Hora_Salida = 0
                        End If
                        
                        'Si es de turno nocturno y sólo tiene una checada anterior a las 12 hrs. pone todo en 0
                        If (Turno_ID = "00002" Or Turno_ID = "00005" Or Turno_ID = "00007") And No_Checadas = 1 And DateDiff("n", Format(Hora_Entrada, "HH:mm"), "12:00") > 0 Then
                            Hora_Entrada = 0
                        End If
                                                
                        If Trim(Hora_Entrada) <> 0 Then
                            'Agrega el registro
                            Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                                & Chr(9) & No_Tarjeta _
                                & Chr(9) & Nombre _
                                & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                                & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                                & Chr(9) & "S" _
                                & Chr(9) & Empleado_ID _
                                & Chr(9) & Checador
                            Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                                Mid(Nombre, 40); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                                Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                            Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                                Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida, "HH:mm:ss"); "|"; _
                                Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                                Format(Hora_Salida, "HH:mm:ss"); "|"; _
                                Val(Horas); "|"; Checador
                        End If
                        Me.Refresh
                    End If
                Else
                    Horas = Val(DateDiff("n", Hora_Entrada, Hora_Salida)) / 60
                    'Valida las horas de comida
                    If (Hora_Entrada = Hora_Comida) Then
                        Hora_Comida = 0
                    End If
                    If (Hora_Salida = Hora_Comida2) Then
                        Hora_Comida2 = 0
                    End If
                    If (Hora_Entrada = Hora_Salida) Then
                        Hora_Salida = 0
                    End If
                    
                    'Si es de turno nocturno y sólo tiene una checada anterior a las 12 hrs. pone todo en 0
                    If (Turno_ID = "00002" Or Turno_ID = "00005" Or Turno_ID = "00007") And No_Checadas = 1 And DateDiff("n", Format(Hora_Entrada, "HH:mm"), "12:00") > 0 Then
                        Hora_Entrada = 0
                    End If
                    
                    If Trim(Hora_Entrada) <> 0 Then
                        'Agrega el registro
                        Grid_Importacion_Lista_Depurada.AddItem Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
                            & Chr(9) & "S" _
                            & Chr(9) & Empleado_ID _
                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
                            Spc(2); Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40))); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Round(Horas, 2), "#0.00"), 8)
                        Print #2, Format(Fecha, "dd/MMM/yyyy"); "|"; No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Hora_Entrada, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida, "HH:mm:ss"); "|"; _
                            Format(Hora_Comida2, "HH:mm:ss"); "|"; _
                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
                            Val(Horas); "|"; Checador
                    End If
                    Me.Refresh
                End If
            Wend
        End If
    End With
    Me.Refresh
    Prbar_Asistencia.Visible = False
    With Grid_Importacion_Lista_Depurada
        If Grid_Importacion_Lista_Depurada.Rows > 1 Then
            .FixedRows = 1
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(0) = 1200     'Fecha
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(1) = 700     'No Tarjeta
            .ColAlignment(1) = flexAlignRightCenter
            .ColWidth(2) = 4200    'Empleado
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(3) = 800     'entrada
            .ColAlignment(3) = flexAlignCenterCenter
            .ColWidth(4) = 0       'comida
            .ColWidth(5) = 0       'comida
            .ColWidth(6) = 800     'salida
            .ColAlignment(6) = flexAlignCenterCenter
            .ColWidth(7) = 700     'horas
            .ColWidth(8) = 0       'horas
            .ColWidth(9) = 0       'Registrado
            .ColWidth(10) = 0      'Empleado_ID
            .ColWidth(11) = 0      'Checador
        End If
        Call Finalizar_Reporte
    End With
    Me.MousePointer = 0
    Me.Refresh
End Sub

Public Sub Exportar_Excel_Bien(Archivo_Exportar As String, Ruta As String)
Dim obj_Excel As Object
Dim Fila As Integer, Columna As Integer
Dim Contenido As String, Lineas As Variant
Dim Datos As Variant, MC As Integer
Dim Encabezado As Boolean
Dim Fila_Encabezado As Integer

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
Private Sub Cmb_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If Chk_Adm_Permisos_Consulta_Supervisor.Value = 1 And Cmb_Adm_Permisos_Consulta_Supervisor.ListIndex > -1 Then
            If IsNumeric(Cmb_Empleado.Text) Then
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE No_Tarjeta='" & Trim(Cmb_Empleado.Text) & "'", Cmb_Empleado, 0, "No_Tarjeta")
            Else
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Empleado, 1, "Apellido_Paterno", " OR Nombre LIKE '%" & Trim(Cmb_Empleado.Text) & "%'" & _
                     " OR Apellido_Materno LIKE '%" & Trim(Cmb_Empleado.Text) & "%'", False, "")
            End If
        'Else
        '    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Adm_Permisos_Consulta_Empleado, 1, "No_Tarjeta", " AND (Nombre like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR " & _
                 "Apellido_Paterno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Permisos_Consulta_Empleado.Text) & "%') ", False, "")
        'End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

