VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Frm_Adm_Importacion_Almacenes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleMode       =   0  'User
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Exportar 
      Caption         =   "Exportar"
      Height          =   690
      Left            =   3960
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Limpiar 
      Caption         =   "Limpiar"
      Height          =   690
      Left            =   5925
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   45
      Tag             =   "A"
      Top             =   7920
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   690
      Left            =   7890
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Guardar 
      Caption         =   "Guardar"
      Height          =   690
      Left            =   120
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "A"
      Top             =   7920
      Width           =   1200
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
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      Begin TabDlg.SSTab Tab_Importacion_Checadas 
         Height          =   6195
         Left            =   60
         TabIndex        =   24
         Top             =   1500
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   10927
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Generar Asistencia"
         TabPicture(0)   =   "Frm_Adm_Importacion_Almacenes.frx":1628
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Informacion_Importación_Kery_Systema"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Log"
         TabPicture(1)   =   "Frm_Adm_Importacion_Almacenes.frx":1644
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   5805
            Left            =   -74955
            TabIndex        =   25
            Top             =   360
            Width           =   9015
            Begin VB.TextBox Txt_Importacion_Keri_Log 
               Height          =   5535
               Left            =   90
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   180
               Width           =   8835
            End
         End
         Begin VB.Frame Fra_Informacion_Importación_Kery_Systema 
            BackColor       =   &H00FFFFFF&
            Height          =   5805
            Left            =   45
            TabIndex        =   27
            Top             =   360
            Width           =   9015
            Begin VB.CommandButton Btn_Generar_Asistencias 
               Caption         =   "Generar Asistencia"
               Height          =   270
               Left            =   6900
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   900
               Width           =   2040
            End
            Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Asistencia 
               Height          =   315
               Left            =   1275
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   180
               Width           =   7620
            End
            Begin VB.ComboBox Cmb_Empleado 
               Height          =   315
               ItemData        =   "Frm_Adm_Importacion_Almacenes.frx":1660
               Left            =   1275
               List            =   "Frm_Adm_Importacion_Almacenes.frx":1662
               TabIndex        =   28
               Top             =   540
               Width           =   7620
            End
            Begin MSComctlLib.ProgressBar Prbar_Asistencia 
               Height          =   105
               Left            =   6960
               TabIndex        =   29
               Top             =   1200
               Visible         =   0   'False
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   185
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Importacion_Lista_Depurada 
               Height          =   4455
               Left            =   60
               TabIndex        =   32
               Top             =   1320
               Width           =   8880
               _ExtentX        =   15663
               _ExtentY        =   7858
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Importacion 
               Height          =   600
               Left            =   60
               TabIndex        =   33
               Top             =   1320
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
               TabIndex        =   34
               Top             =   960
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "ddd dd MMM yyyy"
               Format          =   111542275
               CurrentDate     =   39931
            End
            Begin MSComCtl2.DTPicker Dtp_Asistencia_Fecha_Termino 
               Height          =   315
               Left            =   4755
               TabIndex        =   35
               Top             =   960
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "ddd dd MMM yyyy"
               Format          =   111542275
               CurrentDate     =   39931
            End
            Begin VB.Label Lbl_Progreso_Exportacion 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Archivo"
               Height          =   135
               Left            =   4440
               TabIndex        =   49
               Top             =   3480
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha Termino"
               Height          =   195
               Left            =   3405
               TabIndex        =   39
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Fecha Inicio"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   960
               Width           =   870
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
               TabIndex        =   37
               Top             =   240
               Width           =   735
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
               TabIndex        =   36
               Top             =   600
               Width           =   840
            End
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
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   9105
         Begin VB.TextBox Txt_Adm_Importacion_Ruta_Archivo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   15
            Top             =   960
            Width           =   5205
         End
         Begin VB.CommandButton Btn_Ruta_Checador 
            Caption         =   "..."
            Height          =   315
            Left            =   6570
            TabIndex        =   14
            Top             =   975
            Width           =   450
         End
         Begin VB.CommandButton Btn_Importacion 
            Caption         =   "Importar"
            Height          =   690
            Left            =   7785
            Picture         =   "Frm_Adm_Importacion_Almacenes.frx":1664
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   300
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_Adm_Importacion_Checador 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   600
            Width           =   5700
         End
         Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Manual 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   5700
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Inicio 
            Height          =   315
            Left            =   1350
            TabIndex        =   16
            Top             =   1365
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   111542275
            CurrentDate     =   39931
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Termino 
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   1365
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   111542275
            CurrentDate     =   39931
         End
         Begin MSComctlLib.ProgressBar PrgBar_Importacion 
            Height          =   690
            Left            =   7695
            TabIndex        =   18
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   1035
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   1425
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Termino"
            Height          =   195
            Left            =   3645
            TabIndex        =   21
            Top             =   1425
            Width           =   1065
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
            TabIndex        =   19
            Top             =   300
            Width           =   735
         End
      End
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
         TabIndex        =   1
         Top             =   420
         Width           =   9105
         Begin VB.ComboBox Cmb_Adm_Importacion_Empresa_Automatico 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   5700
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   690
            Left            =   7695
            TabIndex        =   3
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
            Left            =   8505
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Inicio_Automatico 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   111542275
            CurrentDate     =   39931
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Termino_Automatico 
            Height          =   315
            Left            =   5190
            TabIndex        =   5
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   111542275
            CurrentDate     =   39931
         End
         Begin VB.CommandButton Btn_Importacion_Automatica 
            Caption         =   "Importar"
            Height          =   690
            Left            =   7785
            Picture         =   "Frm_Adm_Importacion_Almacenes.frx":1BEE
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Termino"
            Height          =   195
            Left            =   3645
            TabIndex        =   8
            Top             =   660
            Width           =   1065
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
            TabIndex        =   7
            Top             =   300
            Width           =   735
         End
      End
      Begin zkemkeeperCtl.CZKEM CZKEM1 
         Height          =   135
         Left            =   120
         OleObjectBlob   =   "Frm_Adm_Importacion_Almacenes.frx":2178
         TabIndex        =   48
         Top             =   120
         Width           =   210
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTACION ACCESOS ALMACENES"
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
         Left            =   1050
         TabIndex        =   40
         Top             =   0
         Width           =   7125
      End
   End
   Begin MSComctlLib.ProgressBar Prg_Guardar 
      Height          =   690
      Left            =   1320
      TabIndex        =   43
      Top             =   7920
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
      Left            =   3360
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Prbar_Exportacion 
      Height          =   690
      Left            =   5175
      TabIndex        =   47
      Top             =   7920
      Visible         =   0   'False
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   1217
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Btn_Rpt_Generar 
      Caption         =   "Reporte"
      Height          =   690
      Left            =   2085
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":219C
      Style           =   1  'Graphical
      TabIndex        =   50
      Tag             =   "C"
      Top             =   7920
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   690
      Left            =   2085
      Picture         =   "Frm_Adm_Importacion_Almacenes.frx":2726
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "A"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "Frm_Adm_Importacion_Almacenes"
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
            Depurar_Lista
    Else
        MsgBox "Seleccione la empresa a la que se genera la asistencia", vbExclamation
        Cmb_Adm_Importacion_Empresa_Asistencia.SetFocus
    End If
End Sub


Private Sub Btn_Guardar_Click()
    Select Case Opcion
        Case "Importacion_Asistencias_Almacenes":
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
                    Mi_SQL = "SELECT COUNT(*) AS Detalles FROM Adm_Asistencias_Almacenes_Detalles"
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
        Case "Importacion_Asistencias_Almacenes":
            Grid_Importacion.Rows = 0
            Grid_Importacion_Lista_Depurada.Rows = 0
            Dtp_Importacion_Fecha_Inicio.Value = Now
            Dtp_Importacion_Fecha_Termino.Value = Now
            Txt_Adm_Importacion_Ruta_Archivo.Text = ""
            Txt_Importacion_Keri_Log.Text = ""
    End Select
End Sub

Private Sub Btn_Rpt_Generar_Click()
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
             Call Exportar_Excel_Bien(Ruta_Temporal & Opcion & "xls.txt", Ruta_Exportacion)
        End If
    Else
        MsgBox "No existe información para exportar", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
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

Private Sub Cmb_Adm_Importacion_Empresa_Manual_Click()
    If Cmb_Adm_Importacion_Empresa_Manual.ListIndex > -1 Then
        'Llena los checadores de la empresa seleccionada
        Call Conectar_Ayudante.Llena_Combo_Item("CEI.Equipo_ID, CAST(CEI.No_Equipo as varchar)+' '+CEI.Descripcion", _
            "Cat_Equipos_Almacenes_Identificadores CEI, Cat_Empresas_Equipos_Identificacion CEEI WHERE CEI.Equipo_ID = CEEI.Equipo_ID AND CEEI.Empresa_Id = '" & Format(Cmb_Adm_Importacion_Empresa_Manual.ItemData(Cmb_Adm_Importacion_Empresa_Manual.ListIndex), "00000") & "'", Cmb_Adm_Importacion_Checador, 0, "No_Equipo", , False, "TODAS")
        If Cmb_Adm_Importacion_Checador.ListCount > 0 Then
            Cmb_Adm_Importacion_Checador.ListIndex = 0
        End If
    End If
End Sub
Private Sub Cmb_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            If IsNumeric(Cmb_Empleado.Text) Then
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE No_Tarjeta='" & Trim(Cmb_Empleado.Text) & "'", Cmb_Empleado, 0, "No_Tarjeta")
            Else
                Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Empleado, 1, "Apellido_Paterno", " OR Nombre LIKE '%" & Trim(Cmb_Empleado.Text) & "%'" & _
                     " OR Apellido_Materno LIKE '%" & Trim(Cmb_Empleado.Text) & "%'", False, "")
            End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub
Public Sub Inicializa()
    Select Case Opcion
        Case "Importacion_Asistencias_Almacenes":
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
'*******************************************************************************
'NOMBRE_FUNCION: Alta_Importacion_Asistencias
'DESCRIPCION: Genera la lista de informacion del systema Keri-System
'PARAMETROS :
'CREO       : Flores Ramirez Yazmin
'FECHA_CREO : 07-Diciembre-2016
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Alta_Importacion_Asistencias()
Dim Rs_Alta_Adm_Asistencias_Detalles As rdoResultset     'Información de las asistencias
Dim Rs_Modifica_Adm_Asistencias_Detalles As rdoResultset     'Información de las asistencias
Dim Rs_Consulta_Informacion_Turnos As rdoResultset              'Informacion de los turnos
Dim Operacion As String                                         'Consecutivo del maximo del catalogo
Dim Cont_Fila As Integer                                        'Recorre el grid


On Error GoTo HANDLER:
    Me.MousePointer = 11
    If Grid_Importacion_Lista_Depurada.Rows > 0 Then
        Prg_Guardar.Value = 0
        Prg_Guardar.Max = Grid_Importacion_Lista_Depurada.Rows
        Prg_Guardar.Visible = True
    End If
    Conexion_Base.BeginTrans
    For Cont_Fila = 1 To Grid_Importacion_Lista_Depurada.Rows - 1
'        If Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 9) = "S" Then
            'Verifica si el registro ya se ha generado para actualizarlo, si no lo da de alta
            Mi_SQL = "SELECT * FROM Adm_Asistencias_Almacenes_Detalles "
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4)) & "'"
            Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " AND Hora_Entrada='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
            Set Rs_Modifica_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Modifica_Adm_Asistencias_Detalles.EOF Then
                'Cambia sólo si no está validada la asistencia
                    Mi_SQL = "UPDATE Adm_Asistencias_Almacenes_Detalles"
                    Mi_SQL = Mi_SQL & " SET Hora_Entrada='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4)) & "'"
                    Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy") & "'"
                    Conexion_Base.Execute Mi_SQL
            Else
                Mi_SQL = "INSERT INTO Adm_Asistencias_Almacenes_Detalles(Empleado_ID,No_Tarjeta,Fecha,Hora_Entrada,Validada)"
                Mi_SQL = Mi_SQL & " VALUES('" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 4)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Trim(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 0)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 2), "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion_Lista_Depurada.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , 'N')"
                Conexion_Base.Execute Mi_SQL
            End If
            Rs_Modifica_Adm_Asistencias_Detalles.Close
'        End If
        Prg_Guardar.Value = Prg_Guardar.Value + 1
        Me.Refresh
    Next
    Conexion_Base.CommitTrans
    Me.MousePointer = 0
    Prg_Guardar.Visible = False
    Grid_Importacion.Rows = 0
    Txt_Adm_Importacion_Ruta_Archivo.Text = ""
    MsgBox "Lista Registrada", vbInformation + vbOKOnly, Me.Caption
Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Me.MousePointer = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
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

'*******************************************************************************
'NOMBRE_FUNCION: Depurar_Lista
'DESCRIPCION: Depura la lista de información para obtener la hora de entrada y salida
'PARAMETROS :
'CREO       : Flores Ramirez Yazmin
'FECHA_CREO : 07-Diciembre-2016
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
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores_Almacenes AARC"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Asistencia.ItemData(Cmb_Adm_Importacion_Empresa_Asistencia.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Dtp_Asistencia_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Asistencia_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    If Cmb_Empleado.Text <> "" Then
        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID = '" & Format(Cmb_Empleado.ItemData(Cmb_Empleado.ListIndex), "00000") & "' "
    End If
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,CE.Apellido_Paterno,CE.Apellido_Materno,CE.Nombre,AARC.Hora"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    Call Encabezado_Reporte("IMPORTACION ASISTENCIAS ALMACENES", DateAdd("s", 1, Dtp_Asistencia_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Asistencia_Fecha_Termino.Value))
     
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
            Grid_Importacion_Lista_Depurada.AddItem "Tarjeta" & Chr(9) & "Empleado" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Comida" & Chr(9) & "Comida" & Chr(9) & "Salida" & Chr(9) & "Horas" & Chr(9) & "Registrado" & Chr(9) & "Empleado_ID" & Chr(9) & "Checador_ID"
                      '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
            Print #1, "No Nomina   Empleado                              Hora"
            Print #2, "No Nomina |Empleado|||Fecha|Hora"
            Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & "Iniciando depuración de información... (asignando hora de entrada, comida y salida)"
            'LLenado de la informacion
            Empleado_ID = ""
            While Not .EOF
                Prbar_Asistencia.Value = Prbar_Asistencia.Value + 1
                Fecha = .rdoColumns("Fecha")
                'valida el empleado
'                If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                    Empleado_ID = .rdoColumns("Empleado_ID")
                    Empleado_ID_Agregar = Empleado_ID
                    Nombre = .rdoColumns("Apellido_Paterno") & " " & .rdoColumns("Apellido_Materno") & " " & .rdoColumns("Nombre")
                    No_Tarjeta = .rdoColumns("No_Tarjeta")
'                    No_Checadas = 0
'                    Checador = .rdoColumns("Equipo_ID")
                    'Limpia las variables
                    Hora_Entrada = Format(.rdoColumns("Hora"), "HH:mm:ss")
'                    Hora_Salida = "0"
'                    Hora_Comida = "0"
'                    Hora_Comida2 = "0"
'                End If
'                .MoveNext
                If Not .EOF Then
'                    If Empleado_ID <> .rdoColumns("Empleado_ID") Then
                        'Agrega el registro
                        Grid_Importacion_Lista_Depurada.AddItem No_Tarjeta _
                            & Chr(9) & Nombre _
                            & Chr(9) & Format(Fecha, "dd/MMM/yyyy") _
                            & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                            & Chr(9) & Empleado_ID_Agregar
' _
'                            & Chr(9) & Format(Hora_Comida, "HH:mm:ss") _
'                            & Chr(9) & Format(Hora_Comida2, "HH:mm:ss") _
'                            & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
'                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
'                            & Chr(9) & Format(Round(Horas, 2), "#0.00") _
'                            & Chr(9) & "S" _

'                            & Chr(9) & Checador
                        Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); Spc(2); _
                            Mid(Nombre, 40); _
                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40)))
'                            ; _
'                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida, "HH:mm:ss"), 11); _
'                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Comida2, "HH:mm:ss"), 11); _
'                            Conectar_Ayudante.Alinea_Derecha(Format(Hora_Salida, "HH:mm:ss"), 9); _
'                            Conectar_Ayudante.Alinea_Derecha(CStr(Format(Round(Horas, 2), "#0.00")), 8)
                        Print #2, No_Tarjeta; "|"; Nombre; "|||"; _
                            Format(Fecha, "dd/MMM/yyyy"); "|"; _
                            Format(Hora_Entrada, "HH:mm:ss")
'                            ; "|"; _
'                            Format(Hora_Salida, "HH:mm:ss"); "|"; _
'                            Val(Horas); "|"; Checador
                        Me.Refresh
                        Empleado_ID = ""
'                    End If
'                Else
'
'                    'Agrega el registro
'                    Grid_Importacion_Lista_Depurada.AddItem No_Tarjeta _
'                        & Chr(9) & Nombre _
'                        & Chr(9) & Format(Fecha, "dd/MMM/yyyy") _
'                        & Chr(9) & Format(Hora_Entrada, "HH:mm:ss")
'
'                    Print #1, Conectar_Ayudante.Alinea_Derecha(No_Tarjeta, 10); _
'                        Spc(2); Mid(Nombre, 40); _
'                        Conectar_Ayudante.Alinea_Derecha(Format(Hora_Entrada, "HH:mm:ss"), 45 - Len(Mid(Nombre, 1, 40)))
'
'                    Print #2, No_Tarjeta; "|"; Nombre; "|||"; _
'                        Format(Fecha, "dd/MMM/yyyy"); "|"; _
'                        Format(Hora_Entrada, "HH:mm:ss")
'                     Me.Refresh
                End If
                .MoveNext
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
            .ColWidth(1) = 4500     'No Tarjeta
            .ColAlignment(1) = flexAlignLeftCenter
            .ColWidth(2) = 1200    'Empleado
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(3) = 800     'entrada
            .ColAlignment(3) = flexAlignCenterCenter
            .ColWidth(4) = 0       'comida
            .ColWidth(5) = 0       'comida
            .ColWidth(6) = 0     'salida
            .ColWidth(7) = 0     'horas
            .ColWidth(8) = 0       'horas
            .ColWidth(9) = 0       'Registrado
            .ColWidth(10) = 0      'Empleado_ID
            .ColWidth(11) = 0      'Checador
        End If
        If Grid_Importacion_Lista_Depurada.Rows > 0 Then
        Grid_Importacion_Lista_Depurada.Col = 1
        Grid_Importacion_Lista_Depurada.Sort = flexSortGenericAscending
        End If
        Call Finalizar_Reporte
    End With
    Me.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Obtiene_Informacion_Checadas(ByRef Ruta_Archivo As String, ByRef Mensaje As String)
'Genera los archivos de checadas
Dim dwEnrollNumber As Long
'Dim dwEnrollNumber As String
Dim dwEMachineNumber As Long
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
Dim Maquina_1 As Long
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
    Mi_SQL = "SELECT Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Empresas_Equipos_Identificacion.Equipo_ID,Cat_Equipos_Almacenes_Identificadores.No_Equipo"
    Mi_SQL = Mi_SQL & " FROM Cat_Empresas_Equipos_Identificacion,Cat_Equipos_Almacenes_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empresas_Equipos_Identificacion.Equipo_ID=Cat_Equipos_Almacenes_Identificadores.Equipo_ID"
    If Cmb_Adm_Importacion_Empresa_Automatico.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Cat_Empresas_Equipos_Identificacion.Empresa_ID='" & Format(Cmb_Adm_Importacion_Empresa_Automatico.ItemData(Cmb_Adm_Importacion_Empresa_Automatico.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Equipos_Almacenes_Identificadores.No_Equipo"
    Set Rs_Consulta_Dispositivos_Empresa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Dispositivos_Empresa.EOF Then
        With Rs_Consulta_Dispositivos_Empresa
            Set Rs_Alta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias_Registro_Checadores_Almacenes")
            While Not .EOF
                Txt_Importacion_Keri_Log.Text = ""
                'De acuerdo a los checadores de las empresas inicia la extraccion de informacion
                'Consulta la informacion para conectarse
                Mi_SQL = "SELECT Direccion_IP, Puerto_IP, No_Equipo, Descripcion "
                Mi_SQL = Mi_SQL & " FROM Cat_Equipos_Almacenes_Identificadores"
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
                        res = CZKEM1.GetDeviceStatus(Maquina_1, 6, dwvalue)
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
'                         If CZKEM1.ReadAllGLogData(Maquina) Then
                            Me.Refresh
                                Txt_Importacion_Keri_Log.Text = Trim(Txt_Importacion_Keri_Log.Text) & vbCrLf & _
                                    "Equipo listo, inicia generación de información ..."
                                Txt_Importacion_Keri_Log.SelStart = Len(Txt_Importacion_Keri_Log.Text)
                            CZKEM1.ReadAllUserID Maquina
                            Me.Refresh
                            Dim X As Long
                            X = 1
'                            While CZKEM1.GetGeneralLogData(1, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
'                            While CZKEM1.SSR_GetGeneralLogData(Maquina, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode)
'                            While CZKEM1.GetGeneralLogData(1, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode)
'                            While CZKEM1.GetGeneralLogData(x, Maquina_1, dwEnrollNumber, dwEMachineNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)

'                            While CZKEM1.GetGeneralLogData(X, Maquina_1, dwEnrollNumber, dwEMachineNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)
'                            While CZKEM1.GetGeneralLogDataStr(1, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
                            While CZKEM1.GetAllGLogData(X, Maquina_1, dwEnrollNumber, dwEMachineNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)
'                            While CZKEM1.SSR_GetGeneralLogData(1, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode)
'
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
                                                Mi_SQL = "SELECT * FROM Adm_Asistencias_Registro_Checadores_Almacenes AARC"
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
                                                    Mi_SQL = "INSERT INTO Adm_Asistencias_Registro_Checadores_Almacenes(No_Tarjeta,Fecha,Hora,Fecha_Importacion"
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
    Mi_SQL = "SELECT No_Equipo, Direccion_IP FROM Cat_Equipos_Almacenes_Identificadores WHERE Equipo_ID = '" & Format(Cmb_Adm_Importacion_Checador.ItemData(Cmb_Adm_Importacion_Checador.ListIndex), "00000") & "' "
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
            Mi_SQL = "SELECT No_Movimiento,No_Tarjeta FROM Adm_Asistencias_Registro_Checadores_Almacenes"
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
                Mi_SQL = "INSERT INTO Adm_Asistencias_Registro_Checadores_Almacenes "
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



