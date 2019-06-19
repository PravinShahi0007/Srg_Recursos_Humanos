VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Adm_Asistencias_Empleados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASISTENCIAS DE EMPLEADOS"
   ClientHeight    =   6930
   ClientLeft      =   3075
   ClientTop       =   2685
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   13425
   Begin VB.PictureBox Pic_Adm_Asistencias_Empleados_Consulta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   15
      ScaleHeight     =   3885
      ScaleWidth      =   6105
      TabIndex        =   11
      Top             =   15
      Width           =   6135
      Begin VB.Frame Fra_Rpt_Asistencia_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Asistencias Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3750
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   5910
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP 
            Height          =   315
            Left            =   975
            TabIndex        =   1
            Top             =   660
            Width           =   4770
         End
         Begin VB.OptionButton Opt_Horas_Reales 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horas Reales"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   8
            Top             =   2610
            Width           =   1335
         End
         Begin VB.OptionButton Opt_Horas_Aprobadas 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horas Aprobadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   975
            TabIndex        =   7
            Top             =   2610
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Periodo 
            Height          =   315
            ItemData        =   "Frm_Adm_Asistencias_Empleados.frx":0000
            Left            =   975
            List            =   "Frm_Adm_Asistencias_Empleados.frx":000D
            TabIndex        =   4
            Top             =   1770
            Width           =   4770
         End
         Begin VB.CommandButton Btn_Adm_Validacion_Horas_Generar 
            Caption         =   "Generar"
            Height          =   690
            Left            =   975
            Picture         =   "Frm_Adm_Asistencias_Empleados.frx":002E
            Style           =   1  'Graphical
            TabIndex        =   9
            Tag             =   "A"
            Top             =   2940
            Width           =   1200
         End
         Begin VB.CommandButton Btn_Salir 
            Caption         =   "Salir"
            Height          =   690
            Left            =   4515
            Picture         =   "Frm_Adm_Asistencias_Empleados.frx":05B8
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2940
            UseMaskColor    =   -1  'True
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Empresa 
            Height          =   315
            Left            =   975
            TabIndex        =   0
            Top             =   300
            Width           =   4770
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Empleado 
            Height          =   315
            Left            =   975
            TabIndex        =   3
            Top             =   1395
            Width           =   4770
         End
         Begin VB.ComboBox Cmb_Rpt_Asistencia_Empleados_Supervisor 
            Height          =   315
            Left            =   975
            TabIndex        =   2
            Top             =   1020
            Width           =   4770
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Asistencia_Empleados_Fecha_Termino 
            Height          =   315
            Left            =   3405
            TabIndex        =   6
            Top             =   2160
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   124059651
            CurrentDate     =   39872
         End
         Begin MSComCtl2.DTPicker Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio 
            Height          =   315
            Left            =   975
            TabIndex        =   5
            Top             =   2160
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd dd MMM yyyy"
            Format          =   124059651
            CurrentDate     =   39872
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Periodo"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   1830
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empresa"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empleado"
            Height          =   195
            Left            =   135
            TabIndex        =   16
            Top             =   1455
            Width           =   705
         End
         Begin VB.Label Lbl_Rango_Fechas 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fechas"
            Height          =   195
            Left            =   135
            TabIndex        =   15
            Top             =   2220
            Width           =   525
         End
         Begin VB.Label Lbl_Gerencia 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gerencia"
            Height          =   195
            Left            =   105
            TabIndex        =   14
            Top             =   735
            Width           =   645
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supervisor"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   1080
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox Pic_Adm_Asistencia_Empleados_Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6945
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   13395
      TabIndex        =   18
      Top             =   0
      Width           =   13425
      Begin VB.CommandButton Btn_SAP 
         Caption         =   "SAP"
         Enabled         =   0   'False
         Height          =   690
         Left            =   6066
         Picture         =   "Frm_Adm_Asistencias_Empleados.frx":0B42
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "A"
         Top             =   7080
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Exportar 
         Caption         =   "Exportar"
         Enabled         =   0   'False
         Height          =   690
         Left            =   3033
         Picture         =   "Frm_Adm_Asistencias_Empleados.frx":0C8C
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "A"
         Top             =   6165
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.Frame Fra_Validacion_Horas_Trabajo_Lista 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleado"
         Enabled         =   0   'False
         Height          =   6090
         Left            =   0
         TabIndex        =   24
         Top             =   15
         Width           =   13380
         Begin MSFlexGridLib.MSFlexGrid Grid_Asistencia_Empleados 
            Height          =   5550
            Left            =   90
            TabIndex        =   25
            Top             =   450
            Width           =   13245
            _ExtentX        =   23363
            _ExtentY        =   9790
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorFixed  =   -2147483628
            BackColorBkg    =   16777215
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Lbl_Periodo_Consulta 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Periodo"
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
            Left            =   5400
            TabIndex        =   30
            Top             =   225
            Width           =   660
         End
         Begin VB.Label Lbl_Validacion_Horas_Supervisor 
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
            Left            =   90
            TabIndex        =   26
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.CommandButton Btn_Regresar 
         Caption         =   "Regresar"
         Height          =   690
         Left            =   9099
         Picture         =   "Frm_Adm_Asistencias_Empleados.frx":1216
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6165
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Imprimir 
         Caption         =   "Imprimir"
         Height          =   690
         Left            =   0
         Picture         =   "Frm_Adm_Asistencias_Empleados.frx":17A0
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "A"
         Top             =   6165
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Salir_2 
         Caption         =   "Salir"
         Height          =   690
         Left            =   12135
         Picture         =   "Frm_Adm_Asistencias_Empleados.frx":21A2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6165
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   1260
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar Prbar_Exportacion 
         Height          =   165
         Left            =   4305
         TabIndex        =   28
         Top             =   6675
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Pbar_SAP 
         Height          =   165
         Left            =   7440
         TabIndex        =   31
         Top             =   7590
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Lbl_Progreso_SAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportando..."
         Height          =   195
         Left            =   7395
         TabIndex        =   32
         Top             =   7260
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Lbl_Progreso_Exportacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportando..."
         Height          =   195
         Left            =   4335
         TabIndex        =   29
         Top             =   6345
         Visible         =   0   'False
         Width           =   945
      End
   End
End
Attribute VB_Name = "Frm_Adm_Asistencias_Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios
Public Operacion As String
Dim Archivo_Reporte_Abierto As Boolean  'Indica si el archivo de reporte esta abierto
Dim Fecha_Inicio As Date        'Fecha de inicio del reporte
Dim Fecha_Termino As Date       'Fecha de termino del reporte
Public Sub Inicializar()
    Select Case Operacion
        Case "Asistencias_Empleados"
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Asistencia_Empleados_Empresa, 0, "Nombre", "", True, "TODAS")
            If Cmb_Rpt_Asistencia_Empleados_Empresa.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex = 0
            End If
            Call Conectar_Ayudante.Llena_Combo_Item("Gerencia_ID,Nombre", "Cat_Gerencias", Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP, 0, "Nombre", "", True, "TODAS")
            If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex = 0
            End If
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Empresa_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, 0, True, "TODOS")
            'If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
            '    Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
            'End If
            Cmb_Rpt_Asistencia_Empleados_Periodo.ListIndex = 0
            Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value = Now
            Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value = Now
            Cmb_Rpt_Asistencia_Empleados_Empresa.SetFocus
    End Select
End Sub

Private Sub Btn_Adm_Validacion_Horas_Generar_Click()
    Fecha_Inicio = Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value
    Fecha_Termino = Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value
    Generar_Reporte_Asistencia_Empleados
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
    Cmd_Exportar.FileName = Operacion & ".xls"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
        Call Exportar_Excel(Ruta_Temporal & Operacion & "xls.txt", Ruta_Exportacion, Prbar_Exportacion, Lbl_Progreso_Exportacion, Me)
    End If
Exit Sub
HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Imprimir_Click()
Dim linea As String 'Obtiene el texto a imprimir
Dim X As Printer
Dim contar_linea As Integer

On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    Printer.Orientation = vbPRORLandscape
    Printer.FontSize = 7
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 11
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 7
    Printer.Font = "Courier New"
    Open Ruta_Temporal & Operacion & ".txt" For Input As #1
    Do While Not EOF(1)
        contar_linea = contar_linea + 1
        If contar_linea = 90 Then
            Printer.NewPage
        End If
        Line Input #1, linea
        Printer.Print linea
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
End Sub


Private Sub Btn_Regresar_Click()
    Me.Height = 4335
    Me.Width = 6225
    Pic_Adm_Asistencia_Empleados_Lista.Visible = False
    Pic_Adm_Asistencias_Empleados_Consulta.Visible = True
End Sub

Private Sub Btn_Salir_2_Click()
    Unload Me
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Btn_SAP_Click()
Dim RutaArchivo As String
Dim Nombre_Archivo() As String
Dim Nombre_Archivo_SAP As String

On Error GoTo HANDLER
    'Guarda una copia del archivo
    'Nombre del archivo
    Nombre_Archivo_SAP = "P"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "Q"       'P Produccion
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "1"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "545_"    '544 Produccion
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & Format(Now, "yyyyMMddHHmmss") & "_"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "MX"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "FRC3_"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "TIME01_"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "D"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "UT8"
    Nombre_Archivo_SAP = Nombre_Archivo_SAP & "G2I"
    'Genera el archivo de SAP
    MDIFrm_Apl_Principal.MousePointer = 11
    'Call Generar_Reporte_SAP(Nombre_Archivo(0))
    Call Generar_Reporte_SAP(Nombre_Archivo_SAP)
    MDIFrm_Apl_Principal.MousePointer = 0
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Texto |*.TXT|"     '.SAP Produccion
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    MDIFrm_Apl_Principal.CommonDialog1.FileName = Nombre_Archivo_SAP
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    'Nombre_Archivo = Split(MDIFrm_Apl_Principal.CommonDialog1.FileTitle, ".")
    RutaArchivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    SHCopyFile Ruta_Temporal & "SAP" & ".txt", RutaArchivo
    MsgBox "El Archivo ha sido Guardado en " & RutaArchivo, vbInformation
Exit Sub
HANDLER:
    MsgBox "Ocurrio un error al copiar el archivo, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Gerencia_ID, Nombre", "Cat_Gerencias", Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP, 1, "Nombre", "", True, "TODAS")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then
            'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, 0, True, "TODOS")
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%') AND SUpervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 1 Then
                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 1
            End If
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_Paterno", "AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Empleado.Text) & "%')", False, "")
            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 1 Then
                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 1
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empleado_KeyUp(KeyCode As Integer, Shift As Integer)
        Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Empleado, KeyCode)
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_Click()
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > -1 Then
        If Trim(Empleado_Supervisor_ID) = "" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A'", Cmb_Rpt_Asistencia_Empleados_Supervisor, 0, "Apellido_paterno", "", True, "TODOS")
            Cmb_Rpt_Asistencia_Empleados_Supervisor.Enabled = True
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A' AND Empleado_ID='" & Empleado_Supervisor_ID & "'", Cmb_Rpt_Asistencia_Empleados_Supervisor, 0, "Apellido_Paterno", "")
            Cmb_Rpt_Asistencia_Empleados_Supervisor.Enabled = False
        End If
        If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListCount > 0 Then Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Rpt_Asistencia_Empleados_Empresa, 1, "Nombre", True, "TODAS")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Empresa, KeyCode)
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Periodo_Click()
Dim Dias_Mes As Integer
    Select Case Cmb_Rpt_Asistencia_Empleados_Periodo.Text
        Case "ACUMULADO":
            
        Case "SEMANAL":
            Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value = DateAdd("d", 6, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value)
        Case "MENSUAL":
            Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value = DateAdd("d", 30, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value)
    End Select
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_Click()
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados ", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_paterno", "AND Estatus = 'A' AND Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
        If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
            Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
        End If
    Else
        If Trim(Empleado_Supervisor_ID) = "" Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_paterno", "AND Estatus = 'A' ", True, "TODOS")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados", Cmb_Rpt_Asistencia_Empleados_Empleado, 1, "Apellido_paterno", "AND Estatus = 'A' AND Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'", True, "TODOS")
            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
            End If
        End If
        If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
            Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
        End If
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text) & "%')", Cmb_Rpt_Asistencia_Empleados_Supervisor, 0, "Apellido_Paterno", , True, "TODOS")
        If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListCount > 1 Then
            Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex = 1
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Supervisor_ID = '" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "' AND Estatus = 'A'", Cmb_Rpt_Asistencia_Empleados_Empleado, 0, "Apellido_paterno", , True, "TODOS")
            If Cmb_Rpt_Asistencia_Empleados_Empleado.ListCount > 0 Then
                Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex = 0
            End If
        End If
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Rpt_Asistencia_Empleados_Supervisor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Rpt_Asistencia_Empleados_Supervisor, KeyCode)
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_Asistencia_Empleados
'DESCRIPCION: Genera el reporte de asistencia de empleados
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 18-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_Asistencia_Empleados()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset 'Informacion de los tiempo muertos
Dim Rs_Consulta_Empleados As rdoResultset
Dim Mi_SQL As String                            'Cadena de la consulta del reporte
Dim Empresa_ID_Reporte As String                'ID de la empresa
Dim Empleado_ID_Reporte As String               'ID de la empleado
Dim Supervisor_ID As String                     'ID del supervisor
Dim Nombre_Supervisor As String                 'Nombre del supervisor
Dim Nombre_Empresa As String                    'Nombre de la empresa
Dim Cadena_Grid As String                       'Define la informacion del encabezado
Dim Cadena_Grid_2 As String                     'Define la informacion del encabezado
Dim Cadena_Grid_Excel As String                 'Define l ainformacion para la exportacion a excel
Dim Cadena_Grid_Excel_2 As String               'Define l ainformacion para la exportacion a excel
Dim Cadena_Firmas As String                     'Define l ainformacion para la cadena de firmas
Dim Cadena_Firmas_2 As String                   'Define l ainformacion para la cadena de firmas
Dim Cadena_Firmas_Excel As String                     'Define l ainformacion para la cadena de firmas
Dim Cadena_Firmas_Excel_2 As String                   'Define l ainformacion para la cadena de firmas
Dim Cont_Fila As Integer                        'recorre los dias
Dim Encontrado As Boolean                       'Define si el registro se encontro
Dim Fila_Encontrado As Integer                  'Define la fila donde se encontro el registro
Dim Cont_Col As Integer                         'Contador de columnas del grid
Dim Fecha_Encontrada As Boolean                 'Define si la fecha se encontro en el grid
Dim Fecha As Date                               'Indica la fecha del registro
Dim Observaciones As String                     'Mantiene las observaciones de las incidencias
Dim Horas As String                             'Define las horas aprobadas del empleado
Dim Valor As Double                             'Guarda el valor de la celda del grid
Dim No_Empleado As String                       'No de empleado NOI

On Error GoTo HANDLER
    Grid_Asistencia_Empleados.Rows = 0
    Empresa_ID_Reporte = ""
    Empleado_ID_Reporte = ""
    Supervisor_ID = ""
    Nombre_Supervisor = ""
    'Consulta
    Mi_SQL = "SELECT ISNULL(AA.Horas_Aprobadas,0) AS Horas,ISNULL(AA.Horas_Extra,0) AS Horas_Extra"
    Mi_SQL = Mi_SQL & " ,(CAST((CAST(datediff(n,ISNULL(AA.Hora_Entrada,0),ISNULL(AA.Hora_Salida,0)) AS Decimal(18,2))/60) AS decimal(18,2)) - cast((casT(datediff(n,ISNULL(AA.Hora_Entrada_Comida,0),ISNULL(AA.Hora_Salida_Comida,0)) as Decimal(18,2))/60) AS decimal(18,2))) as Horas_Reales"
    Mi_SQL = Mi_SQL & " ,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre,CEM.Empresa_ID,CEM.Nombre AS Nombre_Empresa"
    Mi_SQL = Mi_SQL & " ,AA.Fecha,CE.Empleado_ID,AA.Simbologia,AA.Referencia"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Supervisor_ID,'N') AS Supervisor_ID,ISNULL(AA.Referencia,'') AS No_Movimiento,ISNULL(Tipo_Incidencia,'') AS Tipo_Incidencia"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA,Cat_Empleados CE,Cat_Empresas CEM"
    Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID=CE.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID=CEM.Empresa_ID"
    'Validacion de Empresa
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    'Validacion de Gerencia
    If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Gerencia_UAP='" & Format(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ItemData(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    'Validacion de Empleado
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AA.Empleado_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    'Rango de Fechas
    Mi_SQL = Mi_SQL & " AND AA.Fecha BETWEEN '" & Format(Fecha_Inicio, "MM/dd/yyyy") & "' AND '" & Format(Fecha_Termino, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY CEM.Empresa_ID,CE.Supervisor_ID,CE.No_Tarjeta,AA.Fecha"
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            'Prepara el grid para agregar la información
            Grid_Asistencia_Empleados.Cols = (DateDiff("d", Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value) * 3) + 11
            Cadena_Grid = "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
            Cadena_Grid_2 = "Empresa_ID" & Chr(9) & "Supervisor_ID" & Chr(9) & "Empleado_ID" & Chr(9) & "Depto" & Chr(9) & "Departamento" & Chr(9) & "Empresa/Supervisor"
            For Cont_Fila = 0 To DateDiff("d", Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value)
                Fecha = Format(DateAdd("d", Cont_Fila, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), "MM/dd/yyyy")
                Cadena_Grid = Cadena_Grid & Chr(9) & Fecha & Chr(9) & Fecha & Chr(9) & Fecha
                Cadena_Grid_2 = Cadena_Grid_2 & Chr(9) & "Asist." & Chr(9) & "Hrs." & Chr(9) & "Hrs.Extra"
            Next
            Grid_Asistencia_Empleados.AddItem Cadena_Grid
            Grid_Asistencia_Empleados.AddItem Cadena_Grid_2 & Chr(9) & "Total Hrs." & Chr(9) & "Total Hrs. Extra"
            While Not .EOF
                'Agrega un registro por la empresa
                If Empresa_ID_Reporte <> .rdoColumns("Nombre_Empresa") Then
                    Empresa_ID_Reporte = .rdoColumns("Nombre_Empresa")
                    Grid_Asistencia_Empleados.AddItem .rdoColumns("Empresa_ID") & Chr(9) & .rdoColumns("Supervisor_ID") & Chr(9) & "" & Chr(9) & "" & Chr(9) & "." & Chr(9) & .rdoColumns("Nombre_Empresa") & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                    For Cont_Col = 0 To Grid_Asistencia_Empleados.Cols - 1
                        Grid_Asistencia_Empleados.Col = Cont_Col
                        Grid_Asistencia_Empleados.Row = Grid_Asistencia_Empleados.Rows - 1
                        Grid_Asistencia_Empleados.CellBackColor = &H8000000A
                    Next
                    Supervisor_ID = ""
                End If
                Debug.Print Empleado_ID_Reporte
                Encontrado = False
                Fila_Encontrado = 0
                If .rdoColumns("Empleado_ID") <> Empleado_ID_Reporte Then
                    Empleado_ID_Reporte = .rdoColumns("Empleado_ID")
                    'Supervisor_ID = ""
                End If
                If .rdoColumns("Supervisor_ID") <> Supervisor_ID Then
                    Supervisor_ID = .rdoColumns("Supervisor_ID")
                    Mi_SQL = "SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre"
                    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE"
                    Mi_SQL = Mi_SQL & " WHERE CE.Empleado_ID = '" & Supervisor_ID & "'"
                    Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Nombre")
                    Grid_Asistencia_Empleados.AddItem .rdoColumns("Empresa_ID") & Chr(9) & Supervisor_ID & Chr(9) & "" & Chr(9) & "" & Chr(9) & "-" & Chr(9) & Nombre_Supervisor & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                    For Cont_Col = 0 To Grid_Asistencia_Empleados.Cols - 1
                        Grid_Asistencia_Empleados.Col = Cont_Col
                        Grid_Asistencia_Empleados.Row = Grid_Asistencia_Empleados.Rows - 1
                        Grid_Asistencia_Empleados.CellBackColor = &H80000016
                    Next
                End If
                'busca si el registro ya se ha agregado
                For Cont_Fila = 2 To Grid_Asistencia_Empleados.Rows - 1
                    If Grid_Asistencia_Empleados.TextMatrix(Cont_Fila, 2) = Empleado_ID_Reporte Then
                        Encontrado = True
                        Fila_Encontrado = Cont_Fila
                        Exit For
                    End If
                Next
                If Encontrado = False Then
                    Grid_Asistencia_Empleados.AddItem .rdoColumns("Empresa_ID") & Chr(9) & Supervisor_ID & Chr(9) & .rdoColumns("Empleado_ID") & Chr(9) & "" & Chr(9) & "" & Chr(9) & .rdoColumns("Nombre")
                    Fila_Encontrado = Grid_Asistencia_Empleados.Rows - 1
                End If
                'Busca la columna donde insertara la fecha
                Fecha_Encontrada = False
                For Cont_Col = 0 To Grid_Asistencia_Empleados.Cols - 2
                    If Cont_Col >= 6 Then
                        If DateDiff("d", Format(.rdoColumns("Fecha"), "MM/dd/yyyy"), Format(Grid_Asistencia_Empleados.TextMatrix(0, Cont_Col), "MM/dd/yyyy")) = 0 And Fecha_Encontrada = False Then
                            Observaciones = ""
                            Horas = ""
'                            If Not IsNull(.rdoColumns("Referencia")) Then
'                                Mi_SQL = "SELECT No_Movimiento, Observaciones FROM Adm_Movimientos"
'                                Mi_SQL = Mi_SQL & " WHERE No_Movimiento = '" & Trim(.rdoColumns("Referencia")) & "'"
'                                Mi_SQL = Mi_SQL & " AND Tipo_Incidencia = '" & Trim(.rdoColumns("Tipo_Incidencia")) & "'"
'                                Observaciones = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Observaciones")
'                            End If
                            If Opt_Horas_Aprobadas.Value = True Then
                                Horas = CStr(.rdoColumns("Horas"))
'                                If Val(.rdoColumns("Horas_Reales")) <> Val(.rdoColumns("Horas")) Then
'                                    Horas = Horas & "/"
'                                End If
                            Else
                                If Opt_Horas_Reales.Value = True Then Horas = CStr(.rdoColumns("Horas_Reales"))
                            End If
                            Grid_Asistencia_Empleados.TextMatrix(Fila_Encontrado, Cont_Col) = CStr(.rdoColumns("Simbologia")) & "                  " & .rdoColumns("Tipo_Incidencia") & ":" & .rdoColumns("No_Movimiento")
                            Grid_Asistencia_Empleados.TextMatrix(Fila_Encontrado, Cont_Col + 1) = Horas
                            Grid_Asistencia_Empleados.TextMatrix(Fila_Encontrado, Cont_Col + 2) = .rdoColumns("Horas_Extra") ' Observaciones
                            Fecha_Encontrada = True
                            Exit For
                        End If
                    End If
                Next
                .MoveNext
            Wend
            'Agrega la fila de totales
            Grid_Asistencia_Empleados.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTALES"
            'da formato al grid
            With Grid_Asistencia_Empleados
                .FixedCols = 6
                .FixedRows = 3
                .ColWidth(0) = 0
                .ColWidth(1) = 0
                .ColWidth(2) = 0
                .ColWidth(3) = 250
                .ColWidth(4) = 250
                .ColWidth(5) = 4000
                For Cont_Col = 6 To Grid_Asistencia_Empleados.Cols - 1
                    .ColAlignment(Cont_Col) = flexAlignLeftCenter
                Next
                'Agrega los totales por dia y por empleado
                For Cont_Fila = 4 To .Rows - 2
                    Debug.Print Cont_Fila
                    For Cont_Col = 6 To .Cols - 1
                        If .TextMatrix(1, Cont_Col) = "Hrs." Then
                            'Total de Horas por dia
                            If InStr(1, .TextMatrix(Cont_Fila, Cont_Col), "/") > 0 Then
                                Valor = Mid(.TextMatrix(Cont_Fila, Cont_Col), 1, Len(.TextMatrix(Cont_Fila, Cont_Col)) - 1)
                            Else
                                Valor = Val(.TextMatrix(Cont_Fila, Cont_Col))
                            End If
                            .TextMatrix(.Rows - 1, Cont_Col) = Val(.TextMatrix(.Rows - 1, Cont_Col)) + Val(Valor)
                            .TextMatrix(Cont_Fila, .Cols - 2) = Val(.TextMatrix(Cont_Fila, .Cols - 2)) + Val(Valor)
                        End If
                        If .TextMatrix(1, Cont_Col) = "Hrs.Extra" Then
                            'Total de Horas por dia
                            If InStr(1, .TextMatrix(Cont_Fila, Cont_Col), "/") > 0 Then
                                Valor = Mid(.TextMatrix(Cont_Fila, Cont_Col), 1, Len(.TextMatrix(Cont_Fila, Cont_Col)) - 1)
                            Else
                                Valor = Val(.TextMatrix(Cont_Fila, Cont_Col))
                            End If
                            .TextMatrix(.Rows - 1, Cont_Col) = Val(.TextMatrix(.Rows - 1, Cont_Col)) + Val(Valor)
                            .TextMatrix(Cont_Fila, .Cols - 1) = Val(.TextMatrix(Cont_Fila, .Cols - 1)) + Val(Valor)
                        End If
                    Next
                Next
            End With
            'Agrega el encabezado al reporte
            'Genera el archivo de reportes
            Call Encabezado_Reporte("REPORTE DE ASISTENCIAS DE EMPLEADO", DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value), DateAdd("s", 1, Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value))
            With Grid_Asistencia_Empleados
                Cadena_Grid = Conectar_Ayudante.Agregar_Espacios("Nombre", 25)
                Cadena_Grid_2 = Conectar_Ayudante.Agregar_Espacios("", 25)
                Cadena_Grid_Excel = "No.Empleado|Departamento|Turno|MO|Nombre|Turno|GerenciaUAP|Area|Supervisor|TC|Depto|TN|Ingreso|Secc"
                For Cont_Col = 6 To .Cols - 4 Step 3
                    Cadena_Grid = Cadena_Grid & "          " & Format(.TextMatrix(0, Cont_Col), "dd") & "                   "
                    Cadena_Grid_2 = Cadena_Grid_2 & "Asist." & "   " & "Hrs." & "   " & "Hrs. Extra"
                    Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & Format(.TextMatrix(0, Cont_Col), "dd")
                Next
                Print #1, Cadena_Grid
                Print #1, Cadena_Grid_2
                Print #1, "--------------------------------------------------------------------------------------------------------------------------"
                Print #2, Cadena_Grid_Excel
                Cadena_Grid = ""
                Supervisor_ID = ""
                Empresa_ID_Reporte = ""
                For Cont_Fila = 2 To .Rows - 2
                    If Empresa_ID_Reporte <> Trim(.TextMatrix(Cont_Fila, 0)) Then
                        Empresa_ID_Reporte = Trim(.TextMatrix(Cont_Fila, 0))
                        Nombre_Empresa = Conectar_Ayudante.Busca_Dato_BD("SELECT Nombre FROM Cat_Empresas CE WHERE CE.Empresa_ID = '" & Empresa_ID_Reporte & "'", "Nombre")
                        Print #1,
                        Print #1, Nombre_Empresa
                        Cont_Fila = Cont_Fila + 1
                        Supervisor_ID = ""
                    End If
                    If Supervisor_ID <> Trim(.TextMatrix(Cont_Fila, 1)) Then
                        Supervisor_ID = Trim(.TextMatrix(Cont_Fila, 1))
                        Nombre_Supervisor = Conectar_Ayudante.Busca_Dato_BD("SELECT (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre_Supervisor FROM Cat_Empleados CE WHERE CE.Empleado_ID = '" & Supervisor_ID & "'", "Nombre_Supervisor")
                        If Supervisor_ID = "N" Then
                            Nombre_Supervisor = "Supervisor no asignado"
                        End If
                        Print #1,
                        Print #1, "Supervisor: " & Nombre_Supervisor
                        Cont_Fila = Cont_Fila + 1
                    End If
                    Cadena_Grid = Trim(Mid(.TextMatrix(Cont_Fila, 5), 1, 25))
                    Cadena_Grid = Conectar_Ayudante.Agregar_Espacios(Cadena_Grid, 25) & "  "
                    Cadena_Grid_Excel = ""
                    For Cont_Col = 6 To .Cols - 4 Step 3
                        Cadena_Grid = Cadena_Grid & .TextMatrix(Cont_Fila, Cont_Col) & "     " & _
                                      .TextMatrix(Cont_Fila, Cont_Col + 1) & Conectar_Ayudante.Alinea_Derecha("", 7 - Len(.TextMatrix(Cont_Fila, Cont_Col + 1))) & _
                                      Left(.TextMatrix(Cont_Fila, Cont_Col + 2), 14) & Conectar_Ayudante.Alinea_Derecha("", 17 - Len(Left(.TextMatrix(Cont_Fila, Cont_Col + 2), 14)))
                        If Mid(.TextMatrix(Cont_Fila, Cont_Col), 1, 1) = "A" Then
                            If .TextMatrix(Cont_Fila, Cont_Col + 2) > 0 Then
                                Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & .TextMatrix(Cont_Fila, Cont_Col + 2)
                            Else
                                Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & "A"
                            End If
                        Else
                            If Trim(.TextMatrix(Cont_Fila, Cont_Col)) <> "" Then
                                Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & Trim(Mid(.TextMatrix(Cont_Fila, Cont_Col), 1, 5))
                            Else
                                Cadena_Grid_Excel = Cadena_Grid_Excel & "|" & ""
                            End If
                        End If
                    Next
                    'Consulta el número de empleado
                    Mi_SQL = "SELECT Cat_Empleados.Empleado_ID,Cat_Empleados.No_Tarjeta,Cat_Turnos.Nombre AS Turno,Cat_Gaps.Nombre AS Area,Cat_Empleados.Fecha_Ingreso,Cat_Empleados.Tipo_Contratacion,Cat_Empleados.Tipo_Empleado,Cat_Departamentos.Nombre AS Departamento,Cat_Departamentos.Clave,Cat_Empleados.Turno_ID,Cat_Empleados.Nomipaq_ID,ISNULL(Cat_Gerencias.Nombre,'') AS Gerencia_UAP"
                    Mi_SQL = Mi_SQL & " FROM Cat_Empleados INNER JOIN Cat_Turnos ON Cat_Empleados.Turno_ID=Cat_Turnos.Turno_ID"
                    Mi_SQL = Mi_SQL & " INNER JOIN Cat_Departamentos ON Cat_Empleados.Departamento_ID=Cat_Departamentos.Departamento_ID"
                    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Gaps ON Cat_Empleados.Gap_ID=Cat_Gaps.Gap_ID"
                    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Gerencias ON Cat_Empleados.Gerencia_UAP=Cat_Gerencias.Gerencia_ID"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Empleado_ID='" & .TextMatrix(Cont_Fila, 2) & "'"
                    Set Rs_Consulta_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Empleados.EOF Then
                        Print #1, Rs_Consulta_Empleados.rdoColumns("No_Tarjeta") & "|" & Cadena_Grid
                        Print #2, Rs_Consulta_Empleados.rdoColumns("No_Tarjeta") _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Departamento") _
                            & "|" & Val(Rs_Consulta_Empleados.rdoColumns("Turno_ID")) _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Tipo_Empleado") _
                            & "|" & .TextMatrix(Cont_Fila, 5) _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Turno") _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Gerencia_UAP") _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Area") _
                            & "|" & Nombre_Supervisor _
                            & "|" & Mid(Rs_Consulta_Empleados.rdoColumns("Tipo_Contratacion"), 1, 1) _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Clave") & "|" & "" _
                            & "|" & Format(Rs_Consulta_Empleados.rdoColumns("Fecha_Ingreso"), "MM/dd/yyyy") _
                            & "|" & Rs_Consulta_Empleados.rdoColumns("Nomipaq_ID") _
                            & Cadena_Grid_Excel
                    End If
                    Rs_Consulta_Empleados.Close
                Next
            End With
            Print #1,
            Print #1,
            Print #2,
            Print #2,
            .Close
            Call Finalizar_Reporte(True)
            Btn_Imprimir.Enabled = True
            Btn_Exportar.Enabled = True
            Btn_SAP.Enabled = True
            Btn_Regresar.Enabled = True
            Btn_Salir.Enabled = True
            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados", Me)
        End If
        If Grid_Asistencia_Empleados.Rows > 1 Then
            Pic_Adm_Asistencia_Empleados_Lista.Visible = True
            Pic_Adm_Asistencias_Empleados_Consulta.Visible = False
            Me.Height = 7410
            Me.Width = 13515
            Fra_Validacion_Horas_Trabajo_Lista.Enabled = True
            Lbl_Validacion_Horas_Supervisor.Caption = Trim(Cmb_Rpt_Asistencia_Empleados_Supervisor.Text)
            Lbl_Periodo_Consulta.Caption = "Periodo: " & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "dd MMM yyyy") & " al " & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "dd MMM yyyy")
            Collapsing = True
            Call Collapse_Grid
            Collapsing = False
        Else
            MsgBox "No existe informacion con los parametros seleccionados", vbInformation + vbOKOnly, Me.Caption
        End If
    End With
    Set Rs_Consulta_Adm_Asistencias = Nothing
    'Haya o no haya registros se cambia el Puntero del Mouse
    MDIFrm_Apl_Principal.MousePointer = 0
    Me.Left = 0
    Me.Top = 0
Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    If Archivo_Reporte_Abierto = True Then
        Close #1, #2
    End If
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Reporte_SAP
'DESCRIPCION: Genera el reporte de asistencia de empleados para SAP
'PARAMETROS : Nombre_Archivo- Es el nombre del archivo
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 12-Octubre-2013
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Reporte_SAP(Nombre_Archivo As String)
Dim Rs_Consulta_Datos As rdoResultset
Dim Contador_Registos As Long
Dim Referencia_Incapacidad As String

On Error GoTo HANDLER
    'Genera el archivo para SAP
    Open Ruta_Temporal & "SAP" & ".txt" For Output As #3
    'Encabezado
    Contador_Registos = Contador_Registos + 1
    Print #3, "HEADR|" & Chr(34) & "GVESRG|" & Chr(34) & "SRG|" & Chr(34) & "Fátima Fugueroa|" & Chr(34) & "|" & Chr(34) & "|" & Chr(34) & Nombre_Archivo & "|" & Chr(34) & Format(Now, "yyyyMMdd") & "|" & Chr(34) & Format(Now, "HHmmss") & "|" & Chr(34) & "P|" & Chr(34) & "01|" & Chr(34) & "|" & Chr(34) & "|" & Chr(34) & "|" & Chr(34) & ""
    'Consulta si hay cambios de turnos previstos para el rango de fechas seleccionado
    Mi_SQL = "SELECT Adm_Cambios_Turnos.Empleado_ID,Adm_Cambios_Turnos.Turno_Nuevo_ID,ISNULL(Cat_Empleados.Clave_SAP,'') AS SAP_Empleado,Adm_Cambios_Turnos.Fecha_Cambio,ISNULL(Cat_Turnos.Clave_SAP,'') AS SAP_Turno"
    Mi_SQL = Mi_SQL & " FROM Adm_Cambios_Turnos,Cat_Empleados,Cat_Turnos,Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Adm_Cambios_Turnos.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Cambios_Turnos.Turno_Nuevo_ID=Cat_Turnos.Turno_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID=Cat_Empresas.Empresa_ID"
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then          'Filtro de Empresa
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex > 0 Then     'Filtro de Gerencia
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Gerencia_UAP='" & Format(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ItemData(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then       'Filtro de Supervisor
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then         'Filtro de Empleado
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Cambios_Turnos.Fecha_Cambio BETWEEN '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Set Rs_Consulta_Datos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Datos.EOF
        Contador_Registos = Contador_Registos + 1
        Print #3, "P0007" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Empleado") _
            & "|" & Chr(34) & "MX" _
            & "|" & Chr(34) & "001" _
            & "|" & Chr(34) & "INS" _
            & "|" & Chr(34) & "0007" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Format(Rs_Consulta_Datos.rdoColumns("Fecha_Cambio"), "yyyyMMdd") _
            & "|" & Chr(34) & "99991231" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Turno") _
            & "|" & Chr(34) & "1" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & ""
        Rs_Consulta_Datos.MoveNext
    Wend
    Rs_Consulta_Datos.Close
    'Consulta las incidencias del periodo
    Mi_SQL = "SELECT Adm_Movimientos_Asistencias.Empleado_ID,Adm_Movimientos_Asistencias.Tipo_Falta_ID,ISNULL(Cat_Empleados.Clave_SAP,'') AS SAP_Empleado,Adm_Movimientos_Asistencias.Fecha_Inicio,Adm_Movimientos_Asistencias.Fecha_Termino,ISNULL(Cat_Tipos_Faltas.Clave_SAP,'') AS SAP_Incidencia,Adm_Movimientos_Asistencias.Observaciones"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias,Cat_Empleados,Cat_Tipos_Faltas,Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Adm_Movimientos_Asistencias.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos_Asistencias.Tipo_Falta_ID=Cat_Tipos_Faltas.Tipo_Falta_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID=Cat_Empresas.Empresa_ID"
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then          'Filtro de Empresa
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex > 0 Then     'Filtro de Gerencia
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Gerencia_UAP='" & Format(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ItemData(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then       'Filtro de Supervisor
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then         'Filtro de Empleado
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Movimientos_Asistencias.Fecha_Inicio BETWEEN '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " UNION ALL"   'Une los registros de falta injustificada
    Mi_SQL = Mi_SQL & " SELECT Adm_Asistencias.Empleado_ID,'' AS Tipo_Falta_ID,ISNULL(Cat_Empleados.Clave_SAP,'') AS SAP_Empleado,Adm_Asistencias.Fecha AS Fecha_Inicio,Adm_Asistencias.Fecha AS Fecha_Termino,'3200' AS SAP_Incidencia,'' AS Observaciones"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias,Cat_Empleados,Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID=Cat_Empresas.Empresa_ID"
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then          'Filtro de Empresa
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex > 0 Then     'Filtro de Gerencia
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Gerencia_UAP='" & Format(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ItemData(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then       'Filtro de Supervisor
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then         'Filtro de Empleado
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Fecha BETWEEN '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Simbologia='F'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Inicio"
    Set Rs_Consulta_Datos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Datos.EOF
        'Valida los tipos para enviar la referencia de incapacidad
        If Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") = "2000" _
        Or Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") = "2100" _
        Or Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") = "2200" Then
           Referencia_Incapacidad = Rs_Consulta_Datos.rdoColumns("Observaciones")
           'Referencia_Incapacidad = "10" & Contador_Registos
        Else
           Referencia_Incapacidad = ""
        End If
        'Agrega el registro
        Contador_Registos = Contador_Registos + 1
        Print #3, "P2001" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Empleado") _
            & "|" & Chr(34) & "MX" _
            & "|" & Chr(34) & "001" _
            & "|" & Chr(34) & "INS" _
            & "|" & Chr(34) & "2001" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") _
            & "|" & Chr(34) & Format(Rs_Consulta_Datos.rdoColumns("Fecha_Inicio"), "yyyyMMdd") _
            & "|" & Chr(34) & Format(Rs_Consulta_Datos.rdoColumns("Fecha_Termino"), "yyyyMMdd") _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" & "|" & Chr(34) & "" & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Referencia_Incapacidad
        Rs_Consulta_Datos.MoveNext
    Wend
    Rs_Consulta_Datos.Close
    'Consulta las horas extra del periodo
    Mi_SQL = "SELECT Adm_Asistencias.Empleado_ID,'' AS Tipo_Falta_ID,ISNULL(Cat_Empleados.Clave_SAP,'') AS SAP_Empleado,Adm_Asistencias.Fecha,Adm_Asistencias.Horas_Extra,'9110' AS SAP_Incidencia"
    Mi_SQL = Mi_SQL & " FROM Adm_Asistencias,Cat_Empleados,Cat_Empresas"
    Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID=Cat_Empresas.Empresa_ID"
    If Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex > 0 Then          'Filtro de Empresa
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empresa_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empresa.ItemData(Cmb_Rpt_Asistencia_Empleados_Empresa.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex > 0 Then     'Filtro de Gerencia
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Gerencia_UAP='" & Format(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ItemData(Cmb_Rpt_Asistencia_Empleados_Gerencia_UAP.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex > 0 Then       'Filtro de Supervisor
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Supervisor_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Supervisor.ItemData(Cmb_Rpt_Asistencia_Empleados_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex > 0 Then         'Filtro de Empleado
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID='" & Format(Cmb_Rpt_Asistencia_Empleados_Empleado.ItemData(Cmb_Rpt_Asistencia_Empleados_Empleado.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Fecha BETWEEN '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio.Value, "MM/dd/yyyy") & "' AND '" & Format(Dtp_Rpt_Asistencia_Empleados_Fecha_Termino.Value, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Simbologia<>'F'"
    Mi_SQL = Mi_SQL & " AND Adm_Asistencias.Horas_Extra>0"
    Set Rs_Consulta_Datos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Datos.EOF
        Contador_Registos = Contador_Registos + 1
        Print #3, "P2002" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Empleado") _
            & "|" & Chr(34) & "MX" _
            & "|" & Chr(34) & "001" _
            & "|" & Chr(34) & "INS" _
            & "|" & Chr(34) & "2002" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") _
            & "|" & Chr(34) & Format(Rs_Consulta_Datos.rdoColumns("Fecha"), "yyyyMMdd") _
            & "|" & Chr(34) & Format(Rs_Consulta_Datos.rdoColumns("Fecha"), "yyyyMMdd") _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("SAP_Incidencia") _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & Rs_Consulta_Datos.rdoColumns("Horas_Extra") _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & "" _
            & "|" & Chr(34) & ""
        Rs_Consulta_Datos.MoveNext
    Wend
    Rs_Consulta_Datos.Close
    'Fin del archivo
    Contador_Registos = Contador_Registos + 1
    Print #3, "TRAIL|" & Chr(34) & Contador_Registos
    Close #3
Exit Sub
HANDLER:
    If Archivo_Reporte_Abierto = True Then
        Close #3
    End If
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next
End Sub

Private Sub Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio_Change()
    Cmb_Rpt_Asistencia_Empleados_Periodo_Click
End Sub

Private Sub Dtp_Rpt_Asistencia_Empleados_Fecha_Inicio_Click()
    Cmb_Rpt_Asistencia_Empleados_Periodo_Click
End Sub

Private Sub Dtp_Rpt_Asistencia_Empleados_Fecha_Termino_Change()
    Cmb_Rpt_Asistencia_Empleados_Periodo_Click
End Sub

Private Sub Dtp_Rpt_Asistencia_Empleados_Fecha_Termino_Click()
    Cmb_Rpt_Asistencia_Empleados_Periodo_Click
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
    'Colapsa por Supervisor
    If Grid_Asistencia_Empleados.Rows > 0 Then
        Grid_Asistencia_Empleados.FixedRows = 2
        For Renglon_Procesar = 1 To Grid_Asistencia_Empleados.Rows - 1
            If Grid_Asistencia_Empleados.TextMatrix(Renglon_Procesar, 4) = "-" Then
                Grid_Asistencia_Empleados.Col = 4
                Call Grid_Asistencia_Empleados_Click
            End If
        Next Renglon_Procesar
    End If
    'Colapsa POr empresa
'    If Grid_Asistencia_Empleados.Rows > 0 Then
'        Grid_Asistencia_Empleados.FixedRows = 1
'        For Renglon_Procesar = 1 To Grid_Asistencia_Empleados.Rows - 1
'            If Grid_Asistencia_Empleados.TextMatrix(Renglon_Procesar, 1) = "-" Then
'                Grid_Asistencia_Empleados.Col = 1
'                Call Grid_Asistencia_Empleados_Click
'            End If
'        Next Renglon_Procesar
'    End If
    
End Sub

Private Sub Grid_Asistencia_Empleados_Click()
Dim Renglon As Integer     'Indica que renglon se esta consulltando
Dim Fila As Integer        'Contador de filas
'If Grid_Asistencia_Empleados.Rows > 1 Then
'    If Grid_Asistencia_Empleados.Col <= 4 And Grid_Asistencia_Empleados.Row > 0 Then
'        'And Grid_Importacion_Archivo.TextMatrix(Renglon, 1) = "-"
'        If Collapsing = False Then
'            Renglon = Grid_Asistencia_Empleados.MouseRow
'        Else
'            Renglon = Renglon_Procesar
'        End If
'        If Renglon < 1 Then Exit Sub
'
'        While Renglon > 0 And Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'            Renglon = Renglon - 1
'        Wend
'        If Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "-" Then
'            Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "+"
'        Else
'            Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "-"
'        End If
'
'        Renglon = Renglon + 1
'        If Renglon < Grid_Asistencia_Empleados.Rows Then
'            If Grid_Asistencia_Empleados.RowHeight(Renglon) = 1 Then
'                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'                    Grid_Asistencia_Empleados.RowHeight(Renglon) = -1
'                    Renglon = Renglon + 1
'                    If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
'                Loop
'            Else
'                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'                    Grid_Asistencia_Empleados.RowHeight(Renglon) = 0
'                    Renglon = Renglon + 1
'                    If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
'                Loop
'            End If
'        End If
'        Grid_Asistencia_Empleados.Col = 1
'    End If
'End If
'Dim renglon As Integer          'Almacena el renglon
    
    On Error GoTo Fin
    If Grid_Asistencia_Empleados.Rows > 1 Then
        If Collapsing = False Then
            Renglon = Grid_Asistencia_Empleados.MouseRow
        Else
            Renglon = Renglon_Procesar
        End If
'        If Grid_Asistencia_Empleados.MouseCol = 1 Then
'            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "" Then Exit Sub
'            If Renglon < 1 Then Exit Sub
'            'While renglon > 0 And Grid_Asistencia_Empleados.TextArray(renglon * Grid_Asistencia_Empleados.Cols) = ""
'            While Renglon > 0 And Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'                Renglon = Renglon - 1
'            Wend
'            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "" Then Exit Sub
'            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "-" Then
'                Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "+"
'            Else
'                Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = "-"
'            End If
'            Renglon = Renglon + 1
'            'Si el alto de los renglones es=0, entonces maximiza los renglones
'            If Grid_Asistencia_Empleados.RowHeight(Renglon) = 0 Then
'                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'                    Grid_Asistencia_Empleados.RowHeight(Renglon) = -1
'                    Renglon = Renglon + 1
'                    If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
'                Loop
'            Else 'Minimiza los renglones
'                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 1) = ""
'                  Grid_Asistencia_Empleados.RowHeight(Renglon) = 0
'                  Renglon = Renglon + 1
'                  If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
'                Loop
'            End If
'        Else
            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "" Then Exit Sub
            If Renglon < 1 Then Exit Sub
            'While renglon > 0 And Grid_Asistencia_Empleados.TextArray(renglon * Grid_Asistencia_Empleados.Cols) = ""
            While Renglon > 0 And Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = ""
                Renglon = Renglon - 1
            Wend
            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "" Then Exit Sub
            If Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "-" Then
                Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "+"
            Else
                If Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "+" Then
                    Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = "-"
                End If
            End If
            Renglon = Renglon + 1
            'Si el alto de los renglones es=0, entonces maximiza los renglones
            If Grid_Asistencia_Empleados.RowHeight(Renglon) = 0 Then
                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = ""
                    Grid_Asistencia_Empleados.RowHeight(Renglon) = -1
                    Renglon = Renglon + 1
                    If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
                Loop
            Else 'Minimiza los renglones
                Do While Grid_Asistencia_Empleados.TextMatrix(Renglon, 4) = ""
                  Grid_Asistencia_Empleados.RowHeight(Renglon) = 0
                  Renglon = Renglon + 1
                  If Renglon >= Grid_Asistencia_Empleados.Rows Then Exit Do
                Loop
            End If
'        End If
    End If
    'Oculta los controles
'    Muestra_Controles (False)
Fin:

End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Open Ruta_Temporal & Operacion & ".txt" For Output As #1
    Open Ruta_Temporal & Operacion & "xls.txt" For Output As #2 'Reporte a xls
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
End Sub

Private Sub Grid_Asistencia_Empleados_DblClick()
Dim Frm_Movimientos_Permisos As New Frm_Adm_Solicitud_Permisos      'Forma de Permisos
Dim Frm_Movimientos_Incidencias As New Frm_Adm_Incidencias_Extraordinarias   'Forma de Incidencias extraordinarias
Dim Tipo_Inventario As String       'Tipo de inventario a consultar
Dim No_Movimiento As String         'No movimiento a consultar
Dim Cadena_Movimiento As String     'Cadena temporal con los datos de la referencia
With Grid_Asistencia_Empleados
    If .Rows > 1 Then
        'Obtiene el no. de movimiento y el tipo de incidencia
        Cadena_Movimiento = Right(.TextMatrix(.RowSel, .ColSel), 12)
        Tipo_Inventario = Left(Cadena_Movimiento, 1)
        No_Movimiento = Right(Cadena_Movimiento, 10)
        If IsNumeric(No_Movimiento) = False Then Exit Sub
        If Tipo_Inventario = "P" Then
            If Conectar_Ayudante.Formulario_Cargado("SOLICITUD DE PERMISOS AS_EM") Then
                Conectar_Ayudante.Enfocar ("SOLICITUD DE PERMISOS AS_EM")
            Else
                Load Frm_Movimientos_Permisos
                Frm_Movimientos_Permisos.Top = 0
                Call Conectar_Ayudante.Cargar_Picture(Frm_Movimientos_Permisos.Pic_Solicitud_Permisos, Frm_Movimientos_Permisos)
                Frm_Movimientos_Permisos.Operacion = "Permisos"
                Frm_Movimientos_Permisos.Caption = "SOLICITUD DE PERMISOS AS_EM"
                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Asistencias_Empleados", Frm_Movimientos_Permisos)
                Frm_Movimientos_Permisos.Inicializa
                Frm_Movimientos_Permisos.Llenar_Informacion_Permiso (No_Movimiento)
            End If

        Else
            If Tipo_Inventario = "E" Then
                If Conectar_Ayudante.Formulario_Cargado("INCIDENCIAS EXTRAORDINARIAS AS_EM") Then
                    Conectar_Ayudante.Enfocar ("INCIDENCIAS EXTRAORDINARIAS AS_EM")
                Else
                    Load Frm_Movimientos_Incidencias
                    'Frm_Permisos.Height = 3510
                    'Frm_Permisos.Width = 7080
                    Frm_Movimientos_Incidencias.Top = 0
                    Call Conectar_Ayudante.Cargar_Picture(Frm_Movimientos_Incidencias.Pic_Solicitud_Permisos, Frm_Movimientos_Incidencias)
                    Frm_Movimientos_Incidencias.Operacion = "Permisos"
                    Frm_Movimientos_Incidencias.Pic_Logo.Visible = True
                    Frm_Movimientos_Incidencias.Pic_Logo.ZOrder vbBringToFront
                    Frm_Movimientos_Incidencias.Caption = "INCIDENCIAS EXTRAORDINARIAS AS_EM"
                    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Asistencias_Empleados", Frm_Movimientos_Incidencias)
                    Frm_Movimientos_Incidencias.Inicializa
                    Frm_Movimientos_Incidencias.Llenar_Informacion_Permiso (No_Movimiento)
                End If

            End If
        End If
    End If
End With
End Sub


