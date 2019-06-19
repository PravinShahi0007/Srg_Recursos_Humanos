VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Cambio_Turno 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13455
   Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   135
      ScaleHeight     =   3645
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   495
      Width           =   6990
      Begin VB.Frame Fra_Adm_Validacion_Horas_Trabajadas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opciones"
         Height          =   3600
         Left            =   45
         TabIndex        =   15
         Top             =   45
         Width           =   6855
         Begin VB.ComboBox Cmb_Cambio_Turno_Departamento 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   930
            Width           =   5160
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Turno 
            Height          =   315
            Left            =   1545
            TabIndex        =   5
            Top             =   1650
            Width           =   5160
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Supervisor 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   570
            Width           =   5160
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Empresa 
            Height          =   315
            Left            =   1545
            TabIndex        =   1
            Top             =   225
            Width           =   5160
         End
         Begin VB.CommandButton Btn_Salir 
            Caption         =   "Salir"
            Height          =   690
            Left            =   5415
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2760
            UseMaskColor    =   -1  'True
            Width           =   1200
         End
         Begin VB.CommandButton Btn_Adm_Validacion_Horas_Generar 
            Caption         =   "Consultar"
            Height          =   690
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "A"
            Top             =   2760
            Width           =   1200
         End
         Begin MSComctlLib.ProgressBar PrgBar_Validacion_Horas 
            Height          =   510
            Left            =   1515
            TabIndex        =   19
            Top             =   2115
            Visible         =   0   'False
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   900
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.ComboBox Cmb_Cambio_Turno_Tripulacion 
            Height          =   315
            Left            =   1545
            TabIndex        =   4
            Top             =   1290
            Width           =   5160
         End
         Begin VB.Label Lbl_Tripulacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tripulacion"
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
            TabIndex        =   28
            Top             =   1350
            Width           =   960
         End
         Begin VB.Label Lbl_Turno 
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
            Left            =   135
            TabIndex        =   20
            Top             =   1710
            Width           =   510
         End
         Begin VB.Label Lbl_Supervisor 
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
            Left            =   135
            TabIndex        =   18
            Top             =   630
            Width           =   915
         End
         Begin VB.Label Lbl_Empresa 
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
            Left            =   135
            TabIndex        =   17
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Lbl_Departamento 
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
            Left            =   135
            TabIndex        =   16
            Top             =   990
            Width           =   1200
         End
      End
   End
   Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo_Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   0
      ScaleHeight     =   7260
      ScaleWidth      =   13425
      TabIndex        =   21
      Top             =   0
      Width           =   13425
      Begin VB.Frame Fra_Validacion_Horas_Trabajo_Lista 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleado"
         Enabled         =   0   'False
         Height          =   6465
         Left            =   30
         TabIndex        =   22
         Top             =   15
         Width           =   13380
         Begin VB.ComboBox Cmb_Turnos_Cambio_Todos 
            Height          =   315
            Left            =   11355
            TabIndex        =   29
            Top             =   135
            Width           =   1950
         End
         Begin MSComCtl2.DTPicker Dtp_Cambio_Turno 
            Height          =   315
            Left            =   8265
            TabIndex        =   27
            Top             =   135
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            Format          =   117899264
            CurrentDate     =   41015
         End
         Begin VB.ComboBox Cmb_Turnos_Cambio 
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   8370
            TabIndex        =   26
            Top             =   540
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox Chk_Seleccionar_Todas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Seleccionar Todo"
            Height          =   285
            Left            =   6525
            TabIndex        =   13
            Top             =   150
            Width           =   1620
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Validacion_Horas_Trabajo_Lista 
            Height          =   5925
            Left            =   60
            TabIndex        =   14
            Top             =   465
            Width           =   13245
            _ExtentX        =   23363
            _ExtentY        =   10451
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            AllowUserResizing=   1
            Appearance      =   0
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
            Left            =   150
            TabIndex        =   23
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.CommandButton Btn_Regresar 
         Caption         =   "Regresar"
         Height          =   690
         Left            =   6111
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6525
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Imprimir 
         Caption         =   "Imprimir"
         Height          =   690
         Left            =   2097
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "A"
         Top             =   6525
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Validar_Horas_Empleados 
         Caption         =   "Cambiar"
         Height          =   690
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "A"
         Top             =   6525
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Salir_2 
         Caption         =   "Salir"
         Height          =   690
         Left            =   12135
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6525
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Excel 
         Caption         =   "Excel"
         Height          =   690
         Left            =   4104
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "A"
         Top             =   6525
         Width           =   1200
      End
      Begin MSComctlLib.ProgressBar Prbar_Exportacion 
         Height          =   720
         Left            =   5310
         TabIndex        =   24
         Top             =   6510
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   1270
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   5490
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Lbl_Progreso_Exportacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5520
         TabIndex        =   25
         Top             =   6630
         Width           =   45
      End
   End
End
Attribute VB_Name = "Frm_Adm_Cambio_Turno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean
Dim Fecha As Date
Dim Manejo_Grid As Boolean
Public Opcion As String                     'Define la opcion para los procesos

Private Sub Btn_Adm_Validacion_Horas_Generar_Click()
    If Cmb_Adm_Validacion_Horas_Empresa.ListIndex > -1 Then
        Generar_Lista
    Else
        MsgBox "Seleccione una empresa", vbOKOnly + vbInformation, Me.Caption
        Cmb_Adm_Validacion_Horas_Empresa.SetFocus
    End If
End Sub

Private Sub Btn_Excel_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String
On Error GoTo handler
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 1 Then
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
            Call Exportar_Excel(Ruta_Temporal & Opcion & "xls.txt", Ruta_Exportacion, Prbar_Exportacion, Lbl_Progreso_Exportacion, Me)
        End If
    Else
        MsgBox "No existe información para exportar", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
handler:
    Exit Sub
End Sub

Private Sub Finalizar_Reporte()
    Close #1, #2
End Sub

Private Sub Btn_Imprimir_Click()
    Imprimir
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Adm_Validacion_Horas_Trabajo_Lista.Visible = False
    Pic_Adm_Validacion_Horas_Trabajo.Visible = True
    Me.Height = 4600
    Me.Width = 7180
End Sub

Private Sub Btn_Salir_2_Click()
    If MsgBox("¿Desea salir de la operacion?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Btn_Validar_Horas_Empleados_Click()
Dim Cont_Fila As Integer        'Recorrer el grid
Dim Guardar As Boolean          'Validar si existen registros que guardar
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 0 Then
        'Recorre el grid para saber si al menos se guardara algun registro
        For Cont_Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6) = "SI" Then
                Guardar = True
                Exit For
            End If
        Next
        If Guardar = False Then
            MsgBox "No ha seleccionado información para actualizar", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If MsgBox("¿La información es correcta?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            If Fra_Validacion_Horas_Trabajo_Lista.Enabled = True Then
                Guardar_Lista
            Else
                MsgBox "La lista ya fue guardada, si desea realizar algun cambio" + vbCrLf + _
                       "deberá realizar entrar de nuevo en la opcion de cambio de turno y" + vbCrLf + _
                       "generar nuevamente la lista", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            'Valida que la lista no se hay generado anteriormente
        End If
    End If
End Sub


Private Sub Chk_Seleccionar_Todas_Click()
Dim Fila As Integer     'Contador para recorrer el grid
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 0 Then
        For Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
            If Chk_Seleccionar_Todas.Value = 1 Then
                Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Fila, 6) = "SI"
            Else
                Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Fila, 6) = "NO"
            End If
        Next Fila
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID,Nombre", "Cat_Empresas", Cmb_Adm_Validacion_Horas_Empresa, 1, "Nombre")
        If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Empresa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Validacion_Horas_Empresa, KeyCode)
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Supervisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(Cmb_Adm_Validacion_Horas_Supervisor.Text) Then
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A' AND No_Tarjeta='" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "No_Tarjeta", "", True, "<-SELECCIONE->")
        Else
            Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Adm_Validacion_Horas_Supervisor, 1, "Apellido_Paterno", "AND Tipo='S' AND Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%' OR Apellido_Paterno LIKE '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%')", True, "<-SELECCIONE->")
        End If
        If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Turno_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Turno_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Adm_Validacion_Horas_Turno, KeyCode)
End Sub

Private Sub Cmb_Cambio_Turno_Departamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID,Nombre", "Cat_Departamentos", Cmb_Cambio_Turno_Departamento, 0, "Nombre", "", True, "<-SELECCIONE->")
        If Cmb_Cambio_Turno_Departamento.ListCount > 0 Then Cmb_Cambio_Turno_Departamento.ListIndex = 0
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Cambio_Turno_Tripulacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Gap_ID,Nombre", "Cat_Gaps", Cmb_Cambio_Turno_Tripulacion, 0, "Nombre", "", True, "<-SELECCIONE->")
        If Cmb_Cambio_Turno_Tripulacion.ListCount > 0 Then Cmb_Cambio_Turno_Tripulacion.ListIndex = 0
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Cmb_Turnos_Cambio_Click()
Dim Columna As Integer
    If Cmb_Turnos_Cambio.ListIndex > -1 Then
        With Grid_Validacion_Horas_Trabajo_Lista
            If .RowSel > 0 Then
                .TextMatrix(.RowSel, 7) = Cmb_Turnos_Cambio.Text
                .TextMatrix(.RowSel, 8) = Format(Cmb_Turnos_Cambio.ItemData(Cmb_Turnos_Cambio.ListIndex), "00000")
            End If
        End With
    End If
End Sub

Private Sub Cmb_Turnos_Cambio_Todos_Click()
Dim Cont_Grid As Integer
    If Cmb_Turnos_Cambio_Todos.ListIndex > -1 Then
        If Grid_Validacion_Horas_Trabajo_Lista.RowSel > 0 Then
            For Cont_Grid = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
                If Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Grid, 6)) = "SI" Then
                    Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Grid, 7) = Cmb_Turnos_Cambio_Todos.Text
                    Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Grid, 8) = Format(Cmb_Turnos_Cambio_Todos.ItemData(Cmb_Turnos_Cambio_Todos.ListIndex), "00000")
                End If
            Next
        End If
    End If
End Sub

Private Sub Cmb_Turnos_Cambio_Todos_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Form_Activate()
    If Pic_Adm_Validacion_Horas_Trabajo.Visible = True Then Cmb_Adm_Validacion_Horas_Empresa.SetFocus
End Sub

Private Sub Form_Load()
Dim Mi_SQL As String
Dim Rs_Empleados_Supervisor As rdoResultset
Dim Rs_Empleados_Departamento As rdoResultset

    'agrega la informacion en los combos
    Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Validacion_Horas_Empresa, 0, "Nombre")
    If Cmb_Adm_Validacion_Horas_Empresa.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Empresa.ListIndex = 0
    
    'Consulta Supervisor.
    Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
    Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
    Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
    Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
    Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
    Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Cmb_Adm_Validacion_Horas_Supervisor.Clear
    While Not Rs_Empleados_Supervisor.EOF
        Cmb_Adm_Validacion_Horas_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
        Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
        Rs_Empleados_Supervisor.MoveNext
    Wend
    Rs_Empleados_Supervisor.Close
    Cmb_Adm_Validacion_Horas_Supervisor.Text = "<-SELECCIONE->"
    
'    Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus = 'A'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Apellido_Paterno", "", True, "<-SELECCIONE->")
'    If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
'

    'Consulta Departamento.
    Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
    Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Cmb_Cambio_Turno_Departamento.Clear
    While Not Rs_Empleados_Departamento.EOF
        Cmb_Cambio_Turno_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
        Cmb_Cambio_Turno_Departamento.ItemData(Cmb_Cambio_Turno_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
        Rs_Empleados_Departamento.MoveNext
    Wend
    Rs_Empleados_Departamento.Close
    Cmb_Cambio_Turno_Departamento.Text = "<-SELECCIONE->"
       
'    Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID,Nombre", "Cat_Departamentos", Cmb_Cambio_Turno_Departamento, 0, "Nombre")
'    If Cmb_Cambio_Turno_Departamento.ListCount > 0 Then Cmb_Cambio_Turno_Departamento.ListIndex = 0
    
    Call Conectar_Ayudante.Llena_Combo_Item("Gap_ID,Nombre", "Cat_Gaps", Cmb_Cambio_Turno_Tripulacion, 0, "Nombre", "", True, "<-SELECCIONE->")
    If Cmb_Cambio_Turno_Tripulacion.ListCount > 0 Then Cmb_Cambio_Turno_Tripulacion.ListIndex = 0
    Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Adm_Validacion_Horas_Turno, 0, "Nombre", "", True, "<-SELECCIONE->")
    If Cmb_Adm_Validacion_Horas_Turno.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Turno.ListIndex = 0
    Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID,Nombre", "Cat_Turnos", Cmb_Turnos_Cambio_Todos, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID,Nombre", "Cat_Turnos", Cmb_Turnos_Cambio, 0, "Nombre")
    Dtp_Cambio_Turno.Value = Now
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_Click()
    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 0 And .RowSel > 0 Then
            Cmb_Turnos_Cambio.Visible = False
            If Manejo_Grid = True Then
                Select Case .ColSel
                    Case 7 'Horas
                        If Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 6)) = "SI" Then
                            Call Conectar_Ayudante.Mover_Control_Grid_ComboBox(Grid_Validacion_Horas_Trabajo_Lista, Cmb_Turnos_Cambio)
                        End If
                End Select
            End If
        End If
    End With
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_DblClick()
    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 0 And .RowSel > 0 Then
            Cmb_Turnos_Cambio.Visible = False
            If Manejo_Grid = True Then
                Select Case .ColSel
                    Case 6
                        If Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 6)) = "SI" Then
                            Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 6) = "NO"
                        Else
                            Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 6) = "SI"
                        End If
                End Select
            End If
        End If
    End With
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_EnterCell()
    Grid_Validacion_Horas_Trabajo_Lista_Click
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_LeaveCell()
    'Grid_Validacion_Horas_Trabajo_Lista.CellBackColor = vbWhite
End Sub

Private Sub Cmb_Turnos_Cambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
        'Guarda la informacion y oculta el check
        Cmb_Turnos_Cambio.Visible = False
        Call Mover_Control_Grid_Procesos(KeyCode)
        If KeyCode = 37 Or KeyCode = 39 Then Grid_Validacion_Horas_Trabajo_Lista.SetFocus
    End If
End Sub

Private Sub Cmb_Turnos_Cambio_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Mover_Control_Grid_Procesos
    'DESCRIPCIÓN: Oculta y visuliza los controles del grid de productos segun
    '             la tecla de direccion que se presione
    'PARÁMETROS: Tecla: contiene el numero de la tecla oprimida por el usuario
    'CREO      : José Antonio López Hernández
    'FECHA_CREO: 11/Jul/2007 1:44 pm
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Mover_Control_Grid_Procesos(Tecla As Integer)
    Select Case Tecla
        Case 37 'Izquierda
            'Valida que no sea la primer columna para mover a la anterior
            If Grid_Validacion_Horas_Trabajo_Lista.ColSel > 4 Then
                Grid_Validacion_Horas_Trabajo_Lista.Col = Grid_Validacion_Horas_Trabajo_Lista.ColSel - 1
            Else
                'Si es la ultima columna mueve el cursor hasta la columna de la fecha
                If Grid_Validacion_Horas_Trabajo_Lista.ColSel = 4 Then
                    Grid_Validacion_Horas_Trabajo_Lista.Col = 4
                Else
                    If Grid_Validacion_Horas_Trabajo_Lista.ColSel = 1 Then
                        Grid_Validacion_Horas_Trabajo_Lista.Col = 1
                    Else
                        Grid_Validacion_Horas_Trabajo_Lista_Click
                    End If
                End If
            End If

        Case 38 'Arriba
            'Valida que no sea el ultimo renglon del grid para mover al siguiente
            If Grid_Validacion_Horas_Trabajo_Lista.RowSel > 1 Then
                Grid_Validacion_Horas_Trabajo_Lista.Row = Grid_Validacion_Horas_Trabajo_Lista.RowSel - 1
            End If
            
            Grid_Validacion_Horas_Trabajo_Lista_Click
        
        Case 39 'Derecha
            'Valida que no sea la ultima columna para mover a la siguiente
            If Grid_Validacion_Horas_Trabajo_Lista.ColSel < 9 Then
                Grid_Validacion_Horas_Trabajo_Lista.ColSel = Grid_Validacion_Horas_Trabajo_Lista.ColSel + 1
                
            Else
                'Si es la ultima columna mueve el cursor hasta la columna de la fecha
                If Grid_Validacion_Horas_Trabajo_Lista.ColSel = 10 Then
                    Grid_Validacion_Horas_Trabajo_Lista.Col = 10
                Else
                    Grid_Validacion_Horas_Trabajo_Lista_Click
                End If
            End If
        
        Case 40 'Abajo
            'Valida que no sea el ultimo renglon del grid para mover al siguiente
            If Grid_Validacion_Horas_Trabajo_Lista.RowSel < (Grid_Validacion_Horas_Trabajo_Lista.Rows - 1) Then
                Grid_Validacion_Horas_Trabajo_Lista.Row = Grid_Validacion_Horas_Trabajo_Lista.RowSel + 1
            End If
    
            Grid_Validacion_Horas_Trabajo_Lista_Click
            
    End Select
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Generar_Lista
'DESCRIPCION: Genera la lista de empleados para validar la hora
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 16-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Lista()
Dim Consulta_Cat_Empleados As rdoResultset                      'Informacion de la lista de empleados
Dim Rs_Consulta_Adm_Asistencias_Detalles As rdoResultset        'Obtiene la asistencia
Dim Rs_Consulta_Adm_Permisos As rdoResultset                    'Informacion de permisos del empleado
Dim Rs_Consulta_Adm_Vacaciones As rdoResultset                  'Informacion de Vacaciones del empleado
Dim Rs_Consulta_Adm_Inasistencias As rdoResultset               'Informacion de Vacaciones del empleado
Dim Rs_Consulta_Informacion_Turnos As rdoResultset              'Informacion del empleado
Dim Rs_Consulta_Dia_Feriado As rdoResultset                     'Informacion del empleado
Dim Rs_Consulta_Cat_Empleados As rdoResultset                   'Informacion del empleado
Dim Permiso As String                                           'Informacion del permiso
Dim Tipo_Incidencia As String                                   'Tipo de incidencia generada por el ausentismo
Dim Partida As Integer                                          'No consecutivo de la lista
Dim Turno_Empleado As String                                    'Identificador del turno del empleado
Dim No_Movimiento As Double                                     'No de movimiento
Dim Validada As String                                          'Indica si la asistencia ya fue validada
Dim Empresa_Sindicalizada As Boolean                            'Identifica si la empresa es sindicalizada
Dim Colorear_Fila As Boolean                                    'Define si la fila se coloreara o no
Dim Columna As Integer

On Error GoTo handler:
    Partida = 0
    Grid_Validacion_Horas_Trabajo_Lista.Rows = 0
    Grid_Validacion_Horas_Trabajo_Lista.Cols = 9
    'Informacion para la barra de progreso
    Mi_SQL = "SELECT ISNULL(COUNT(No_Tarjeta),0) AS Empleados"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Estatus='A'"
    Mi_SQL = Mi_SQL & " AND Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
    If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Supervisor_ID = '" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Cambio_Turno_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Departamento_ID='" & Format(Cmb_Cambio_Turno_Departamento.ItemData(Cmb_Cambio_Turno_Departamento.ListIndex), "00000") & "'"
    End If
    If Cmb_Cambio_Turno_Tripulacion.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Gap_ID='" & Format(Cmb_Cambio_Turno_Tripulacion.ItemData(Cmb_Cambio_Turno_Tripulacion.ListIndex), "00000") & "'"
    End If
    If Cmb_Adm_Validacion_Horas_Turno.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Turno_ID='" & Format(Cmb_Adm_Validacion_Horas_Turno.ItemData(Cmb_Adm_Validacion_Horas_Turno.ListIndex), "00000") & "'"
    End If
    Set Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'obtiene la informacion para configurar el progress bar
        If Val(Consulta_Cat_Empleados.rdoColumns("Empleados")) > 0 Then
            PrgBar_Validacion_Horas.Max = Val(Consulta_Cat_Empleados.rdoColumns("Empleados"))
        End If
    Consulta_Cat_Empleados.Close
    Set Consulta_Cat_Empleados = Nothing
    'Informacion para la lista
    Mi_SQL = "SELECT CE.Empleado_ID,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre,CE.No_Tarjeta,CE.Turno_ID,CD.Nombre AS Departamento,Cat_Turnos.Nombre AS Turno,ISNULL(Cat_Gaps.Nombre,'') AS Tripulacion"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE INNER JOIN Cat_Departamentos CD ON CE.Departamento_ID=CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " INNER JOIN Cat_Turnos ON CE.Turno_ID=Cat_Turnos.Turno_ID"
    Mi_SQL = Mi_SQL & " LEFT JOIN Cat_Gaps ON CE.Gap_ID=Cat_Gaps.Gap_ID"
    Mi_SQL = Mi_SQL & " WHERE CE.Estatus='A'"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
    If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
    End If
    If Cmb_Adm_Validacion_Horas_Turno.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Turno_ID='" & Format(Cmb_Adm_Validacion_Horas_Turno.ItemData(Cmb_Adm_Validacion_Horas_Turno.ListIndex), "00000") & "'"
    End If
    If Cmb_Cambio_Turno_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Cambio_Turno_Departamento.ItemData(Cmb_Cambio_Turno_Departamento.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " ORDER BY CE.No_Tarjeta"
    Set Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Consulta_Cat_Empleados
        If Not .EOF Then
            Me.MousePointer = 11
            Me.Refresh
            PrgBar_Validacion_Horas.Visible = True
            PrgBar_Validacion_Horas.Value = 0
            Empresa_Sindicalizada = False
            'Identifica si la empresa es o no sindicalizada
            If InStr(1, Cmb_Adm_Validacion_Horas_Empresa.Text, "SINDI") > 0 Then
                Empresa_Sindicalizada = True
            End If
            Call Encabezado_Reporte("CAMBIO DE TURNO", DateAdd("s", 1, Now), DateAdd("s", 1, Now))
            'Agrega el encabezado
            Grid_Validacion_Horas_Trabajo_Lista.AddItem "Empleado_ID" _
                & Chr(9) & "No." & Chr(9) & "Departamento" _
                & Chr(9) & "Tripulacion" & Chr(9) & "Nombre" _
                & Chr(9) & "Turno" & Chr(9) & "Cambiar" _
                & Chr(9) & "Nuevo" & Chr(9) & "ID_Nuevo"
            Print #1, ""
            Print #2, "No.|Departamento|Tripulacion|Empleado|Turno"
            Partida = 0
            While Not .EOF
                Debug.Print .rdoColumns("Empleado_ID")
                Me.Refresh
                'Agrega el dato en el grid
                Grid_Validacion_Horas_Trabajo_Lista.AddItem .rdoColumns("Empleado_ID") _
                    & Chr(9) & .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Departamento") _
                    & Chr(9) & .rdoColumns("Tripulacion") & Chr(9) & .rdoColumns("Nombre") _
                    & Chr(9) & .rdoColumns("Turno") _
                    & Chr(9) & "NO" _
                    & Chr(9) & "" _
                    & Chr(9) & ""
                Print #1, ""
                Print #2, .rdoColumns("No_Tarjeta"); "|"; .rdoColumns("Departamento"); "|"; .rdoColumns("Tripulacion"); "|"; .rdoColumns("Nombre"); "|"; .rdoColumns("Turno")
                Me.Refresh
                PrgBar_Validacion_Horas.Value = PrgBar_Validacion_Horas.Value + 1
                .MoveNext
            Wend
            Call Finalizar_Reporte
        End If
    End With
    Me.Refresh
    Manejo_Grid = True
    'Configuracion del grid
    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 1 Then .FixedRows = 1
            .FixedCols = 2
            .ColWidth(0) = 0        'Empleado_ID
            .ColWidth(1) = 800      'No_Tarjeta
            .ColAlignment(1) = flexAlignCenterCenter
            .ColWidth(2) = 2000     'Departamento
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(3) = 2000     'Tripulacion
            .ColAlignment(3) = flexAlignLeftCenter
            .ColWidth(4) = 4500     'Empleado
            .ColAlignment(4) = flexAlignLeftCenter
            .ColWidth(5) = 1500     'Turno
            .ColAlignment(5) = flexAlignLeftCenter
            .ColWidth(6) = 600      'Validar
            .ColWidth(7) = 1500     'Nuevo
            .ColAlignment(7) = flexAlignLeftCenter
            .ColWidth(8) = 0       'ID_Nuevo
    End With
    Me.Refresh
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 1 Then
        Pic_Adm_Validacion_Horas_Trabajo_Lista.Visible = True
        Pic_Adm_Validacion_Horas_Trabajo.Visible = False
        Me.Height = 7740
        Me.Width = 13545
        Me.Top = 0
        Me.Left = 0
        Fra_Validacion_Horas_Trabajo_Lista.Enabled = True
        Lbl_Validacion_Horas_Supervisor.Caption = Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text)
        Btn_Validar_Horas_Empleados.Enabled = False
        If DateDiff("d", Now, Fecha) < 0 Then
            Btn_Validar_Horas_Empleados.Enabled = True
        End If
        Chk_Seleccionar_Todas.Value = 0
    Else
        Lbl_Validacion_Horas_Supervisor.Caption = Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text)
        MsgBox "No se encontraron empleados con los parámetros seleccionados", vbExclamation + vbOKOnly, Me.Caption
    End If
    Me.Refresh
    PrgBar_Validacion_Horas.Visible = False
    Me.MousePointer = 0
Exit Sub
handler:
    Call Finalizar_Reporte
    Me.MousePointer = 0
    Debug.Print Err.Description
    PrgBar_Validacion_Horas.Visible = False
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    Open Ruta_Temporal & Opcion & ".txt" For Output As #1
    Open Ruta_Temporal & Opcion & "xls.txt" For Output As #2 'Reporte a xls
    'Archivo_Reporte_Abierto = True
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

'*******************************************************************************
'NOMBRE_FUNCION: Guardar_Lista
'DESCRIPCION: Actualiza los datos del turno
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 16-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Guardar_Lista()
Dim Rs_Adm_Cambios_Turnos As rdoResultset
Dim Cont_Fila As Integer
Dim cadena_mensaje As String

On Error GoTo handler
    Conexion_Base.BeginTrans
    For Cont_Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
        Me.MousePointer = 11
        Debug.Print Cont_Fila & " " & Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 1)
        cadena_mensaje = ""
        If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6) = "SI" And Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7) <> "" And Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8)) <> "" Then
            Set Rs_Adm_Cambios_Turnos = Conectar_Ayudante.Recordset_Agregar("Adm_Cambios_Turnos")
            With Rs_Adm_Cambios_Turnos
                .AddNew
                    .rdoColumns("Empleado_ID") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 0))
                    .rdoColumns("Turno_Nuevo_ID") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8))
                    .rdoColumns("Fecha_Cambio") = Format(Dtp_Cambio_Turno.Value, "MM/dd/yyyy")
                    .rdoColumns("Estatus") = "PENDIENTE"
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
            End With
            Rs_Adm_Cambios_Turnos.Close
        End If
    Next
    Conexion_Base.CommitTrans
    Fra_Validacion_Horas_Trabajo_Lista.Enabled = False
    Me.MousePointer = 0
    MsgBox "Información Guardada Correctamente", vbInformation + vbOKOnly, Me.Caption
    'Valida si hace el cambio al día de hoy para que se ejecute el cambio automático
    If Format(Dtp_Cambio_Turno.Value, "yyyyMMdd") <= Format(Now, "yyyyMMdd") Then
        'If MsgBox("¿Desea ejecutar el cambio de turnos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            Actualiza_Turnos_Programacion
        'End If
    End If
Exit Sub
handler:    'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
    Me.MousePointer = 0
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Imprimir()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset
Dim Movimiento As String
Dim linea As String 'Obtiene el texto a imprimir
Dim X As Printer
Dim Horas_Trabajadas As Double
Dim contar_linea As Integer

On Error GoTo handler
    If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex < 0 Then
        MsgBox "No ha seleccionado un supervisor", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    Mi_SQL = "SELECT AA.No_Tarjeta, AA.Fecha, AA.Hora_Entrada, AA.Hora_Salida, AA.Hora_Entrada_Comida, AA.Hora_Salida_Comida,"
    Mi_SQL = Mi_SQL & " AA.Horas_Aprobadas, AA.Simbologia,"
    Mi_SQL = Mi_SQL & " AA.Empleado_ID, (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE, Adm_Asistencias AA"
    Mi_SQL = Mi_SQL & " WHERE CE.Empleado_ID = AA.Empleado_ID"
    'Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND AA.Supervisor_ID = '" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
    Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
    Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Now, "MM/dd/yyyy") & "'"
    If Cmb_Adm_Validacion_Horas_Turno.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND AA.Turno_ID = '" & Format(Cmb_Adm_Validacion_Horas_Turno.ItemData(Cmb_Adm_Validacion_Horas_Turno.ListIndex), "00000") & "'"
    End If
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            Call Encabezado_Reporte("VALIDACION DE HORAS TRABAJADAS", DateAdd("s", 1, Now), DateAdd("s", 1, Now))
            Print #1, "Empresa:   " & Cmb_Adm_Validacion_Horas_Empresa.Text
            Print #1, "Supervisor: " & Cmb_Adm_Validacion_Horas_Supervisor.Text
            Print #1, "Turno: " & Cmb_Adm_Validacion_Horas_Turno.Text
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            Print #1, "No Tarjeta   Nombre                                         E      S   Hr. Hrs Acuerdo Incidencia               Tipo      "
            Print #1, "--------------------------------------------------------------------------------------------------------------------------"
            While Not .EOF
                Horas_Trabajadas = 0
                Select Case .rdoColumns("Simbologia")
                    Case "AS": Movimiento = "Asistencia"
                    Case "FI": Movimiento = "Falta Injustificada"
                    Case "FJ": Movimiento = "Falta Justificada"
                    Case "II": Movimiento = "Inasistencia por incapacidad"
                    Case "ID": Movimiento = "Inasistencia por derecho"
                    Case "RE": Movimiento = "Retardo"
                End Select
                Horas_Trabajadas = Format((DateDiff("n", Format(.rdoColumns("Hora_Entrada"), "HH:mm"), Format(.rdoColumns("Hora_Salida"), "HH:mm"))) / 60, "#0.00") - Format((DateDiff("n", Format(.rdoColumns("Hora_Entrada_Comida"), "HH:mm"), Format(.rdoColumns("Hora_Salida_Comida"), "HH:mm"))) / 60, "#0.00")
                Print #1, Conectar_Ayudante.Alinea_Derecha(.rdoColumns("No_Tarjeta"), 10); Spc(3); _
                          Mid(.rdoColumns("Nombre"), 1, 40); Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Hora_Entrada"), "HH:mm"), 47 - Len(Mid(.rdoColumns("Nombre"), 1, 40))); _
                          Conectar_Ayudante.Alinea_Derecha(Format(.rdoColumns("Hora_Salida"), "HH:mm"), 7); Conectar_Ayudante.Alinea_Derecha(CStr(Horas_Trabajadas), 6); Spc(1); Val(.rdoColumns("Horas_Aprobadas")); _
                          Spc(12 - Len(.rdoColumns("Horas_Aprobadas"))); Mid(Movimiento, 1, 25); Spc(25 - Len(Mid(Movimiento, 1, 25))); .rdoColumns("Simbologia"); Spc(6 - Len(.rdoColumns("Simbologia")))
                .MoveNext
            Wend
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1, Conectar_Ayudante.Alinea_Derecha("__________________________", 123)
            Print #1, Conectar_Ayudante.Alinea_Derecha("           FIRMA          ", 123)
            .Close
            Finalizar_Reporte
            Printer.FontSize = 8
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 11
            Printer.Font = "COURIER NEW"
            Printer.Print
            Printer.FontSize = 8
            Printer.Font = "Courier New"
            Open Ruta_Temporal & "ListaValidacion.txt" For Input As #1
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
        Else
            MsgBox "No existe información que imprimir", vbInformation + vbOKOnly, Me.Caption
        End If
        Set Rs_Consulta_Adm_Asistencias = Nothing
    End With
Exit Sub
handler:
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox Err.Description
    For Each Er In rdoErrors
        If Mid(Er, 1, 5) = "01S02" Then MsgBox "No se encontro la impresora", vbCritical + vbOKOnly, Me.Caption
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Enviar_Correo
    'DESCRIPCIÓN:           Envia el correo con los parametros establecido
    'PARÁMETROS :           From_Email: correo de quien envia
    '                       Nombre_From: Nombre quien envia
    '                       To_Email:correo a quien se envia
    '                       Nombre_To: nombre a quien se envia
    '                       Asunto: asunto del correo
    '                       Mensaje_Email: mensaje del correo
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           19 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Enviar_Correo(From_Email As String, Nombre_From As String, To_Email As String, Nombre_To As String, Asunto As String, Mensaje_Email As String)
    Set poSendMail = New clsSendMail
    Me.MousePointer = vbHourglass

    With poSendMail
        ' Propiedades opcionales para envio de correo, deberan ser primero configuradas si se utilizan
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
        ' Propiedades básicas para envio de correos
        .SMTPHost = Servidor_SMTP           ' Required the fist time, optional thereafter
        .From = From_Email                  ' Required the fist time, optional thereafte
        .FromDisplayName = Nombre_From      ' Optional, saved after first use
        .Recipient = To_Email               ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = Nombre_To   ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName          ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc              ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text      ' Optional, used when different than 'From' address
        .Subject = Asunto                   ' Optional
        .Message = Mensaje_Email            ' Optional
        '.Attachment = Trim(txtAttach.Text)  ' Optional, separate multiple entries with delimiter character

        ' Propiedades opcionales adicionales, utilizar si son requeridas por la aplicacion
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        '.UserName = txtUserName                     ' Optional, default = Null String
        '.Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' Propiedades avanzadas, cambiar solo si tienes una buena razon para hacerlos
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
         .SMTPPort = Puerto_SMTP                    ' Optional, default = 25

        ' Envio de correo
        ' .Connect                                  ' Optional, use when sending bulk mail
        .send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'txtServer.Text = .SMTPhost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Set poSendMail = Nothing
    Me.MousePointer = vbDefault
End Function

