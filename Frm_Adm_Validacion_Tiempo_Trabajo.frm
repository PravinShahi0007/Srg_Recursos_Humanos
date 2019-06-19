VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Adm_Validación_Tiempo_Trabajo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14760
   Begin VB.TextBox Txt_Validacion_Justificacion_Horas_Grid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8160
      MaxLength       =   200
      TabIndex        =   43
      Top             =   600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   150
      ScaleHeight     =   3420
      ScaleWidth      =   6960
      TabIndex        =   0
      Top             =   495
      Width           =   6990
      Begin VB.Frame Fra_Adm_Validacion_Horas_Trabajadas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opciones"
         Height          =   3345
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   6855
         Begin VB.CommandButton Btn_Adm_Validacion_Horas_Generar 
            Caption         =   "Generar"
            Height          =   690
            Left            =   1260
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "A"
            Top             =   2550
            Width           =   1200
         End
         Begin VB.CommandButton Btn_Salir 
            Caption         =   "Salir"
            Height          =   690
            Left            =   5295
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2550
            UseMaskColor    =   -1  'True
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Empresa 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   225
            Width           =   5235
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Supervisor 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   618
            Width           =   5235
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Turno 
            Height          =   315
            Left            =   1440
            TabIndex        =   4
            Top             =   1011
            Width           =   5235
         End
         Begin VB.CommandButton Btn_Imprimir_2 
            Caption         =   "Imprimir"
            Height          =   690
            Left            =   3277
            Style           =   1  'Graphical
            TabIndex        =   3
            Tag             =   "A"
            Top             =   2550
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_Adm_Validacion_Horas_Departamento 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   1410
            Width           =   5235
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Adm_Validacion_Horas_Fecha 
            Height          =   315
            Left            =   1440
            TabIndex        =   9
            Top             =   1815
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   112787456
            CurrentDate     =   39940
         End
         Begin MSComctlLib.ProgressBar PrgBar_Validacion_Horas 
            Height          =   240
            Left            =   300
            TabIndex        =   10
            Top             =   2235
            Visible         =   0   'False
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   423
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
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
            TabIndex        =   15
            Top             =   1875
            Width           =   540
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
            TabIndex        =   14
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label1 
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
            TabIndex        =   13
            Top             =   678
            Width           =   915
         End
         Begin VB.Label Label2 
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
            TabIndex        =   12
            Top             =   1071
            Width           =   510
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
            Left            =   120
            TabIndex        =   11
            Top             =   1470
            Width           =   1200
         End
      End
   End
   Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo_Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7965
      Left            =   0
      ScaleHeight     =   7965
      ScaleWidth      =   14700
      TabIndex        =   16
      Top             =   0
      Width           =   14700
      Begin VB.CommandButton Btn_Excel 
         Caption         =   "Excel"
         Height          =   690
         Left            =   4514
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "A"
         Top             =   7275
         Width           =   1200
      End
      Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencido 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   5400
         ScaleHeight     =   240
         ScaleWidth      =   285
         TabIndex        =   32
         Top             =   6960
         Width           =   285
      End
      Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencer 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2970
         ScaleHeight     =   240
         ScaleWidth      =   285
         TabIndex        =   31
         Top             =   6960
         Width           =   285
      End
      Begin VB.PictureBox Pic_Adm_Validacion_Horas_Trabajo_Lista_Incosistencia 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   90
         ScaleHeight     =   240
         ScaleWidth      =   285
         TabIndex        =   30
         Top             =   6960
         Width           =   285
      End
      Begin VB.CommandButton Btn_Anterior 
         Caption         =   "Anterior"
         Height          =   690
         Left            =   8938
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7275
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Siguiente 
         Caption         =   "Siguiente"
         Height          =   690
         Left            =   11150
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7275
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Salir_2 
         Caption         =   "Salir"
         Height          =   690
         Left            =   13365
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7275
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Validar_Horas_Empleados 
         Caption         =   "Validar"
         Height          =   690
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "A"
         Top             =   7275
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Imprimir 
         Caption         =   "Imprimir"
         Height          =   690
         Left            =   2302
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "A"
         Top             =   7275
         Width           =   1200
      End
      Begin VB.CommandButton Btn_Regresar 
         Caption         =   "Regresar"
         Height          =   690
         Left            =   6726
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   7275
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin MSComctlLib.ProgressBar Pbar_Validacion 
         Height          =   720
         Left            =   1290
         TabIndex        =   17
         Top             =   7260
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   1270
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Prbar_Exportacion 
         Height          =   720
         Left            =   5715
         TabIndex        =   34
         Top             =   7260
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
         Left            =   6015
         Top             =   7395
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Fra_Validacion_Horas_Trabajo_Lista 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleado"
         Enabled         =   0   'False
         Height          =   6915
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   14655
         Begin VB.CheckBox Chk_Seleccionar_Todas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Seleccionar Todo"
            Height          =   285
            Left            =   12660
            TabIndex        =   20
            Top             =   135
            Width           =   1815
         End
         Begin VB.TextBox Txt_Validacion_Horas_Grid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6975
            MaxLength       =   5
            TabIndex        =   19
            Top             =   540
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Validacion_Horas_Trabajo_Lista 
            Height          =   6375
            Left            =   45
            TabIndex        =   21
            Top             =   465
            Width           =   14550
            _ExtentX        =   25665
            _ExtentY        =   11245
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Lbl_Validacion_Horas_Fecha_Lista 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha: dd/MMM/yyyy"
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
            Left            =   6285
            TabIndex        =   23
            Top             =   225
            Width           =   1875
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
            TabIndex        =   22
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.Label Lbl_Progreso_Exportacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5520
         TabIndex        =   42
         Top             =   6630
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contrato vencido"
         Height          =   195
         Left            =   5715
         TabIndex        =   41
         Top             =   6983
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contrato por vencer"
         Height          =   195
         Left            =   3285
         TabIndex        =   40
         Top             =   6983
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incosistencia en la checada"
         Height          =   195
         Left            =   405
         TabIndex        =   39
         Top             =   6983
         Width           =   1995
      End
      Begin VB.Label Lbl_Empleados_12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mas de límite de horas trabajadas"
         Height          =   195
         Left            =   7875
         TabIndex        =   38
         Top             =   6990
         Width           =   2385
      End
      Begin VB.Label Lbl_Empleados_12_Color 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   7500
         TabIndex        =   37
         Top             =   6960
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empleados Sin Checada E/S"
         Height          =   195
         Left            =   11460
         TabIndex        =   36
         Top             =   6990
         Width           =   2070
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   11085
         TabIndex        =   35
         Top             =   6960
         Width           =   315
      End
   End
End
Attribute VB_Name = "Frm_Adm_Validación_Tiempo_Trabajo"
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
        If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > -1 Then
            If Cmb_Adm_Validacion_Horas_Turno.ListIndex > -1 Then
                Fecha = Format(Dtp_Adm_Adm_Validacion_Horas_Fecha.Value, "MM/dd/yyyy")
                Generar_Lista
            Else
                MsgBox "Seleccione un turno", vbOKOnly + vbInformation, Me.Caption
                Cmb_Adm_Validacion_Horas_Turno.SetFocus
            End If
        Else
        If Es_RH Then
            Fecha = Format(Dtp_Adm_Adm_Validacion_Horas_Fecha.Value, "MM/dd/yyyy")
            Generar_Lista
        Else
            MsgBox "Seleccione un supervisor", vbOKOnly + vbInformation, Me.Caption
            Cmb_Adm_Validacion_Horas_Supervisor.SetFocus
            End If
        End If
    Else
        MsgBox "Seleccione una empresa", vbOKOnly + vbInformation, Me.Caption
        Cmb_Adm_Validacion_Horas_Empresa.SetFocus
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Es_RH
    'DESCRIPCIÓN:           Consulta los el área del usuario
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           25 Febrero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Es_RH() As Boolean
Dim Rs_Consulta_Area As rdoResultset       'Informacion de los registros


    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Cat_Areas.Nombre FROM Cat_Areas, Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Usuario_ID = " & Usuario_ID
    Mi_SQL = Mi_SQL & " AND Cat_Usuarios.Area_ID = Cat_Areas.Area_ID"
    Set Rs_Consulta_Area = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Es_RH = False
    With Rs_Consulta_Area
    
        If Not .EOF Then
           Dim Nombre As String
           Nombre = .rdoColumns("Nombre")
        End If
    End With
    If Nombre = "RECURSOS HUMANOS" Then
    Es_RH = True
    Else
    Es_RH = False
    End If
    'Cierra el manejador del registro
    Set Rs_Consulta_Area = Nothing

End Function


Private Sub Btn_Anterior_Click()
Dim Cont_Dias As Integer
Dim Dia_Valido As Boolean
    
'    Dia_Valido = False
    Cont_Dias = 1
    'Valida la fecha que se importara
'    Do While Dia_Valido = False
        Fecha = Format(DateAdd("d", -(Cont_Dias), Fecha), "MM/dd/yyyy")
        'Valida que no sea domingo ni sabado
'        If (Weekday(Fecha) <> vbSunday) Then
'            Dia_Valido = True
'        End If
'        Cont_Dias = Cont_Dias + 1
'    Loop
    Txt_Validacion_Horas_Grid.Visible = False
    Txt_Validacion_Justificacion_Horas_Grid.Visible = False
    'Generar_Lista
    Dtp_Adm_Adm_Validacion_Horas_Fecha.Value = Fecha
    Btn_Adm_Validacion_Horas_Generar_Click
End Sub

Private Sub Btn_Excel_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String

On Error GoTo HANDLER
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
            Call Exportar_Excel_Bien(Ruta_Temporal & Opcion & "xls.txt", Ruta_Exportacion)
        End If
    Else
        MsgBox "No existe información para exportar", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    MsgBox Err.Description
    Exit Sub
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

Private Sub Finalizar_Reporte()
    Close #1, #2
End Sub

Private Sub Btn_Imprimir_2_Click()
    Generar_Reporte
    Imprimir
End Sub

Private Sub Btn_Imprimir_Click()
    Generar_Reporte
    Imprimir
End Sub

Private Sub Btn_Regresar_Click()
    Pic_Adm_Validacion_Horas_Trabajo_Lista.Visible = False
    Pic_Adm_Validacion_Horas_Trabajo.Visible = True
    Me.Height = 4300
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

Private Sub Btn_Siguiente_Click()
Dim Cont_Dias As Integer
Dim Dia_Valido As Boolean
    
'    Dia_Valido = False
    Cont_Dias = 1
'    'Valida la fecha que se importara
'    Do While Dia_Valido = False
        Fecha = Format(DateAdd("d", (Cont_Dias), Fecha), "MM/dd/yyyy")
'        If Conectar_Ayudante.Es_Dia_No_Laboral(Fecha) = False Then
'            Dia_Valido = True
'        End If
'        Cont_Dias = Cont_Dias + 1
'    Loop
    Txt_Validacion_Horas_Grid.Visible = False
    Txt_Validacion_Justificacion_Horas_Grid.Visible = False
    'Generar_Lista
    Dtp_Adm_Adm_Validacion_Horas_Fecha.Value = Fecha
    Btn_Adm_Validacion_Horas_Generar_Click
End Sub

Private Sub Btn_Validar_Horas_Empleados_Click()
Dim Cont_Fila As Integer        'Recorrer el grid
Dim Guardar As Boolean          'Validar si existen registros que guardar
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 0 Then
        'Recorre el grid para saber si al menos se guardara algun registro
        For Cont_Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 19) = "SI" Then
                Guardar = True
                Exit For
            End If
        Next
        If Guardar = False Then
            MsgBox "No ha seleccionado información para validar", vbInformation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If MsgBox("¿La información es correcta?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            'Valida que la lista no se hay generado anteriormente
            If Fra_Validacion_Horas_Trabajo_Lista.Enabled = True Then
                Guardar_Lista
            Else
                MsgBox "La lista ya fue guardada, si desea realizar algun cambio" + vbCrLf + _
                       "deberá realizarlo en la opcion de Mantenimiento Asistencias ó" + vbCrLf + _
                       "Generar nuevamente la lista", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Chk_Seleccionar_Todas_Click()
Dim Fila As Integer     'Contador para recorrer el grid
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 0 Then
        For Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Fila, 19) <> "V" Then
                If Chk_Seleccionar_Todas.Value = 1 Then
                    Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Fila, 19) = "SI"
                Else
                    Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Fila, 19) = "NO"
                End If
            End If
        Next Fila
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Departamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID,Nombre", "Cat_Departamentos", Cmb_Adm_Validacion_Horas_Departamento, 0, "Nombre", "", True, "<-SELECCIONE->")
        If Cmb_Adm_Validacion_Horas_Departamento.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Departamento.ListIndex = 0
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Empresa_Click()
    If Cmb_Adm_Validacion_Horas_Empresa.ListIndex > -1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S' AND Empresa_ID = '" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "' ORDER BY Apellido_Paterno", Cmb_Adm_Validacion_Horas_Supervisor, 0, "")
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Estatus = 'A' AND Tipo='S'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Apellido_Paterno", "", False, "")
        If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
    End If
End Sub

Private Sub Cmb_Adm_Validacion_Horas_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Validacion_Horas_Empresa, 1, "Nombre")
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
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus = 'A' AND (Nombre like '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%' OR " & _
             "Apellido_Paterno like '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%' OR Apellido_Materno like '%" & Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text) & "%')", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Apellido_Paterno", "", False, "")
        'Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus = 'A'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "")
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

Private Sub Form_Activate()
    If Pic_Adm_Validacion_Horas_Trabajo.Visible = True Then Cmb_Adm_Validacion_Horas_Empresa.SetFocus
End Sub

Private Sub Form_Load()
Dim Dia_Valido As Boolean
Dim Cont_Dias As Integer
Dim Fecha_Generar As Date
Dim Rs_Empleados_Supervisor As rdoResultset
Dim Rs_Empleados_Departamento As rdoResultset

    'agrega la informacion en los combos
    Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Validacion_Horas_Empresa, 0, "Nombre")
    If Cmb_Adm_Validacion_Horas_Empresa.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Empresa.ListIndex = 0
    
    Cmb_Adm_Validacion_Horas_Supervisor.Text = "<-SELECCIONE->"
    If Rol_ID = 4 Then
        Call Conectar_Ayudante.Llena_Combo_Item("DISTINCT Supervisores.Empleado_ID, (Supervisores.Apellido_Paterno+' '+Supervisores.Apellido_Materno+' '+Supervisores.Nombre) as Nombre", "Cat_Empleados, Cat_Usuarios, Cat_Areas_Detalles, Cat_Empleados Supervisores WHERE Cat_Empleados.Estatus = 'A' AND Cat_Usuarios.Usuario_ID = '" & Usuario_ID & "' AND Cat_Usuarios.Area_ID = Cat_Areas_Detalles.Area_ID AND Cat_Areas_Detalles.Empleado_ID = Cat_Empleados.Empleado_ID AND Cat_Empleados.Supervisor_ID = Supervisores.Empleado_ID", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Cat_Empleados.Apellido_Paterno", "", False, "")
        If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then
            Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
            Cmb_Adm_Validacion_Horas_Supervisor.Enabled = False
        End If
    End If
    
'''    If Trim(Empleado_Supervisor_ID) = "" Then
'''        'Consulta Supervisor.
'''        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
'''        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
'''        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
'''        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'''        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
'''        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
'''        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
'''        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
'''        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'''        Cmb_Adm_Validacion_Horas_Supervisor.Clear
'''        While Not Rs_Empleados_Supervisor.EOF
'''            Cmb_Adm_Validacion_Horas_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
'''            Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
'''            Rs_Empleados_Supervisor.MoveNext
'''        Wend
'''        Rs_Empleados_Supervisor.Close
'''        Cmb_Adm_Validacion_Horas_Supervisor.Text = "<-SELECCIONE->"
'''
''''        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Apellido_Paterno", "", True, "<-SELECCIONE->")
'''        Cmb_Adm_Validacion_Horas_Supervisor.Enabled = True
'''    Else
'''        'Consulta Supervisor.
'''        Mi_SQL = "SELECT Cat_Areas_Detalles.Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) as Nombre "
'''        Mi_SQL = Mi_SQL & " ,(SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)as Supervisor"
'''        Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles,Cat_Empleados"
'''        Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
'''        Mi_SQL = Mi_SQL & " AND Not (SELECT TOP 1 Supervisor_ID from Cat_Empleados where Cat_Areas_Detalles.Empleado_id=Cat_Empleados.Supervisor_ID)   is null"
'''        Mi_SQL = Mi_SQL & " AND Tipo='S' AND Estatus = 'A'"
'''        Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
'''        Mi_SQL = Mi_SQL & " AND Empleado_ID='" & Empleado_Supervisor_ID & "'"
'''        Mi_SQL = Mi_SQL & " ORDER BY Apellido_Paterno"
'''        Set Rs_Empleados_Supervisor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'''        Cmb_Adm_Validacion_Horas_Supervisor.Clear
'''        While Not Rs_Empleados_Supervisor.EOF
'''            Cmb_Adm_Validacion_Horas_Supervisor.AddItem Rs_Empleados_Supervisor.rdoColumns("Nombre")
'''            Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.NewIndex) = Rs_Empleados_Supervisor.rdoColumns("Empleado_ID")
'''            Rs_Empleados_Supervisor.MoveNext
'''        Wend
'''        Rs_Empleados_Supervisor.Close
'''        Cmb_Adm_Validacion_Horas_Supervisor.Text = "<-SELECCIONE->"
'''
''''        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Tipo='S' AND Estatus='A' AND Empleado_ID='" & Empleado_Supervisor_ID & "'", Cmb_Adm_Validacion_Horas_Supervisor, 0, "Apellido_Paterno", "")
'''        Cmb_Adm_Validacion_Horas_Supervisor.Enabled = False
'''    End If
'    If Cmb_Adm_Validacion_Horas_Supervisor.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Supervisor.ListIndex = 0
    
    Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Adm_Validacion_Horas_Turno, 0, "Nombre", "", True, "<-SELECCIONE->")
    If Cmb_Adm_Validacion_Horas_Turno.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Turno.ListIndex = 0
    
    'Consulta Departamento.
    Mi_SQL = "SELECT DISTINCT Cat_Empleados.Departamento_ID,Cat_Departamentos.Nombre FROM Cat_Departamentos,Cat_Empleados,Cat_Areas_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Departamentos.Departamento_ID=Cat_Empleados.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Empleado_ID=Cat_Empleados.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND Area_ID ='" & Format(Area_ID, "00000") & "'"
    Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Cmb_Adm_Validacion_Horas_Departamento.Clear
    While Not Rs_Empleados_Departamento.EOF
        Cmb_Adm_Validacion_Horas_Departamento.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
        Cmb_Adm_Validacion_Horas_Departamento.ItemData(Cmb_Adm_Validacion_Horas_Departamento.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Departamento_ID")
        Rs_Empleados_Departamento.MoveNext
    Wend
    Rs_Empleados_Departamento.Close
    Cmb_Adm_Validacion_Horas_Departamento.Text = "<-SELECCIONE->"
    
'    Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID,Nombre", "Cat_Departamentos", Cmb_Adm_Validacion_Horas_Departamento, 0, "Nombre", "", True, "<-SELECCIONE->")
'    If Cmb_Adm_Validacion_Horas_Departamento.ListCount > 0 Then Cmb_Adm_Validacion_Horas_Departamento.ListIndex = 0
'
    Cont_Dias = 1
    'Valida la fecha que se importara
'    Do While Dia_Valido = False
'        Fecha_Generar = Format(DateAdd("d", -(Cont_Dias), Now), "MM/dd/yyyy")
'        'If Conectar_Ayudante.Es_Dia_No_Laboral(Fecha_Generar) = False Then
'            'Valida que no sea domingo ni sabado
'            'If (Weekday(Fecha_Generar) <> vbSunday) And (Weekday(Fecha_Generar) <> vbSaturday) Then
'            '    Dia_Valido = True
'            'End If
'        'End If
'        Cont_Dias = Cont_Dias + 1
'    Loop
'    Dtp_Adm_Adm_Validacion_Horas_Fecha.Value = Fecha_Generar
    Dtp_Adm_Adm_Validacion_Horas_Fecha.Value = Format(DateAdd("d", -1, Now), "MM/dd/yyyy")
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_Click()
    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 0 And .RowSel > 0 Then
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) <> "V" And PG_Calcula_Horas_Extra = "1" Then
                If .TextMatrix(.RowSel, 16) = "FI" Then
                    Txt_Validacion_Horas_Grid.Visible = False
                    Txt_Validacion_Justificacion_Horas_Grid.Visible = False
                    Exit Sub
                End If
                Txt_Validacion_Horas_Grid.Visible = False
                Txt_Validacion_Justificacion_Horas_Grid.Visible = False
                If Manejo_Grid = True Then
                    Select Case .ColSel
                        Case 12 'Horas extra
                            Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Validacion_Horas_Trabajo_Lista, Txt_Validacion_Horas_Grid)
                        Case 13 'Justificacion Horas extra
                            Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Validacion_Horas_Trabajo_Lista, Txt_Validacion_Justificacion_Horas_Grid)
                    End Select
                End If
            End If
        End If
    End With
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_DblClick()
Dim Frm_Movimientos_Permisos As New Frm_Adm_Solicitud_Permisos      'Forma de Permisos
Dim Frm_Movimientos_Incidencias As New Frm_Adm_Incidencias_Extraordinarias   'Forma de Incidencias extraordinarias
Dim Tipo_Inventario As String       'Tipo de inventario a consultar
Dim No_Movimiento As String         'No movimiento a consultar
Dim Cadena_Movimiento As String     'Cadena temporal con los datos de la referencia

    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 0 And .RowSel > 0 Then
            Txt_Validacion_Horas_Grid.Visible = False
            Txt_Validacion_Justificacion_Horas_Grid.Visible = False
            Select Case .ColSel
                Case 19
                    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) <> "V" Then
                        If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) = "SI" Then
                            Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) = "NO"
                        Else
                            Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) = "SI"
                        End If
                    End If
                Case 16, 20
                    'Obtiene el no. de movimiento y el tipo de incidencia
                    Cadena_Movimiento = Trim(.TextMatrix(.RowSel, 17))
                    Tipo_Inventario = Trim(.TextMatrix(.RowSel, 18))
                    No_Movimiento = Right(Cadena_Movimiento, 20)
                    If IsNumeric(No_Movimiento) = False Then Exit Sub
                    If Tipo_Inventario = "P" Then
                        If Conectar_Ayudante.Formulario_Cargado("SOLICITUD DE PERMISOS VE") Then
                            Conectar_Ayudante.Enfocar ("SOLICITUD DE PERMISOS VE")
                        Else
                            Load Frm_Movimientos_Permisos
                            Frm_Movimientos_Permisos.Top = 0
                            Call Conectar_Ayudante.Cargar_Picture(Frm_Movimientos_Permisos.Pic_Solicitud_Permisos, Frm_Movimientos_Permisos)
                            Frm_Movimientos_Permisos.Operacion = "Permisos"
                            Frm_Movimientos_Permisos.Caption = "SOLICITUD DE PERMISOS VE"
                            Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Validacion_Tiempo_Trabajo", Frm_Movimientos_Permisos)
                            Frm_Movimientos_Permisos.Inicializa
                            Frm_Movimientos_Permisos.Llenar_Informacion_Permiso (No_Movimiento)
                        End If
                    Else
                        If Tipo_Inventario = "E" Then
                            If Conectar_Ayudante.Formulario_Cargado("INCIDENCIAS EXTRAORDINARIAS VE") Then
                                Conectar_Ayudante.Enfocar ("INCIDENCIAS EXTRAORDINARIAS VE")
                            Else
                                Load Frm_Movimientos_Incidencias
                                'Frm_Permisos.Height = 3510
                                'Frm_Permisos.Width = 7080
                                Frm_Movimientos_Incidencias.Top = 0
                                Call Conectar_Ayudante.Cargar_Picture(Frm_Movimientos_Incidencias.Pic_Solicitud_Permisos, Frm_Movimientos_Incidencias)
                                Frm_Movimientos_Incidencias.Operacion = "Permisos"
                                Frm_Movimientos_Incidencias.Pic_Logo.Visible = True
                                Frm_Movimientos_Incidencias.Pic_Logo.ZOrder vbBringToFront
                                Frm_Movimientos_Incidencias.Caption = "INCIDENCIAS EXTRAORDINARIAS VE"
                                Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Adm_Validacion_Tiempo_Trabajo", Frm_Movimientos_Incidencias)
                                Frm_Movimientos_Incidencias.Inicializa
                                Frm_Movimientos_Incidencias.Llenar_Informacion_Permiso (No_Movimiento)
                            End If
                        End If
                    End If
            End Select
        End If
    End With
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_EnterCell()
    Grid_Validacion_Horas_Trabajo_Lista_Click
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_LeaveCell()
    'Grid_Validacion_Horas_Trabajo_Lista.CellBackColor = vbWhite
End Sub

Private Sub Grid_Validacion_Horas_Trabajo_Lista_Scroll()
    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.RowSel, 19) <> "V" And PG_Calcula_Horas_Extra = "1" Then
        If Grid_Validacion_Horas_Trabajo_Lista.ColSel = 12 Then
            Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Validacion_Horas_Trabajo_Lista, Txt_Validacion_Horas_Grid)
        ElseIf Grid_Validacion_Horas_Trabajo_Lista.ColSel = 13 Then
            Call Conectar_Ayudante.Mover_Control_Grid_TextBox(Grid_Validacion_Horas_Trabajo_Lista, Txt_Validacion_Justificacion_Horas_Grid)
        End If
    End If
End Sub

Private Sub Txt_Validacion_Horas_Grid_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Solo_Numeros(KeyAscii, Txt_Validacion_Horas_Grid, True)
End Sub

Private Sub Txt_Validacion_Horas_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
        'Guarda la informacion y oculta el check
        Txt_Validacion_Horas_Grid.Visible = False
        Call Mover_Control_Grid_Procesos(KeyCode)
        If KeyCode = 37 Or KeyCode = 39 Then Grid_Validacion_Horas_Trabajo_Lista.SetFocus
    End If
End Sub

Private Sub Txt_Validacion_Horas_Grid_Change()
Dim Columna As Integer
    With Grid_Validacion_Horas_Trabajo_Lista
        If .RowSel > 0 Then
            If .TextMatrix(.RowSel, 16) = "FI" Then
                Txt_Validacion_Horas_Grid.Visible = False
                Exit Sub
            End If
            .TextMatrix(.RowSel, 12) = Txt_Validacion_Horas_Grid.Text
            If .TextMatrix(.RowSel, 16) = "HI" And Val(.TextMatrix(.RowSel, 10)) >= Val(.TextMatrix(.RowSel, 22)) Then
                .TextMatrix(.RowSel, 16) = "AS"
                .TextMatrix(.RowSel, 15) = ""
                Manejo_Grid = False
                .Col = 0
                For Columna = 1 To .Cols - 1
                    .Col = Columna
                    .Row = .RowSel
                    .CellBackColor = &HFFFFFF
                Next Columna
                Manejo_Grid = True
                .Col = 10
                .Row = .RowSel
                SendKeys "{END}"
            End If
            If .TextMatrix(.RowSel, 16) = "AS" And Val(.TextMatrix(.RowSel, 10)) < Val(.TextMatrix(.RowSel, 22)) Then
                .TextMatrix(.RowSel, 16) = "HI"
                .TextMatrix(.RowSel, 15) = "Horas_Incompletas"
                Manejo_Grid = False
                .Col = 0
                For Columna = 1 To .Cols - 1
                    .Col = Columna
                    .Row = .RowSel
                    .CellBackColor = &H80FFFF
                Next Columna
                Manejo_Grid = True
                .Col = 10
                .Row = .RowSel
                SendKeys "{END}"
            End If
        End If
    End With
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
'FECHA_CREO : 08-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Generar_Lista()
Dim Consulta_Cat_Empleados As rdoResultset                      'Informacion de la lista de empleados
Dim Rs_Consulta_Adm_Asistencias_Detalles As rdoResultset        'Obtiene la asistencia
Dim Rs_Consulta_Adm_Permisos As rdoResultset                    'Informacion de permisos del empleado
Dim Rs_Consulta_Adm_Vacaciones As rdoResultset                  'Informacion de Vacaciones del empleado
Dim Rs_Consulta_Informacion_Turnos As rdoResultset              'Informacion del empleado
Dim Rs_Consulta_Dia_Feriado As rdoResultset                     'Informacion del empleado
Dim Rs_Consulta_Cat_Empleados As rdoResultset                   'Informacion del empleado
Dim Rs_Consulta_Adm_Asistencias_Validadas As rdoResultset
Dim Rs_Consulta_Rutas_Transportes As rdoResultset
Dim Rs_Consulta_Checada As rdoResultset
Dim Bool_Proceso As Boolean
Dim Permiso As String                                           'Informacion del permiso
Dim Tipo_Incidencia As String                                   'Tipo de incidencia generada por el ausentismo
Dim Partida As Integer                                          'No consecutivo de la lista
Dim Horas_Laboradas As Double                                   'Obtiene las horas laboradas del empleado
Dim Hora_Entrada As String                                      'Obtiene la hora de entrada del empleado
Dim Hora_Entrada_Auxiliar As Date                                      'Obtiene la hora de entrada del empleado
Dim Hora_Salida_Auxiliar As Date
Dim Hora_Entrada_Calculo As String                              'Obtiene la hora de entrada del empleado
Dim Hora_Salida As String                                       'Obtiene la hora de salida del empleado
Dim Hora_Inicio_Turno As Date                                   'Obtiene la hora de inicio del turno del empleado
Dim Hora_Termino_Turno As Date                                  'Obtiene la hora de termino del turno del empleado
Dim Hora_Comida_Salida As Date                                  'Obtiene la hora de inicio del turno del empleado
Dim Hora_Comida_Entrada As Date                                 'Obtiene la hora de termino del turno del empleado
Dim Horas_Turno As Double                                       'Obtiene las horas del turno del empleado
Dim Turno_Empleado As String                                    'Identificador del turno del empleado
Dim Turno_ID As String
Dim Calendario_Turno_ID As String
Dim Calendario_Turno_Detalle_ID As String
Dim Nombre_Turno As String
Dim Referencia As String                                        'No Referencia de la incidencia
Dim Simbologia As String                                        'Simbologia de la incidencia
Dim SubSimbologia As String                                     'Subsimbologia de la incidencia
Dim Hora_Real As Double                                         'Horas real trabajadas
Dim Hora_Real_Extra As Double
Dim No_Movimiento As Double                                     'No de movimiento
Dim Validada As String                                          'Indica si la asistencia ya fue validada
Dim Empresa_Sindicalizada As Boolean                            'Identifica si la empresa es sindicalizada
Dim Colorear_Fila As Boolean                                    'Define si la fila se coloreara o no
Dim Columna As Integer
Dim Hora_Comida_Inicio As Date
Dim Hora_Comida_Termino As Date
Dim Horas_Extra As Double
Dim Horas_Extra_Definidas As Double
Dim Horas_Extra_Paga As Double
Dim Horas_Extra_Adicionales As Double
Dim Horas_Laboradas_Turno As Double
Dim Justificacion_Horas_Extra As String
Dim Nombre_Supervisor As String
Dim Permiso_Validado As String
Dim Permiso_Referencia As String
Dim Ruta_Transporte As String
Dim Dia_Descanso As String

On Error GoTo HANDLER:
    Partida = 0
    Grid_Validacion_Horas_Trabajo_Lista.Rows = 0
    Grid_Validacion_Horas_Trabajo_Lista.Cols = 27
    'Informacion para la barra de progreso
    Mi_SQL = "SELECT COUNT(CE.No_Tarjeta) AS Empleados"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE"
    Mi_SQL = Mi_SQL & " WHERE CE.Estatus='A'"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
    
    If Trim(Empleado_Supervisor_ID) = "" Then
        If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
        Else
            If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > -1 And Cmb_Adm_Validacion_Horas_Supervisor.Text <> "<-SELECCIONE->" Then
                Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
            End If
        End If
    Else
        If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
        End If
    End If
    
    If Cmb_Adm_Validacion_Horas_Turno.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Turno_ID='" & Format(Cmb_Adm_Validacion_Horas_Turno.ItemData(Cmb_Adm_Validacion_Horas_Turno.ListIndex), "00000") & "'"
    End If
    If Cmb_Adm_Validacion_Horas_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Adm_Validacion_Horas_Departamento.ItemData(Cmb_Adm_Validacion_Horas_Departamento.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND CE.Fecha_Ingreso<=" & Par_Fecha & Format(Fecha, "MM/dd/yyyy") & Par_Fecha
    Set Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Consulta_Cat_Empleados.EOF Then
        'Obtiene la informacion para configurar el progress bar
        If Val(Consulta_Cat_Empleados.rdoColumns("Empleados")) > 0 Then
            PrgBar_Validacion_Horas.Max = Val(Consulta_Cat_Empleados.rdoColumns("Empleados"))
        End If
    End If
    Consulta_Cat_Empleados.Close
    Set Consulta_Cat_Empleados = Nothing
    'Informacion para la lista
    Mi_SQL = "SELECT CE.Empleado_ID,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) AS Nombre,CE.No_Tarjeta,CE.Turno_ID,CE.Supervisor_ID,CE.Tipo_Empleado,CE.Cedula_Identidad_Ciudadana"
    Mi_SQL = Mi_SQL & " ,CD.Nombre AS Departamento,CD.Clave,CE.Transporte_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Departamentos CD"
    Mi_SQL = Mi_SQL & " WHERE CE.Departamento_ID=CD.Departamento_ID"
    Mi_SQL = Mi_SQL & " AND CE.Estatus='A'"
    Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
    If Trim(Empleado_Supervisor_ID) = "" Then
        If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > 0 Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
        Else
            If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > -1 And Cmb_Adm_Validacion_Horas_Supervisor.Text <> "<-SELECCIONE->" Then
                Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
            End If
        End If
    Else
        If Cmb_Adm_Validacion_Horas_Supervisor.ListIndex > -1 Then
            Mi_SQL = Mi_SQL & " AND CE.Supervisor_ID='" & Format(Cmb_Adm_Validacion_Horas_Supervisor.ItemData(Cmb_Adm_Validacion_Horas_Supervisor.ListIndex), "00000") & "'"
        End If
    End If
    If Cmb_Adm_Validacion_Horas_Turno.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Turno_ID='" & Format(Cmb_Adm_Validacion_Horas_Turno.ItemData(Cmb_Adm_Validacion_Horas_Turno.ListIndex), "00000") & "'"
    End If
    If Cmb_Adm_Validacion_Horas_Departamento.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND CE.Departamento_ID='" & Format(Cmb_Adm_Validacion_Horas_Departamento.ItemData(Cmb_Adm_Validacion_Horas_Departamento.ListIndex), "00000") & "'"
    End If
    Mi_SQL = Mi_SQL & " AND CE.Fecha_Ingreso<='" & Format(Fecha, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " ORDER BY CE.No_Tarjeta"
    Set Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Consulta_Cat_Empleados.EOF Then
        With Consulta_Cat_Empleados
            Me.MousePointer = 11
            Me.Refresh
            PrgBar_Validacion_Horas.Visible = True
            PrgBar_Validacion_Horas.Value = 0
            Empresa_Sindicalizada = False
            'Identifica si la empresa es o no sindicalizada
            If InStr(1, Cmb_Adm_Validacion_Horas_Empresa.Text, "SINDI") > 0 Then
                Empresa_Sindicalizada = True
            End If
            Call Encabezado_Reporte("VALIDACION DE HORAS", DateAdd("s", 1, Dtp_Adm_Adm_Validacion_Horas_Fecha.Value), DateAdd("s", 1, Dtp_Adm_Adm_Validacion_Horas_Fecha.Value))
            'Agrega el encabezado
            Grid_Validacion_Horas_Trabajo_Lista.AddItem "Referencia" _
                & Chr(9) & "Empleado_ID" _
                & Chr(9) & "No." _
                & Chr(9) & "Departamento" _
                & Chr(9) & "Nombre" _
                & Chr(9) & "Turno" _
                & Chr(9) & "Entrada" _
                & Chr(9) & "Comida S" _
                & Chr(9) & "Comida E" _
                & Chr(9) & "Salida" _
                & Chr(9) & "Hrs" _
                & Chr(9) & "Acuerdo" _
                & Chr(9) & "Extra" _
                & Chr(9) & "Justificacion Hrs. Extra" _
                & Chr(9) & "Sistema" _
                & Chr(9) & "Observaciones" _
                & Chr(9) & "Tipo" _
                & Chr(9) & "SubTipo" _
                & Chr(9) & "No_Detalle" _
                & Chr(9) & "Validar" _
                & Chr(9) & "No Movimiento" _
                & Chr(9) & "Mov." _
                & Chr(9) & "Horas_Turno" _
                & Chr(9) & "TurnoID" & Chr(9) & "Fecha y Usuario Validó" _
                & Chr(9) & "CalendarioTurnoID" & Chr(9) & "CalendarioTurnoDetalleID"
            Print #1, ""
            Print #2, "No.|Departamento|Empleado|Supervisor|Entrada|S.Comida|E.Comida|Salida|Horas|Hrs.Acuerdo|Hrs.Extra|Justificacion Hrs. Extra|Calculadas|Turno|Observaciones|Tipo|No.Detalle|Departamento|Depto|MO|Subdivision|Ruta|Transporte"
            Partida = 0
            
            
            '   consulta las horas de asistencia de cada persona
            While Not .EOF
                Colorear_Fila = False
                Permiso = ""
                Hora_Entrada = "0"
                Hora_Salida = "0"
                Horas_Laboradas = 0
                Hora_Real = 0
                Tipo_Incidencia = ""
                Simbologia = ""
                SubSimbologia = ""
                Validada = ""
                'Valores para las incindecias
                Referencia = ""
                No_Movimiento = 0
                
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                'Consulta la asistencia del empleado
                Mi_SQL = "SELECT * FROM Adm_Asistencias_Detalles"
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
                Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Fecha, "MM/dd/yyyy") & "'"
                Set Rs_Consulta_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Adm_Asistencias_Detalles.EOF Then
                    Simbologia = "A"
                    SubSimbologia = ""
                    Validada = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Validada")
                    No_Movimiento = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("No_Operacion")
                    Hora_Entrada = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Entrada")
                    'Hora_Comida_Entrada = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Comida_Entrada")
                    Hora_Entrada_Calculo = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Entrada")
                    Hora_Comida_Entrada = "0"
                    Hora_Entrada_Calculo = "0"
                    Hora_Salida = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Salida")
                    'If Not IsNull(Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Comida_Salida")) Then
                    '    Hora_Comida_Salida = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Hora_Comida_Salida")
                    'End If
                    If Not IsNull(Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Justificacion_Horas_Extra")) Then
                    Justificacion_Horas_Extra = Rs_Consulta_Adm_Asistencias_Detalles.rdoColumns("Justificacion_Horas_Extra")
                    Else
                    Justificacion_Horas_Extra = ""
                    End If
                    If (DateDiff("n", Hora_Salida, Hora_Entrada) / 60) > 0 Then
                        Hora_Real = DateDiff("n", Hora_Salida, Hora_Entrada) / 60
                    Else
                        Hora_Real = 24 + (DateDiff("n", Hora_Salida, Hora_Entrada) / 60)
                    End If
                    'Hora_Real = Hora_Real - DateDiff("n", Hora_Comida_Salida, Hora_Comida_Entrada) / 60
                    Horas_Laboradas = Hora_Real
                Else
                    Simbologia = "F"
                    SubSimbologia = ""
                    Validada = "N"
                End If
                Rs_Consulta_Adm_Asistencias_Detalles.Close
                Me.Refresh
                
                
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                'Consulta el nombre del supervisor
                If Not IsNull(.rdoColumns("Supervisor_ID")) Then
                    Mi_SQL = "SELECT (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Supervisor FROM Cat_Empleados"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Supervisor_ID") & "'"
                    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Empleados.EOF Then
                        Nombre_Supervisor = Rs_Consulta_Cat_Empleados.rdoColumns("Supervisor")
                    Else
                        Nombre_Supervisor = ""
                    End If
                    Rs_Consulta_Cat_Empleados.Close
                Else
                    Nombre_Supervisor = ""
                End If
                
                
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------
                'Consulta la zona y transporte
                If Not IsNull(.rdoColumns("Transporte_ID")) Then
                    Mi_SQL = "SELECT Cat_Zonas.Nombre AS Ruta,Cat_Transportes.Nombre AS Transporte"
                    Mi_SQL = Mi_SQL & " FROM Cat_Zonas,Cat_Transportes"
                    Mi_SQL = Mi_SQL & " WHERE Cat_Zonas.Zona_ID=Cat_Transportes.Zona_ID "
                    Mi_SQL = Mi_SQL & " AND Cat_Transportes.Transporte_ID='" & .rdoColumns("Transporte_ID") & "'"
                    Set Rs_Consulta_Rutas_Transportes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Rutas_Transportes.EOF Then
                        Ruta_Transporte = Rs_Consulta_Rutas_Transportes.rdoColumns("Ruta") & "|" & Rs_Consulta_Rutas_Transportes.rdoColumns("Transporte")
                    Else
                        Ruta_Transporte = "" & "|" & ""
                    End If
                    Rs_Consulta_Rutas_Transportes.Close
                Else
                    Ruta_Transporte = "" & "|" & ""
                End If
                If Validada = "N" Then
                    
                    '--------------------------------------------------------------------------------------------
                    '--------------------------------------------------------------------------------------------
                    '--------------------------------------------------------------------------------------------
                    '--------------------------------------------------------------------------------------------
                    'Obtiene la informacion del turno del empleado
                    Mi_SQL = "SELECT CE.Turno_ID,Cat_Turnos_Detalles.Hora_Inicio,Cat_Turnos_Detalles.Hora_Termino,Cat_Turnos_Detalles.Comida_Inicio,Cat_Turnos_Detalles.Comida_Termino,Cat_Turnos_Detalles.Horas_Turno,Cat_Turnos_Detalles.Horas_Comida,Cat_Turnos_Detalles.Dia_Descanso,CT.Nombre"
                    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Turnos CT,Cat_Turnos_Detalles"
                    Mi_SQL = Mi_SQL & " WHERE CE.Turno_ID=CT.Turno_ID"
                    Mi_SQL = Mi_SQL & " AND CT.Turno_ID=Cat_Turnos_Detalles.Turno_ID"
                    Mi_SQL = Mi_SQL & " AND CE.Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
                    If UCase(Format(Fecha, "dddd")) = "SABADO" Or UCase(Format(Fecha, "dddd")) = "SÁBADO" Or UCase(Format(Fecha, "dddd")) = "SATURDAY" Then
                        Mi_SQL = Mi_SQL & " AND Cat_Turnos_Detalles.Dia_Semana IN ('Sabado')"
                    Else
                        If UCase(Format(Fecha, "dddd")) = "MIERCOLES" Or UCase(Format(Fecha, "dddd")) = "MIÉRCOLES" Or UCase(Format(Fecha, "dddd")) = "WEDNESDAY" Then
                            Mi_SQL = Mi_SQL & " AND Cat_Turnos_Detalles.Dia_Semana IN ('Miercoles')"
                        Else
                            Mi_SQL = Mi_SQL & " AND Cat_Turnos_Detalles.Dia_Semana='" & Format(Fecha, "dddd") & "'"
                        End If
                    End If
                    Mi_SQL = Mi_SQL & " AND NOT EXISTS ("
                    Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
                    Mi_SQL = Mi_SQL & "     FROM ("
                    Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
'                    Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
                    Mi_SQL = Mi_SQL & "         FROM Cat_Calendarios_Turnos"
                    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
                    Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
                    Mi_SQL = Mi_SQL & "         WHERE Cat_Calendarios_Turnos_Roles.No_Tarjeta = CE.No_Tarjeta"
                    Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) = '" & Format(Fecha, "YYYYMMDD") & "'"
                    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
                    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
                    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
                    Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
'                    Mi_SQL = Mi_SQL & "             AND (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) >=0"
                    Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
'                    Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
'                    Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = '" & Format(Fecha, "YYYYMMDD") & "'"
                    Mi_SQL = Mi_SQL & " )"
                    Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    
                    If Not Rs_Consulta_Informacion_Turnos.EOF Then
                        Turno_Empleado = Rs_Consulta_Informacion_Turnos.rdoColumns("Turno_ID")
                        Nombre_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Nombre")
                        Hora_Inicio_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                        Hora_Termino_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Termino"), "HH:mm:ss")
                        Hora_Comida_Inicio = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                        Hora_Comida_Termino = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Termino"), "HH:mm:ss")
                        Horas_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Horas_Turno")
                        Horas_Laboradas_Turno = Horas_Turno
                        Dia_Descanso = Rs_Consulta_Informacion_Turnos.rdoColumns("Dia_Descanso")
                        
                        Bool_Proceso = False
                        
                        '   se actualizaran los tiempos si se trata del 3ro turno con id '00003'
                        If Turno_Empleado = "00003" Then
                        
                            'Consulta la hora de entrada del día
                            Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,"
                            Mi_SQL = Mi_SQL & " Adm_Asistencias_Registro_Checadores.Hora"
                            Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                            Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
                            Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Fecha & "'"
                            Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora>'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                            Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                            Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                            Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                            Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            
                            '   se consulta la informacion de las horas del checador
                            If Not Rs_Consulta_Checada.EOF Then
                            
                                Hora_Entrada_Auxiliar = Rs_Consulta_Checada.rdoColumns("Hora")
                            
                                If (Hora_Entrada_Auxiliar >= "22:00:00") Then
                                    Hora_Entrada = Rs_Consulta_Checada.rdoColumns("Hora")
                                    Hora_Salida = Rs_Consulta_Checada.rdoColumns("Hora")
                                    Bool_Proceso = True
                                End If
                            
                               
                            End If
                            Rs_Consulta_Checada.Close
                            
                            
                            If (Bool_Proceso = True) Then
                                'Consulta la hora de salida del día siguiente
                                 Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                                 Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                                 Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Format(DateAdd("d", 1, Fecha), "MM/dd/yyyy") & "'"
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora<'1899-12-30 12:00:00.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                                 Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                                 Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                                 Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                 If Not Rs_Consulta_Checada.EOF Then
                                     If Hora_Entrada = "0" Then 'Si no tuvo hora de entrada le asigna su entrada coo su salida y no le calcula tiempo extra
                                         Hora_Entrada = Rs_Consulta_Checada.rdoColumns("Hora")
                                     End If
                                     Hora_Salida = Rs_Consulta_Checada.rdoColumns("Hora")
                                 End If
                                 Rs_Consulta_Checada.Close
                                       
                            '-----------------------------------------------------------
                            Else
                            '-----------------------------------------------------------
                                'Consulta la hora de salida del día actual
                                 Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta"
                                 Mi_SQL = Mi_SQL & ", Cast(Adm_Asistencias_Registro_Checadores.Hora as DATETIME) as Hora"
                                 Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                                 Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='" & Format(Cmb_Adm_Validacion_Horas_Empresa.ItemData(Cmb_Adm_Validacion_Horas_Empresa.ListIndex), "00000") & "'"
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Fecha='" & Format(Fecha, "MM/dd/yyyy") & "'"
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.Hora<'1899-12-30 23:59:59.000'"       'Debe ser menor de las 12 hrs. del día siguiente para cerrar el ciclo de un día
                                 Mi_SQL = Mi_SQL & " AND Adm_Asistencias_Registro_Checadores.No_Tarjeta='" & .rdoColumns("No_Tarjeta") & "'"
                                 Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
                                 Mi_SQL = Mi_SQL & " ORDER BY Adm_Asistencias_Registro_Checadores.Fecha,Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora desc"
                                 Set Rs_Consulta_Checada = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                 If Not Rs_Consulta_Checada.EOF Then
                                     If Hora_Entrada = "0" Then 'Si no tuvo hora de entrada le asigna su entrada coo su salida y no le calcula tiempo extra
                                         Hora_Entrada = Rs_Consulta_Checada.rdoColumns("Hora")
                                     End If
                                     
                                     Hora_Salida = CDate(Rs_Consulta_Checada.rdoColumns("Hora"))
                                     Hora_Salida_Auxiliar = Rs_Consulta_Checada.rdoColumns("Hora")
                                 End If
                                 Rs_Consulta_Checada.Close
                                       
                            End If
                            
                           
                          
                                  
                        End If
                        
                        
'                        '   se calculan los nuevos totales
'                        If (DateDiff("n", Hora_Salida, Hora_Entrada) / 60) > 0 Then
'                            Hora_Real = DateDiff("n", Format(Hora_Salida, "HH:mm:ss"), Format(Hora_Entrada, "HH:mm:ss")) / 60
'
'                            Dim x As Integer
'
'                            'Format(Datos(1), "MM/dd/yyyy")
'                            x = DateDiff("h", Format(Hora_Salida, "HH:mm:ss"), Format(Hora_Entrada, "HH:mm:ss"))
'
'                        Else
'                            Hora_Real = 24 + (DateDiff("n", Hora_Salida, Hora_Entrada) / 60)
'                        End If
                               
                               
                    Else    'Si no encuentra día de la semana busca en el general
                        Rs_Consulta_Informacion_Turnos.Close
                        Mi_SQL = "SELECT CE.Turno_ID,CT.Hora_Inicio,CT.Hora_Termino,CT.Comida_Inicio,CT.Comida_Termino,CT.Horas_Turno,CT.Horas_Comida,CT.Nombre"
                        Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Cat_Turnos CT"
                        Mi_SQL = Mi_SQL & " WHERE CE.Turno_ID=CT.Turno_ID"
                        Mi_SQL = Mi_SQL & " AND CE.Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
                        Mi_SQL = Mi_SQL & " AND NOT EXISTS ("
                        Mi_SQL = Mi_SQL & "     SELECT Roles_Calendarios.No_Tarjeta"
                        Mi_SQL = Mi_SQL & "     FROM ("
                        Mi_SQL = Mi_SQL & "         SELECT Cat_Calendarios_Turnos_Roles.No_Tarjeta"
'                        Mi_SQL = Mi_SQL & "             ,DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) Fecha_Calculada"
                        Mi_SQL = Mi_SQL & "         From Cat_Calendarios_Turnos"
                        Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Detalles"
                        Mi_SQL = Mi_SQL & "             ,Cat_Calendarios_Turnos_Roles"
                        Mi_SQL = Mi_SQL & "         WHERE Cat_Calendarios_Turnos_Roles.No_Tarjeta = CE.No_Tarjeta"
                        Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
                        Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
                        Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
                        Mi_SQL = Mi_SQL & "             AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
                        Mi_SQL = Mi_SQL & "             AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) = '" & Format(Fecha, "YYYYMMDD") & "'"
'                        Mi_SQL = Mi_SQL & "             AND (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) >=0"
                        Mi_SQL = Mi_SQL & "         ) Roles_Calendarios"
'                        Mi_SQL = Mi_SQL & "     WHERE Roles_Calendarios.No_Tarjeta = CE.No_Tarjeta"
'                        Mi_SQL = Mi_SQL & "     AND Roles_Calendarios.Fecha_Calculada = '" & Format(Fecha, "YYYYMMDD") & "'"
                        Mi_SQL = Mi_SQL & " )"
                        Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Informacion_Turnos.EOF Then
                            Turno_Empleado = Rs_Consulta_Informacion_Turnos.rdoColumns("Turno_ID")
                            Nombre_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Nombre")
                            Hora_Inicio_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                            Hora_Termino_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Termino"), "HH:mm:ss")
                            Hora_Comida_Inicio = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                            Hora_Comida_Termino = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Termino"), "HH:mm:ss")
                            Horas_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Horas_Turno")
                            Horas_Laboradas_Turno = Horas_Turno
                            Dia_Descanso = "NO"
                        Else
                            Mi_SQL = "SELECT Cat_Calendarios_Turnos.Calendario_Turno_ID"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Roles.No_Tarjeta"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Hora_Inicio"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Hora_Termino"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Comida_Inicio"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Comida_Termino"
                            Mi_SQL = Mi_SQL & "     ,CASE WHEN (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) >= 0 THEN (Cat_Calendarios_Turnos_Detalles.Hora_Termino - Cat_Calendarios_Turnos_Detalles.Hora_Inicio) ELSE (Cat_Calendarios_Turnos_Detalles.Hora_Inicio - Cat_Calendarios_Turnos_Detalles.Hora_Termino) END AS Horas_Turno"
                            Mi_SQL = Mi_SQL & "     ,CASE WHEN (Cat_Calendarios_Turnos_Detalles.Comida_Termino - Cat_Calendarios_Turnos_Detalles.Comida_Inicio) >=0 THEN (Cat_Calendarios_Turnos_Detalles.Comida_Termino - Cat_Calendarios_Turnos_Detalles.Comida_Inicio) ELSE (Cat_Calendarios_Turnos_Detalles.Comida_Inicio - Cat_Calendarios_Turnos_Detalles.Comida_Termino) END AS Horas_Comida"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles.Nombre_Turno"
                            Mi_SQL = Mi_SQL & " FROM Cat_Calendarios_Turnos"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Detalles"
                            Mi_SQL = Mi_SQL & "     ,Cat_Calendarios_Turnos_Roles"
                            Mi_SQL = Mi_SQL & "     ,Cat_Empleados"
                            Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
                            Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos_Roles.No_Tarjeta = Cat_Empleados.No_Tarjeta"
                            Mi_SQL = Mi_SQL & "     AND DATEADD(DAY, dbo.Obtener_Numero_Dia_Semana(Cat_Calendarios_Turnos_Detalles.Dia_Semana) - 1, DATEADD(WEEK, Cat_Calendarios_Turnos_Detalles.Semana - 1, CAST(YEAR(Cat_Calendarios_Turnos.Fecha_Inicio) AS VARCHAR) + '0101')) = '" & Format(Fecha, "YYYYMMDD") & "'"
                            Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos_Detalles.Estatus <> 'ELIMINADO'"
                            Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos.Calendario_Turno_ID = Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID"
                            Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_ID"
                            Mi_SQL = Mi_SQL & "     AND Cat_Calendarios_Turnos_Detalles.Calendario_Turno_Detalle_ID = Cat_Calendarios_Turnos_Roles.Calendario_Turno_Detalle_ID"
                            Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Informacion_Turnos.EOF Then
                                Calendario_Turno_ID = Rs_Consulta_Informacion_Turnos.rdoColumns("Calendario_Turno_ID")
                                Calendario_Turno_Detalle_ID = Rs_Consulta_Informacion_Turnos.rdoColumns("Calendario_Turno_Detalle_ID")
                                Nombre_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Nombre_Turno")
                                Hora_Inicio_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                                Hora_Termino_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Termino"), "HH:mm:ss")
                                Hora_Comida_Inicio = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                                Hora_Comida_Termino = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Termino"), "HH:mm:ss")
                                Horas_Turno = Rs_Consulta_Informacion_Turnos.rdoColumns("Horas_Turno")
                                Horas_Laboradas_Turno = Horas_Turno
                                Dia_Descanso = "NO"
                            End If
                        End If
                    End If
                    Rs_Consulta_Informacion_Turnos.Close
                    Set Rs_Consulta_Informacion_Turnos = Nothing
                    Me.Refresh
                    Hora_Real = 0
                    Hora_Real_Extra = 0
                    If Simbologia = "A" Then
                        If DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Entrada_Calculo, "HH:mm:ss")) <= 0 Then
                            Hora_Entrada_Calculo = Hora_Inicio_Turno
                        End If
                        'Valida si el horario de salida es un día posterior para evitar horas negativas
                        If DateDiff("n", Hora_Inicio_Turno, Hora_Termino_Turno) > 0 Then
                            If Hora_Entrada <> "01/01/1900" Then    'Valida si tuvo hora de entrada
                                Hora_Real = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                If DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Entrada_Calculo, "HH:mm:ss")) <= 0 Then
'                                    Hora_Real_Extra = (DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                Else
                                    Hora_Real_Extra = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                End If
                            End If
                            Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
                        Else
                            'Valida si salió el mismo día que entró para el segundo turno
                            If DateDiff("n", Hora_Entrada, Hora_Salida) > 0 Then
                                If Hora_Entrada <> "01/01/1900" Then    'Valida si tuvo hora de entrada
                                    Hora_Real = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                    If DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Entrada_Calculo, "HH:mm:ss")) <= 0 Then
'                                        Hora_Real_Extra = (DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                    Else
                                        Hora_Real_Extra = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
'                                    End If
                                End If
                                Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
                            Else
                                If Hora_Entrada <> "01/01/1900" Then    'Valida si tuvo hora de entrada
                                
                                    If (Bool_Proceso = True) Then
                                        Hora_Real = DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), "23:59:59") / 60 + DateDiff("n", "00:00:00", Format(Hora_Salida, "HH:mm:ss")) / 60
                                        Hora_Real_Extra = DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), "23:59:59") / 60 + DateDiff("n", "00:00:00", Format(Hora_Salida, "HH:mm:ss")) / 60
                                    Else
                                        Hora_Real = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
                                        Hora_Real_Extra = (DateDiff("n", Format(Hora_Entrada, "HH:mm:ss"), Format(Hora_Salida, "HH:mm:ss"))) / 60
                                    End If
                                    
                                
                                    
'                                    If DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), Format(Hora_Entrada_Calculo, "HH:mm:ss")) <= 0 Then
'                                        Hora_Real_Extra = DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), "23:59:59") / 60 + DateDiff("n", "00:00:00", Format(Hora_Salida, "HH:mm:ss")) / 60
'                                    Else
                                        
'                                    End If
                                End If
                                Horas_Turno = DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), "23:59:59") / 60 + DateDiff("n", "00:00:00", Format(Hora_Termino_Turno, "HH:mm:ss")) / 60
                            End If
                        End If
                        'Horas_Turno = Horas_Turno - (Abs(DateDiff("n", Format(Hora_Comida_Termino, "HH:mm:ss"), Format(Hora_Comida_Inicio, "HH:mm:ss"))) / 60)
                        Horas_Laboradas = Format(Horas_Turno, "#0.0")
                        'Valida si calcula retardo
                        If PG_Aplica_Retardos = "1" Then
                            If DateDiff("n", Format(Hora_Entrada, "HH:mm"), DateAdd("n", PG_Tolerancia_Retardos, Format(Hora_Inicio_Turno, "HH:mm"))) < 0 Then
                                Permiso = "Retardo"
                                Simbologia = "RE"
                                SubSimbologia = ""
                            End If
                        End If
                        If DateDiff("n", Hora_Termino_Turno, Hora_Salida) < 0 Then
                            Permiso = "Registro de salida antes de horario"
                            'Horas_Laboradas = Hora_Real
                            Simbologia = "HI"
                            SubSimbologia = ""
                        End If
                        'If DateDiff("n", Hora_Termino_Turno, Hora_Salida) > 0 And Empresa_Sindicalizada = True Then
                        If Empresa_Sindicalizada = True Then
                            If Hora_Real >= (DateDiff("n", Hora_Inicio_Turno, Hora_Termino_Turno) / 60) Then
                                Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
                                'Horas_Turno = Horas_Turno - (Abs(DateDiff("n", Format(Hora_Comida_Termino, "HH:mm:ss"), Format(Hora_Comida_Inicio, "HH:mm:ss"))) / 60)
                                Horas_Laboradas = Horas_Turno
                            Else
                                Horas_Laboradas = Hora_Real
                            End If
                        End If
                    End If
                    Me.Refresh
                    'Consulta los movimientos del empleado para la fecha
                    Mi_SQL = "SELECT ISNULL(No_Movimiento,'') as No_Movimiento, "
                    Mi_SQL = Mi_SQL & " ISNULL(Tipo_Incidencia,'') AS Tipo_Incidencia,"
                    Mi_SQL = Mi_SQL & " Empleado_ID, Motivo, Simbologia, "
                    Mi_SQL = Mi_SQL & " SubSimbologia, Horas_Acuerdo"
                    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias "
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
                    Mi_SQL = Mi_SQL & " AND (" & Par_Fecha & Format(Fecha, "MM/dd/yyyy") & Par_Fecha
                    Mi_SQL = Mi_SQL & " BETWEEN Fecha_Inicio AND Fecha_Termino)"
                    Mi_SQL = Mi_SQL & " AND Estatus='A'"
                    Mi_SQL = Mi_SQL & " AND Tipo_Incidencia='E'"
                    Set Rs_Consulta_Adm_Permisos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Adm_Permisos.EOF Then
                        Permiso = Rs_Consulta_Adm_Permisos.rdoColumns("Motivo")
                        Referencia = Trim(Rs_Consulta_Adm_Permisos.rdoColumns("No_Movimiento"))
                        Simbologia = Rs_Consulta_Adm_Permisos.rdoColumns("Simbologia")
                        SubSimbologia = Rs_Consulta_Adm_Permisos.rdoColumns("SubSimbologia")
                        Tipo_Incidencia = Rs_Consulta_Adm_Permisos.rdoColumns("Tipo_Incidencia")
                        If Not IsNull(Rs_Consulta_Adm_Permisos.rdoColumns("Horas_Acuerdo")) Then
                            If Val(Rs_Consulta_Adm_Permisos.rdoColumns("Horas_Acuerdo")) > 0 Then
                                Horas_Laboradas = Val(Rs_Consulta_Adm_Permisos.rdoColumns("Horas_Acuerdo"))
                            End If
                        End If
                    End If
                    Rs_Consulta_Adm_Permisos.Close
                    Me.Refresh
                    'Valida que el dia festivo
                    Mi_SQL = "SELECT Fecha,Comentarios FROM Cat_Dias_No_Laborales"
                    Mi_SQL = Mi_SQL & " WHERE Fecha='" & Format(Fecha, "MM/dd/yyyy") & "'"
                    Set Rs_Consulta_Dia_Feriado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Dia_Feriado.EOF Then
                        Permiso = Rs_Consulta_Dia_Feriado.rdoColumns("Comentarios")
                        Simbologia = "FE"
                        SubSimbologia = ""
                        'Valida si el horario de salida es un día posterior para evitar horas negativas
                        If DateDiff("n", Hora_Inicio_Turno, Hora_Termino_Turno) > 0 Then
                            Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
                        Else
                            Horas_Turno = DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm:ss"), "23:59:59") / 60 + DateDiff("n", "00:00:00", Format(Hora_Termino_Turno, "HH:mm:ss")) / 60
                        End If
                        Horas_Laboradas = Horas_Turno
                    End If
                    Set Rs_Consulta_Dia_Feriado = Nothing
                    'Si fue su dia de descanso lo pone como trabajado
                    If Dia_Descanso = "SI" And Simbologia = "F" Then
'                        Permiso = "DESCANSO"
'                        Simbologia = "DE"
'                        SubSimbologia = ""
'                        Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
'                        Horas_Turno = Horas_Turno - (Abs(DateDiff("n", Format(Hora_Comida_Termino, "HH:mm:ss"), Format(Hora_Comida_Inicio, "HH:mm:ss"))) / 60)
'                        Horas_Laboradas = (DateDiff("n", Hora_Inicio_Turno, Hora_Termino_Turno) / 60)
'                        If Hora_Real - Fix(Hora_Real) < 0.5 Then
'                            Horas_Extra = Fix(Hora_Real)
'                        Else
'                            Horas_Extra = Fix(Hora_Real) + 0.5
'                        End If
                        Permiso = ""
                        Simbologia = ""
                        SubSimbologia = ""
                        Horas_Turno = Abs(DateDiff("n", Format(Hora_Termino_Turno, "HH:mm:ss"), Format(Hora_Inicio_Turno, "HH:mm:ss"))) / 60
                        Horas_Turno = Horas_Turno - (Abs(DateDiff("n", Format(Hora_Comida_Termino, "HH:mm:ss"), Format(Hora_Comida_Inicio, "HH:mm:ss"))) / 60)
                        Horas_Laboradas = (DateDiff("n", Hora_Inicio_Turno, Hora_Termino_Turno) / 60)
                        If Hora_Real - Fix(Hora_Real) < 0.5 Then
                            Horas_Extra = Fix(Hora_Real)
                        Else
                            Horas_Extra = Fix(Hora_Real) + 0.5
                        End If
                    End If
                    'Valida las horas extra trabajadas
                    Horas_Extra_Adicionales = 0
                    Horas_Extra = 0
                    If PG_Calcula_Horas_Extra = "1" Then
                        If Hora_Salida <> "" And Hora_Salida <> "01/01/1900" And Hora_Salida <> "0" Then
                            If Horas_Extra_Paga > 0 Then
                                'If Format(Hora_Real - Horas_Laboradas, "#0") >= Horas_Extra_Paga Then
                                If Fix((Hora_Real_Extra - Horas_Laboradas) + 0.25) >= Horas_Extra_Paga Then
                                    'Horas_Extra_Adicionales = Format(Hora_Real - Horas_Laboradas, "#0") - Horas_Extra_Paga
                                    Horas_Extra_Adicionales = Fix(Hora_Real_Extra - Horas_Laboradas + 0.25) - Horas_Extra_Paga
                                    'Horas_Extra = Val(Format(Horas_Extra_Adicionales, "#0")) + Val(Format(Horas_Extra_Definidas, "#0.00"))
                                    Horas_Extra = Val(Fix(Horas_Extra_Adicionales)) + Val(Format(Horas_Extra_Definidas, "#0.00"))
                                Else
                                    Horas_Extra_Adicionales = Fix(Hora_Real_Extra - Horas_Laboradas + 0.25)
                                    Horas_Extra = Horas_Extra_Adicionales
                                End If
                            Else
                                'Horas_Extra_Adicionales = Hora_Real - Horas_Laboradas
                                Horas_Extra_Adicionales = Fix(Hora_Real_Extra - Horas_Laboradas + 0.25)
                                Horas_Extra = Val(Format(Horas_Extra_Adicionales, "#0"))
                            End If
                            If Horas_Extra < 0 Then Horas_Extra = 0
                        End If
                    End If
                    Me.Refresh
                    Partida = Partida + 1
                    'Si no tiene checada de algún turno (entrada/salida) le asigan las horas trabajadas del turno sin horas extra
                    If Hora_Real < 0 Then Hora_Real = 0
                    If Hora_Entrada <> "01/01/1900" And Hora_Salida = "01/01/1900" Then
                        Hora_Real = Horas_Laboradas
                    End If
                    
                    
                    
                    'Agrega el dato en el grid
                    Grid_Validacion_Horas_Trabajo_Lista.AddItem Referencia _
                        & Chr(9) & .rdoColumns("Empleado_ID") _
                        & Chr(9) & .rdoColumns("No_Tarjeta") _
                        & Chr(9) & .rdoColumns("Departamento") _
                        & Chr(9) & .rdoColumns("Nombre") _
                        & Chr(9) & Nombre_Turno _
                        & Chr(9) & Format(Hora_Entrada, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida_Entrada, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Comida_Salida, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Salida, "HH:mm:ss") _
                        & Chr(9) & Format(Hora_Real, "#0.00") _
                        & Chr(9) & Format(Horas_Laboradas, "#0.00") _
                        & Chr(9) & Horas_Extra _
                        & Chr(9) & Justificacion_Horas_Extra _
                        & Chr(9) & Horas_Extra _
                        & Chr(9) & Permiso _
                        & Chr(9) & Simbologia _
                        & Chr(9) & SubSimbologia _
                        & Chr(9) & No_Movimiento _
                        & Chr(9) & "NO" _
                        & Chr(9) & Referencia _
                        & Chr(9) & Tipo_Incidencia _
                        & Chr(9) & Horas_Laboradas_Turno _
                        & Chr(9) & Turno_Empleado _
                        & Chr(9) & "" & Chr(9) & Calendario_Turno_ID & Chr(9) & Calendario_Turno_Detalle_ID
                    Print #1, ""
                    Print #2, .rdoColumns("No_Tarjeta"); _
                        "|"; .rdoColumns("Departamento"); _
                        "|"; .rdoColumns("Nombre"); _
                        "|"; Nombre_Supervisor; _
                        "|"; Format(Hora_Entrada, "HH:mm:ss"); _
                        "|"; Format(Hora_Comida_Entrada, "HH:mm:ss"); _
                        "|"; Format(Hora_Comida_Salida, "HH:mm:ss"); _
                        "|"; Format(Hora_Salida, "HH:mm:ss"); _
                        "|"; Format(Hora_Real, "#0.00"); _
                        "|"; Format(Horas_Laboradas, "#0.00"); _
                        "|"; Horas_Extra; _
                        "|"; Justificacion_Horas_Extra; _
                        "|"; Horas_Extra; _
                        "|"; Turno_Empleado; _
                        "|"; Permiso; _
                        "|"; Simbologia; _
                        "|"; No_Movimiento; _
                        "|"; .rdoColumns("Departamento"); _
                        "|"; .rdoColumns("Clave"); _
                        "|"; .rdoColumns("Tipo_Empleado"); _
                        "|"; .rdoColumns("Cedula_Identidad_Ciudadana"); _
                        "|"; Ruta_Transporte
                    Me.Refresh
                    'Realiza la validacion de inconsistencias
                    With Grid_Validacion_Horas_Trabajo_Lista
                        Select Case .TextMatrix(.Rows - 1, 12)
                            Case "HI":
                                Colorear_Fila = True
                            Case "VN", "FE":
                                If Val(.TextMatrix(.Rows - 1, 7)) <> Horas_Turno Then
                                    Colorear_Fila = True
                                End If
                            Case "PE":
                                If Val(.TextMatrix(.Rows - 1, 7)) = 0 Then
                                    Colorear_Fila = True
                                End If
                            Case "II":
                                Select Case .TextMatrix(.Rows - 1, 11)
                                    Case "EG"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> 0 Then
                                            Colorear_Fila = True
                                        End If
                                    Case "MA"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> 0 Then
                                            Colorear_Fila = True
                                        End If
                                    Case "RT"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> 0 Then
                                            Colorear_Fila = True
                                        End If
                                End Select
                            Case "ID"
                                Select Case .TextMatrix(.Rows - 1, 11)
                                    Case "VA"
                                        If Val(.TextMatrix(.Rows - 1, 6)) <> 0 Then
                                            Colorear_Fila = True
                                        End If
                                    Case "AL"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> Horas_Turno Then
                                            Colorear_Fila = True
                                        End If
                                    Case "DE"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> Horas_Turno Then
                                            Colorear_Fila = True
                                        End If
                                    Case "MO"
                                        If Val(.TextMatrix(.Rows - 1, 7)) <> Horas_Turno Then
                                            Colorear_Fila = True
                                        End If
                                End Select
                            Case "D", "F"   'Descanso y falta injustificada
                                If Val(.TextMatrix(.Rows - 1, 7)) <> 0 Then
                                    Colorear_Fila = True
                                End If
                        End Select
                        Manejo_Grid = False
                        If Colorear_Fila = True Then
                            .Col = 0
                            For Columna = 3 To .Cols - 1
                                .Col = Columna
                                .Row = .Rows - 1
                                .CellBackColor = &H80FFFF
                            Next Columna
                        End If
                        
                        'Colorea tiempos de trabajo excedidos de 14 hrs.
                        If PG_Horas_Maximas_Turno > 0 Then
                            If Val(.TextMatrix(.Rows - 1, 10)) > PG_Horas_Maximas_Turno Then
                                .Col = 0
                                For Columna = 3 To .Cols - 1
                                    .Col = Columna
                                    .Row = .Rows - 1
                                    .CellBackColor = &H80FF&
                                Next Columna
                            End If
                        End If
                        'Colorea los registros sin hora de salida
                        If Trim(.TextMatrix(.Rows - 1, 9)) = "00:00:00" Then
                            .Col = 0
                            For Columna = 3 To .Cols - 1
                                .Col = Columna
                                .Row = .Rows - 1
                                .CellBackColor = &HFFFF00
                            Next Columna
                        End If
                        
                        Me.Refresh
                        'Verfica si el empleado es temporal para revisar la finalizacion de contrato
                        Mi_SQL = "SELECT Empleado_ID, Tipo_Contratacion, Fecha_Termino_Contrato"
                        Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
                        Mi_SQL = Mi_SQL & " WHERE Empleado_ID = '" & .TextMatrix(.Rows - 1, 1) & "'"
                        Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                        If Not Rs_Consulta_Cat_Empleados.EOF Then
                            If UCase(Rs_Consulta_Cat_Empleados.rdoColumns("Tipo_Contratacion")) = "EVENTUAL" Then
                                If DateDiff("d", Now, Rs_Consulta_Cat_Empleados.rdoColumns("Fecha_Termino_Contrato")) <= Dias_Aviso_Contrato_Eventual And _
                                    DateDiff("d", Now, Rs_Consulta_Cat_Empleados.rdoColumns("Fecha_Termino_Contrato")) > 0 Then
                                    'If Colorear_Fila = True Then
                                        .Col = 0
                                        For Columna = 1 To .Cols - 1
                                            .Col = Columna
                                            .Row = .Rows - 1
                                            .CellBackColor = &H8080FF
                                        Next Columna
                                    'End If
                                End If
                                If DateDiff("d", Now, Rs_Consulta_Cat_Empleados.rdoColumns("Fecha_Termino_Contrato")) <= 0 Then
                                    'If Colorear_Fila = True Then
                                        .Col = 0
                                        For Columna = 1 To .Cols - 1
                                            .Col = Columna
                                            .Row = .Rows - 1
                                            .CellBackColor = &HC0&
                                        Next Columna
                                    'End If
                                End If
                            End If
                        End If
                        Set Rs_Consulta_Cat_Empleados = Nothing
                        Me.Refresh
                    End With
                Else
                    'Consulta la asistencia ya validada para mostrarla en pantalla sin posibilidad de editarla
                    Mi_SQL = "SELECT * FROM Adm_Asistencias"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
                    Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Fecha, "MM/dd/yyyy") & "'"
                    Set Rs_Consulta_Adm_Asistencias_Validadas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Adm_Asistencias_Validadas.EOF Then
                        Turno_ID = Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Turno_ID")
                        If Not IsNull(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Calendario_Turno_ID")) Then
                            Calendario_Turno_ID = Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Calendario_Turno_ID")
                        End If
                        If Not IsNull(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Calendario_Turno_Detalle_ID")) Then
                            Calendario_Turno_Detalle_ID = Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Calendario_Turno_Detalle_ID")
                        End If
                        'Consulta el nombre del supervisor
                        If Not IsNull(.rdoColumns("Supervisor_ID")) Then
                            Mi_SQL = "SELECT (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Supervisor FROM Cat_Empleados"
                            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Supervisor_ID") & "'"
                            Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Cat_Empleados.EOF Then
                                Nombre_Supervisor = Rs_Consulta_Cat_Empleados.rdoColumns("Supervisor")
                            Else
                                Nombre_Supervisor = ""
                            End If
                            Rs_Consulta_Cat_Empleados.Close
                        Else
                            Nombre_Supervisor = ""
                        End If
                        If Trim(Turno_ID) <> "" Then
                            Mi_SQL = "SELECT Calendario_Turno_ID,Calendario_Turno_Detalle_ID,Nombre_Turno,Hora_Inicio,Hora_Termino FROM Cat_Calendarios_Turnos_Detalles"
                            Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID='" & Calendario_Turno_ID & "'"
                            Mi_SQL = Mi_SQL & " AND Calendario_Turno_Detalle_ID='" & Calendario_Turno_Detalle_ID & "'"
                            Mi_SQL = Mi_SQL & " AND Estatus<>'ELIMINADO'"
                            Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Cat_Empleados.EOF Then
                                Nombre_Turno = Rs_Consulta_Cat_Empleados.rdoColumns("Nombre_Turno")
                                Horas_Turno = DateDiff("n", Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Inicio"), "HH:mm:ss"), Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Termino"), "HH:mm:ss")) / 60
                                If Horas_Turno < 0 Then
                                    Horas_Turno = 24 + Horas_Turno
                                End If
                            Else
                                Mi_SQL = "SELECT Turno_ID,Nombre,Hora_Inicio,Hora_Termino FROM Cat_Turnos"
                                Mi_SQL = Mi_SQL & " WHERE Turno_ID='" & Turno_ID & "'"
                                Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                If Not Rs_Consulta_Cat_Empleados.EOF Then
                                    Nombre_Turno = Rs_Consulta_Cat_Empleados.rdoColumns("Nombre")
                                    Horas_Turno = DateDiff("n", Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Inicio"), "HH:mm:ss"), Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Termino"), "HH:mm:ss")) / 60
                                    If Horas_Turno < 0 Then
                                        Horas_Turno = 24 + Horas_Turno
                                    End If
                                Else
                                    Nombre_Turno = ""
                                    Horas_Turno = 0
                                End If
                            End If
                            Rs_Consulta_Cat_Empleados.Close
                        Else
                            Mi_SQL = "SELECT Calendario_Turno_ID,Calendario_Turno_Detalle_ID,Nombre_Turno,Hora_Inicio,Hora_Termino FROM Cat_Calendarios_Turnos_Detalles"
                            Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID='" & Calendario_Turno_ID & "'"
                            Mi_SQL = Mi_SQL & " AND Calendario_Turno_Detalle_ID='" & Calendario_Turno_Detalle_ID & "'"
                            Mi_SQL = Mi_SQL & " AND Estatus<>'ELIMINADO'"
                            Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Cat_Empleados.EOF Then
                                Nombre_Turno = Rs_Consulta_Cat_Empleados.rdoColumns("Nombre_Turno")
                                Horas_Turno = DateDiff("n", Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Inicio"), "HH:mm:ss"), Format(Rs_Consulta_Cat_Empleados.rdoColumns("Hora_Termino"), "HH:mm:ss")) / 60
                                If Horas_Turno < 0 Then
                                    Horas_Turno = 24 + Horas_Turno
                                End If
                            Else
                                Nombre_Turno = ""
                                Horas_Turno = 0
                            End If
                            Rs_Consulta_Cat_Empleados.Close
                        End If
                        'Consulta los movimientos del empleado para la fecha
                        If Not IsNull(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Referencia")) And Trim(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Tipo_Incidencia")) <> "" Then
                            Mi_SQL = "SELECT No_Movimiento,Tipo_Incidencia,Empleado_ID,Motivo,Simbologia,SubSimbologia,Horas_Acuerdo"
                            Mi_SQL = Mi_SQL & " FROM Adm_Movimientos_Asistencias"
                            Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Referencia") & "'"
                            Mi_SQL = Mi_SQL & " AND Estatus='A'"
                            Mi_SQL = Mi_SQL & " AND Tipo_Incidencia='E'"
                            Set Rs_Consulta_Adm_Permisos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Consulta_Adm_Permisos.EOF Then
                                Permiso_Validado = Rs_Consulta_Adm_Permisos.rdoColumns("Motivo")
                                Permiso_Referencia = Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Referencia")
                            Else
                                Permiso_Validado = ""
                                Permiso_Referencia = ""
                            End If
                            Rs_Consulta_Adm_Permisos.Close
                        Else
                            Permiso_Validado = ""
                            Permiso_Referencia = ""
                        End If
                        'Agrega el dato en el grid
                        Dim JHE As String
                        If Not IsNull(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Justificacion_Horas_Extra")) Then
                            JHE = Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Justificacion_Horas_Extra")
                        Else
                            JHE = ""
                        End If
                        Grid_Validacion_Horas_Trabajo_Lista.AddItem "" _
                            & Chr(9) & .rdoColumns("Empleado_ID") _
                            & Chr(9) & .rdoColumns("No_Tarjeta") _
                            & Chr(9) & .rdoColumns("Departamento") _
                            & Chr(9) & .rdoColumns("Nombre") _
                            & Chr(9) & Nombre_Turno _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Entrada_Turno"), "HH:mm:ss") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Entrada_Comida"), "HH:mm:ss") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Salida_Comida"), "HH:mm:ss") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Salida_Turno"), "HH:mm:ss") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Aprobadas"), "#0.00") _
                            & Chr(9) & Format(Horas_Turno, "#0.00") _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Extra") _
                            & Chr(9) & JHE _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Calculadas") _
                            & Chr(9) & Permiso_Validado _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Simbologia") _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Subsimbologia") _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Referencia") _
                            & Chr(9) & "V" _
                            & Chr(9) & Permiso_Referencia _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Tipo_Incidencia") _
                            & Chr(9) & "" _
                            & Chr(9) & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Turno_ID") _
                            & Chr(9) & Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Fecha_Creo"), "dd/MMM/yyyy HH:mm") & " - " & Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Usuario_Creo") & Chr(9) & Calendario_Turno_ID & Chr(9) & Calendario_Turno_Detalle_ID
                        Print #1, ""
                        Print #2, .rdoColumns("No_Tarjeta"); _
                            "|"; .rdoColumns("Departamento"); _
                            "|"; .rdoColumns("Nombre"); _
                            "|"; Nombre_Supervisor; _
                            "|"; Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Entrada_Turno"), "HH:mm:ss"); _
                            "|"; Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Entrada_Comida"), "HH:mm:ss"); _
                            "|"; Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Salida_Comida"), "HH:mm:ss"); _
                            "|"; Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Hora_Salida_tURNO"), "HH:mm:ss"); _
                            "|"; Format(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Aprobadas"), "#0.00"); _
                            "|"; Format(0, "#0.00"); _
                            "|"; Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Extra"); _
                            "|"; JHE; _
                            "|"; Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Horas_Calculadas"); _
                            "|"; Val(Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Turno_ID")); _
                            "|"; Permiso_Validado; _
                            "|"; Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Simbologia"); _
                            "|"; Rs_Consulta_Adm_Asistencias_Validadas.rdoColumns("Referencia"); _
                            "|"; .rdoColumns("Departamento"); _
                            "|"; .rdoColumns("Clave"); _
                            "|"; .rdoColumns("Tipo_Empleado"); _
                            "|"; .rdoColumns("Cedula_Identidad_Ciudadana"); _
                            "|"; Ruta_Transporte
                        'Colorea los registros de horas extra diferentes entre el calculado y manual
                        If Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.Rows - 1, 12)) <> Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Grid_Validacion_Horas_Trabajo_Lista.Rows - 1, 14)) Then
                            Grid_Validacion_Horas_Trabajo_Lista.Col = 0
                            For Columna = 12 To 14
                                Grid_Validacion_Horas_Trabajo_Lista.Col = Columna
                                Grid_Validacion_Horas_Trabajo_Lista.Row = Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
                                Grid_Validacion_Horas_Trabajo_Lista.CellBackColor = vbRed
                            Next
                        End If
                        
                        Me.Refresh
                    End If
                    Rs_Consulta_Adm_Asistencias_Validadas.Close
                End If
                PrgBar_Validacion_Horas.Value = PrgBar_Validacion_Horas.Value + 1
                .MoveNext
            Wend
            Call Finalizar_Reporte
        End With
    End If
    Me.Refresh
    
    Manejo_Grid = True
    'Configuracion del grid
    With Grid_Validacion_Horas_Trabajo_Lista
        If .Rows > 1 Then .FixedRows = 1
            .FixedCols = 3
            .ColWidth(0) = 0        'Referencia
            .ColWidth(1) = 0        'Empleado_ID
            .ColWidth(2) = 600      'No_Tarjeta
            .ColAlignment(2) = flexAlignCenterCenter
            .ColWidth(3) = 1850     'Departamento
            .ColWidth(4) = 3000     'Nombre Empleado
            .ColWidth(5) = 600      'Turno
            .ColAlignment(5) = flexAlignLeftCenter
            .ColWidth(6) = 800      'Entrada
            .ColAlignment(6) = flexAlignCenterCenter
            .ColWidth(7) = 0        'Comida Salida
            .ColWidth(8) = 0        'Comida Entrada
            .ColWidth(9) = 800      'Salida
            .ColAlignment(9) = flexAlignCenterCenter
            .ColWidth(10) = 700     'Horas
            .ColWidth(11) = 700     'Horas Acuerdo
            .ColWidth(12) = 700     'Horas Extra
            .ColWidth(13) = 2650     'Justificacion Horas Extra
            .ColWidth(14) = 700     'Calculadas
            .ColWidth(15) = 2650    'Observaciones
            .ColWidth(16) = 500     'Simbologia
            .ColWidth(17) = 0       'SubSimbologia
            .ColWidth(18) = 0       'No_Detalle
            .ColWidth(19) = 600     'Validar
            .ColWidth(20) = 1200    'Referencia
            .ColWidth(21) = 600     'Movimiento
            .ColWidth(22) = 0       'Horas Turno
            .ColWidth(23) = 0       'TurnoID
            .ColWidth(24) = 3000    'Fecha Validó
            .ColAlignment(24) = flexAlignLeftCenter
            .ColWidth(25) = 0       'CalendarioTurnoID
            .ColWidth(26) = 0       'CalendarioTurnoDetalleID
    End With
    Me.Refresh
    If Grid_Validacion_Horas_Trabajo_Lista.Rows > 1 Then
        Pic_Adm_Validacion_Horas_Trabajo_Lista.Visible = True
        Pic_Adm_Validacion_Horas_Trabajo.Visible = False
        Me.Height = 8500
        Me.Width = 14900
        Me.Top = 0
        Me.Left = 0
        Fra_Validacion_Horas_Trabajo_Lista.Enabled = True
        Lbl_Validacion_Horas_Supervisor.Caption = Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text)
        Lbl_Validacion_Horas_Fecha_Lista.Caption = "Fecha: " & Format(Fecha, "dd/MMM/yyyy")
        Btn_Validar_Horas_Empleados.Enabled = False
        'Bloquea el botón de validar los días después del hoy
        If DateDiff("d", Now, Fecha) < 1 Then
            Btn_Validar_Horas_Empleados.Enabled = True
        End If
        Chk_Seleccionar_Todas.Value = 0
    Else
        Lbl_Validacion_Horas_Supervisor.Caption = Trim(Cmb_Adm_Validacion_Horas_Supervisor.Text)
        Lbl_Validacion_Horas_Fecha_Lista.Caption = "Fecha: " & Format(Fecha, "dd/MMM/yyyy")
        MsgBox "No existe informacion con los parametros seleccionados ó" & vbCrLf & _
                "ya se valido la información del dia seleccionado en su totalidad." & vbCrLf & _
                "Si desea conocer la información validada, puede intentar con la opción Imprimir", vbInformation + vbOKOnly, Me.Caption
    End If
    Me.Refresh
    PrgBar_Validacion_Horas.Visible = False
    Me.MousePointer = 0
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Incosistencia.Visible = True
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Incosistencia.ZOrder vbBringToFront
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencer.Visible = True
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencer.ZOrder vbBringToFront
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencido.Visible = True
    Pic_Adm_Validacion_Horas_Trabajo_Lista_Contrato_x_Vencido.ZOrder vbBringToFront
Exit Sub
HANDLER:
    Close #1
    Close #2
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
'DESCRIPCION: Actualiza los datos de los empleados
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 18-Mayo-2011
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Guardar_Lista()
Dim Rs_Adm_Asistencias As rdoResultset     'Informacion del detalles
Dim Rs_Consulta_Informacion_Turnos As rdoResultset
Dim Rs_Modifica_Cat_Empleados As rdoResultset
Dim Rs_Alta_Adm_Asistencias_Detalles As rdoResultset
Dim Cont_Fila As Integer                                'Se utiliza para recorrer el grid de la lista
Dim Supervidor_ID_Empleado As String
Dim Turno_Empleado As String                            'Guarda el Turno del empleado
Dim Calendario_Turno_ID As String
Dim Calendario_Turno_Detalle_ID As String
Dim Hora_Inicio_Turno As Date                           'Guarda la hora de inicio del turno
Dim Hora_Termino_Turno As Date                          'Guarda la hora de inicio del turno
Dim Hora_Comida_Salida As Date
Dim Hora_Comida_Entrada As Date
Dim Hora_Entrada As Date
Dim Hora_Salida As Date
Dim Horas As Integer
Dim Horas_Acuerdo As Integer
Dim Horas_Turno As Double
Dim cadena_mensaje As String

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Set Rs_Adm_Asistencias = Conectar_Ayudante.Recordset_Agregar("Adm_Asistencias")
    Pbar_Validacion.Visible = True
    Pbar_Validacion.Max = Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
    Pbar_Validacion.Value = 0
    For Cont_Fila = 1 To Grid_Validacion_Horas_Trabajo_Lista.Rows - 1
        MDIFrm_Apl_Principal.MousePointer = 11
        cadena_mensaje = ""
        If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 19) = "SI" Then
            Supervidor_ID_Empleado = ""
            'Consulta los datos del turno
            Mi_SQL = "SELECT * FROM Cat_Turnos"
            Mi_SQL = Mi_SQL & " WHERE Turno_ID='" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 23)) & "'"
            Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Informacion_Turnos.EOF Then
                Turno_Empleado = Rs_Consulta_Informacion_Turnos.rdoColumns("Turno_ID")
                Hora_Inicio_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                Hora_Termino_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Termino"), "HH:mm:ss")
                Hora_Comida_Salida = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                Hora_Comida_Entrada = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Termino"), "HH:mm:ss")
            Else
                Mi_SQL = "SELECT * FROM Cat_Calendarios_Turnos_Detalles"
                Mi_SQL = Mi_SQL & " WHERE Calendario_Turno_ID='" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 25)) & "'"
                Mi_SQL = Mi_SQL & " AND Calendario_Turno_Detalle_ID='" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 26)) & "'"
                Mi_SQL = Mi_SQL & " AND Estatus<>'ELIMINADO'"
                Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Consulta_Informacion_Turnos.EOF Then
                    Calendario_Turno_ID = Rs_Consulta_Informacion_Turnos.rdoColumns("Calendario_Turno_ID")
                    Calendario_Turno_Detalle_ID = Rs_Consulta_Informacion_Turnos.rdoColumns("Calendario_Turno_Detalle_ID")
                    Hora_Inicio_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Inicio"), "HH:mm:ss")
                    Hora_Termino_Turno = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Hora_Termino"), "HH:mm:ss")
                    Hora_Comida_Salida = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Inicio"), "HH:mm:ss")
                    Hora_Comida_Entrada = Format(Rs_Consulta_Informacion_Turnos.rdoColumns("Comida_Termino"), "HH:mm:ss")
                End If
            End If
            Rs_Consulta_Informacion_Turnos.Close
                Hora_Inicio_Turno = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6), "HH:mm:ss")
                Hora_Termino_Turno = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9), "HH:mm:ss")
                Hora_Comida_Salida = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8), "HH:mm:ss")
                Hora_Comida_Entrada = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7), "HH:mm:ss")
            Set Rs_Consulta_Informacion_Turnos = Nothing
'            With Rs_Adm_Asistencias
'                .AddNew
'                    'Obtiene el maximo de la tabla
'                    .rdoColumns("No_Asistencia") = Conectar_Ayudante.Maximo_Catalogo("Adm_Asistencias", "No_Asistencia")
'                    .rdoColumns("Empleado_ID") = Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 1)
'                    .rdoColumns("Turno_ID") = Turno_Empleado
'                    If Supervidor_ID_Empleado <> "" Then
'                        .rdoColumns("Supervisor_ID") = Supervidor_ID_Empleado
'                    End If
'                    .rdoColumns("Referencia") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 0))
'                    .rdoColumns("Tipo_Incidencia") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 20))
'                    .rdoColumns("No_Tarjeta") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 2))
'                    .rdoColumns("Fecha") = Format(Fecha, "MM/dd/yyyy")
'                    .rdoColumns("Hora_Entrada_Turno") = Format(Hora_Inicio_Turno, "HH:mm:ss")
'                    .rdoColumns("Hora_Salida_Turno") = Format(Hora_Termino_Turno, "HH:mm:ss")
'                    .rdoColumns("Hora_Salida_Comida_Turno") = Format(Hora_Comida_Salida, "HH:mm:ss")
'                    .rdoColumns("Hora_Entrada_Comida_Turno") = Format(Hora_Comida_Entrada, "HH:mm:ss")
'                    .rdoColumns("Tiempo_Retardo") = 0
'                    .rdoColumns("Simbologia") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 15))
'                    .rdoColumns("SubSimbologia") = Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 16))
'                    'Valida horas de entrada
'                    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6) <> "" Then
'                        .rdoColumns("Hora_Entrada") = Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss")
'                        Hora_Entrada = Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss")
'                        'Verifica si no tiene permisos o incidencia extraordinaria para acumular el retardo
'                        If Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 19)) = "" And Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 20)) = "" Then
'                            If DateDiff("n", DateAdd("n", Minutos_Tolerancia, Format(Hora_Inicio_Turno, "HH:mm:ss")), Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss")) > 0 Then
'                                .rdoColumns("Tiempo_Retardo") = DateDiff("n", DateAdd("n", Minutos_Tolerancia, Format(Hora_Inicio_Turno, "HH:mm:ss")), Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss"))
'                            End If
'                        End If
'                    Else
'                        .rdoColumns("Hora_Entrada") = Format(0, "HH:mm:ss")
'                        Hora_Entrada = Format(0, "HH:mm")
'                    End If
'                    'Hora salida comida
'                    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7) <> "" Then
'                        .rdoColumns("Hora_Salida_Comida") = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7), "HH:mm:ss")
'                    Else
'                        .rdoColumns("Hora_Salida_Comida") = Format(0, "HH:mm:ss")
'                    End If
'                    'Hora salida comida
'                    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8) <> "" Then
'                        .rdoColumns("Hora_Entrada_Comida") = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8), "HH:mm:ss")
'                    Else
'                        .rdoColumns("Hora_Entrada_Comida") = Format(0, "HH:mm:ss")
'                    End If
'                    'Valida hora de salida
'                    If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9) <> "" Then
'                        .rdoColumns("Hora_Salida") = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9), "HH:mm:ss")
'                        Hora_Salida = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9), "HH:mm")
'                    Else
'                        .rdoColumns("Hora_Salida") = Format(0, "HH:mm:ss")
'                        Hora_Salida = Format(0, "HH:mm")
'                    End If
'                    Horas_Turno = Val(DateDiff("n", Format(Hora_Inicio_Turno, "HH:mm"), Format(Hora_Termino_Turno, "HH:mm"))) - Val(DateDiff("n", Format(Hora_Comida_Salida, "HH:mm"), Format(Hora_Comida_Entrada, "HH:mm")))
'                    .rdoColumns("Horas_Extra") = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12))
'                    .rdoColumns("Horas_Calculadas") = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 13))
'                    If Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12)) > 0 Then
'                        .rdoColumns("Simbologia") = "A"
'                    End If
'                    .rdoColumns("Horas_Aprobadas") = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 10))
'                    Horas_Acuerdo = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 11))
'                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
'                    .rdoColumns("Fecha_Creo") = Now
'                .Update
'            End With
            Mi_SQL = "INSERT INTO Adm_Asistencias(Empleado_ID,Turno_ID,No_Tarjeta,Fecha,Hora_Entrada_Turno,Hora_Salida_Turno,Hora_Entrada_Comida_Turno,Hora_Salida_Comida_Turno"
            Mi_SQL = Mi_SQL & " ,Hora_Entrada_Comida,Hora_Salida_Comida,Hora_Entrada,Hora_Salida,Horas_Extra,Justificacion_Horas_Extra,Horas_Aprobadas,Tiempo_Retardo,Simbologia,SubSimbologia"
            Mi_SQL = Mi_SQL & " ,Referencia,Tipo_Incidencia,Horas_Calculadas,Usuario_Creo,Fecha_Creo,Calendario_Turno_ID,Calendario_Turno_Detalle_ID)"
            Mi_SQL = Mi_SQL & " VALUES('" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 1)) & "'"
            Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 23)) & "'"
            Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 2)) & "'"
            Mi_SQL = Mi_SQL & " ,'" & Format(Fecha, "MM/dd/yyyy") & "'"
            Mi_SQL = Mi_SQL & " ,'" & Format(Hora_Inicio_Turno, "HH:mm:ss") & "'"
            Mi_SQL = Mi_SQL & " ,'" & Format(Hora_Termino_Turno, "HH:mm:ss") & "'"
            Mi_SQL = Mi_SQL & " ,'" & Format(Hora_Comida_Entrada, "HH:mm:ss") & "'"
            Mi_SQL = Mi_SQL & " ,'" & Format(Hora_Comida_Salida, "HH:mm:ss") & "'"
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 8), "HH:mm:ss") & "'"
            Else
                Mi_SQL = Mi_SQL & " ,'" & Format(0, "HH:mm:ss") & "'"
            End If
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 7), "HH:mm:ss") & "'"
            Else
                Mi_SQL = Mi_SQL & " ,'" & Format(0, "HH:mm:ss") & "'"
            End If
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6), "HH:mm:ss") & "'"
                Hora_Entrada = Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss")
            Else
                Mi_SQL = Mi_SQL & " ,'" & Format(0, "HH:mm:ss") & "'"
                Hora_Entrada = Format(0, "HH:mm")
            End If
            If Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9), "HH:mm:ss") & "'"
                Hora_Salida = Format(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 9), "HH:mm")
            Else
                Mi_SQL = Mi_SQL & " ,'" & Format(0, "HH:mm:ss") & "'"
                Hora_Salida = Format(0, "HH:mm")
            End If
            Mi_SQL = Mi_SQL & " ," & Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12))
            Mi_SQL = Mi_SQL & " ,'" & Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 13) & "'"
            Mi_SQL = Mi_SQL & " ," & Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 10))
            'Verifica si no tiene permisos o incidencia extraordinaria para acumular el retardo
            If Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)) <> "" And Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 20)) = "" And Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 21)) = "" Then
                If DateDiff("n", DateAdd("n", Minutos_Tolerancia, Format(Hora_Inicio_Turno, "HH:mm:ss")), Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss")) > 0 Then
                    Mi_SQL = Mi_SQL & " ," & DateDiff("n", DateAdd("n", Minutos_Tolerancia, Format(Hora_Inicio_Turno, "HH:mm:ss")), Format((Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 6)), "HH:mm:ss"))
                Else
                    Mi_SQL = Mi_SQL & " ,0"
                End If
            Else
                Mi_SQL = Mi_SQL & " ,0"
            End If
            If Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12)) > 0 Then
                Mi_SQL = Mi_SQL & " ,'A'"
            Else
                Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 16)) & "'"
            End If
            Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 17)) & "'"
            Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 0)) & "'"
            Mi_SQL = Mi_SQL & " ,'" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 21)) & "'"
            Mi_SQL = Mi_SQL & " ," & Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 14))
            Mi_SQL = Mi_SQL & " ,'" & Nombre_Usuario & "'"
            Mi_SQL = Mi_SQL & " ,GETDATE()"
            If Trim(Calendario_Turno_ID) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Calendario_Turno_ID & "'"
            Else
                Mi_SQL = Mi_SQL & " ,NULL"
            End If
            If Trim(Calendario_Turno_Detalle_ID) <> "" Then
                Mi_SQL = Mi_SQL & " ,'" & Calendario_Turno_Detalle_ID & "')"
            Else
                Mi_SQL = Mi_SQL & " ,NULL)"
            End If
            Conexion_Base.Execute Mi_SQL
            'Actualiza la informacion del detalles de importacion
            Mi_SQL = "SELECT * FROM Adm_Asistencias_Detalles"
            Mi_SQL = Mi_SQL & " WHERE No_Operacion=" & Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 18)
            Set Rs_Consulta_Informacion_Turnos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Consulta_Informacion_Turnos.EOF Then
                Rs_Consulta_Informacion_Turnos.Edit
                    Rs_Consulta_Informacion_Turnos.rdoColumns("Validada") = "S"
                    Rs_Consulta_Informacion_Turnos.rdoColumns("Fecha_Valido") = Now
                    Rs_Consulta_Informacion_Turnos.rdoColumns("Horas_Extra") = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12))
                    Rs_Consulta_Informacion_Turnos.rdoColumns("Justificacion_Horas_Extra") = Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 13))
                Rs_Consulta_Informacion_Turnos.Update
            Else
                'Inserta el registr de detalles
                Mi_SQL = "INSERT INTO Adm_Asistencias_Detalles(Empleado_ID,No_Tarjeta,Fecha,Hora_Entrada,Hora_Salida,Hora_Comida_Entrada,Hora_Comida_Salida,Horas_Laboradas,Validada"
                Mi_SQL = Mi_SQL & " ,Fecha_Importacion,Proceso,Fecha_Valido,Horas_Extra,Justificacion_Horas_Extra)"
                Mi_SQL = Mi_SQL & " VALUES('" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 1)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Trim(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 2)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Fecha, "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " , '00:00:00'"
                Mi_SQL = Mi_SQL & " , '00:00:00'"
                Mi_SQL = Mi_SQL & " , '00:00:00'"
                Mi_SQL = Mi_SQL & " , '00:00:00'"
                Mi_SQL = Mi_SQL & " , 0"
                Mi_SQL = Mi_SQL & " , 'S'"
                Mi_SQL = Mi_SQL & " , GETDATE()"
                Mi_SQL = Mi_SQL & " , 'MANUAL-VALIDACION'"
                Mi_SQL = Mi_SQL & " , GETDATE()"
                Mi_SQL = Mi_SQL & " , " & Val(Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 12))
                Mi_SQL = Mi_SQL & " , '" & Grid_Validacion_Horas_Trabajo_Lista.TextMatrix(Cont_Fila, 13) & "')"
                Conexion_Base.Execute Mi_SQL
            End If
            Rs_Consulta_Informacion_Turnos.Close
        End If
        Pbar_Validacion.Value = Pbar_Validacion.Value + 1
        Me.Refresh
    Next
'    Rs_Adm_Asistencias.Close
    Conexion_Base.CommitTrans
    Fra_Validacion_Horas_Trabajo_Lista.Enabled = False
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox "Información Guardada Correctamente", vbInformation + vbOKOnly, Me.Caption
    Pbar_Validacion.Visible = False
Exit Sub
HANDLER:    'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
    MDIFrm_Apl_Principal.MousePointer = 0
    Pbar_Validacion.Visible = False
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Imprimir()
Dim linea As String 'Obtiene el texto a imprimir
Dim x As Printer
Dim contar_linea As Integer
Dim Foto_Empleado As New StdPicture
Dim No_Tarjeta As String
Dim Cont_Fila As Integer
Dim Cordenada_Y_Imagen  As Double
Dim Cont_Saltos As Integer
Dim Mi_SQL As String
Dim Rs_Conssultar_Foto As rdoResultset
Dim Contar_Filas As Integer
Dim Numero_Pagina As Integer

On Error GoTo HANDLER
    MDIFrm_Apl_Principal.MousePointer = 11
    Cordenada_Y_Imagen = 0
    
    Printer.FontSize = 8
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 11
    Printer.Font = "COURIER NEW"
    Printer.Print
    Printer.FontSize = 8
    Printer.Font = "Courier New"
    
    Open Ruta_Temporal & "Reporte_Validacion_Tiempo_Trabajo.txt" For Input As #1
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
    Printer.EndDoc
    Close #1
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Generar_Reporte()
Dim Rs_Consulta_Adm_Asistencias As rdoResultset
Dim Movimiento As String
Dim linea As String 'Obtiene el texto a imprimir
Dim Horas_Trabajadas As Double
Dim contar_linea As Integer

    Mi_SQL = "SELECT AA.No_Tarjeta, AA.Fecha, AA.Hora_Entrada, AA.Hora_Salida, AA.Hora_Entrada_Comida, AA.Hora_Salida_Comida,"
    Mi_SQL = Mi_SQL & " AA.Horas_Aprobadas, AA.Simbologia,"
    Mi_SQL = Mi_SQL & " AA.Empleado_ID, (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE, Adm_Asistencias AA"
    Mi_SQL = Mi_SQL & " WHERE CE.Empleado_ID = AA.Empleado_ID"
    Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
    Mi_SQL = Mi_SQL & " AND Fecha = '12/01/2015'"
    
    Set Rs_Consulta_Adm_Asistencias = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Asistencias
        If Not .EOF Then
            MDIFrm_Apl_Principal.MousePointer = 11
            Call Encabezado_Reporte("VALIDACION DE HORAS TRABAJADAS", Now, Now)
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
        Else
            MsgBox "No existe información que imprimir", vbInformation + vbOKOnly, Me.Caption
        End If
    End With
    Set Rs_Consulta_Adm_Asistencias = Nothing
    MDIFrm_Apl_Principal.MousePointer = 0
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

Private Sub Txt_Validacion_Justificacion_Horas_Grid_Change()
Dim Columna As Integer
With Grid_Validacion_Horas_Trabajo_Lista
        If .RowSel > 0 Then
            If .TextMatrix(.RowSel, 16) = "FI" Then
                Txt_Validacion_Justificacion_Horas_Grid.Visible = False
                Exit Sub
            End If
            .TextMatrix(.RowSel, 13) = Txt_Validacion_Justificacion_Horas_Grid.Text
            If .TextMatrix(.RowSel, 16) = "HI" And Val(.TextMatrix(.RowSel, 10)) >= Val(.TextMatrix(.RowSel, 22)) Then
                .TextMatrix(.RowSel, 16) = "AS"
                .TextMatrix(.RowSel, 15) = ""
                Manejo_Grid = False
                .Col = 0
                For Columna = 1 To .Cols - 1
                    .Col = Columna
                    .Row = .RowSel
                    .CellBackColor = &HFFFFFF
                Next Columna
                Manejo_Grid = True
                .Col = 10
                .Row = .RowSel
                SendKeys "{END}"
            End If
            If .TextMatrix(.RowSel, 16) = "AS" And Val(.TextMatrix(.RowSel, 10)) < Val(.TextMatrix(.RowSel, 22)) Then
                .TextMatrix(.RowSel, 16) = "HI"
                .TextMatrix(.RowSel, 15) = "Horas_Incompletas"
                Manejo_Grid = False
                .Col = 0
                For Columna = 1 To .Cols - 1
                    .Col = Columna
                    .Row = .RowSel
                    .CellBackColor = &H80FFFF
                Next Columna
                Manejo_Grid = True
                .Col = 10
                .Row = .RowSel
                SendKeys "{END}"
            End If
        End If
    End With
End Sub

Private Sub Txt_Validacion_Justificacion_Horas_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 37 And KeyCode <= 40) Or KeyCode = 13 Then
        'Guarda la informacion y oculta el check
        Txt_Validacion_Justificacion_Horas_Grid.Visible = False
        Call Mover_Control_Grid_Procesos(KeyCode)
        If KeyCode = 37 Or KeyCode = 39 Then Grid_Validacion_Horas_Trabajo_Lista.SetFocus
    End If
End Sub

Private Sub Txt_Validacion_Justificacion_Horas_Grid_KeyPress(KeyAscii As Integer)
    'Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, False)
End Sub
