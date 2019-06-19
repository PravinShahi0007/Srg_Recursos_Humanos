VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Importacion_Archivo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10560
   Begin VB.CommandButton Btn_Limpiar 
      Caption         =   "Limpiar"
      Height          =   690
      Left            =   4695
      Picture         =   "Frm_Adm_Importacion_Archivo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "A"
      Top             =   5655
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Guardar 
      Caption         =   "Archivo Nomipaq"
      Height          =   690
      Left            =   90
      Picture         =   "Frm_Adm_Importacion_Archivo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "A"
      Top             =   5655
      Width           =   1200
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   690
      Left            =   9300
      Picture         =   "Frm_Adm_Importacion_Archivo.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5655
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar Prg_Guardar 
      Height          =   690
      Left            =   1305
      TabIndex        =   27
      Top             =   5655
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   1217
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Cmd_Exportar 
      Left            =   1575
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_Adm_Exportacion_NOI 
      BackColor       =   &H00FFFFFF&
      Height          =   5595
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   10440
      TabIndex        =   3
      Top             =   0
      Width           =   10500
      Begin VB.Frame Fra_Archivo_Importacion_Resultado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultados"
         Height          =   3750
         Left            =   45
         TabIndex        =   24
         Top             =   1740
         Width           =   10350
         Begin MSFlexGridLib.MSFlexGrid Grid_Importacion_Archivo 
            Height          =   3375
            Left            =   45
            TabIndex        =   25
            Top             =   225
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Archivo_Importacion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General"
         Height          =   1335
         Left            =   45
         TabIndex        =   4
         Top             =   405
         Width           =   10350
         Begin VB.Frame P006 
            BackColor       =   &H00FFFFFF&
            Height          =   1005
            Left            =   90
            TabIndex        =   10
            Top             =   1995
            Width           =   10170
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Permisos"
               Height          =   315
               Index           =   5
               Left            =   3884
               TabIndex        =   19
               Top             =   585
               Width           =   1185
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vacaciones"
               Height          =   315
               Index           =   6
               Left            =   5691
               TabIndex        =   18
               Top             =   135
               Width           =   1185
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Retardos"
               Height          =   315
               Index           =   3
               Left            =   2077
               TabIndex        =   17
               Top             =   585
               Width           =   1185
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Faltas"
               Height          =   315
               Index           =   2
               Left            =   2077
               TabIndex        =   16
               Top             =   135
               Width           =   1185
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Horas Triples"
               Height          =   315
               Index           =   1
               Left            =   135
               TabIndex        =   15
               Top             =   585
               Width           =   1320
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Horas Dobles"
               Height          =   315
               Index           =   0
               Left            =   135
               TabIndex        =   14
               Top             =   135
               Width           =   1320
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Dia Doble"
               Height          =   315
               Index           =   4
               Left            =   3884
               TabIndex        =   13
               Top             =   135
               Width           =   1185
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apoyo Alimentacion"
               Height          =   315
               Index           =   7
               Left            =   5700
               TabIndex        =   12
               Top             =   600
               Width           =   1725
            End
            Begin VB.CheckBox Chk_Opciones 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apoyo Transporte"
               Height          =   315
               Index           =   8
               Left            =   7500
               TabIndex        =   11
               Top             =   135
               Width           =   1725
            End
         End
         Begin VB.CommandButton Btn_Importar_Archivo 
            Caption         =   "Generar"
            Height          =   645
            Left            =   9210
            Picture         =   "Frm_Adm_Importacion_Archivo.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "A"
            Top             =   180
            Width           =   1050
         End
         Begin VB.ComboBox Cmb_Adm_Importacion_Empresa 
            Height          =   315
            Left            =   1170
            TabIndex        =   5
            Top             =   225
            Width           =   6975
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Importacion_Fecha_Inicio 
            Height          =   315
            Left            =   1170
            TabIndex        =   7
            Top             =   660
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110166019
            CurrentDate     =   39940
         End
         Begin MSComCtl2.DTPicker Dtp_Adm_Importacion_Fecha_Termino 
            Height          =   315
            Left            =   6240
            TabIndex        =   8
            Top             =   660
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   110166019
            CurrentDate     =   39940
         End
         Begin MSComctlLib.ProgressBar PrgBar_Importacion_Archivo 
            Height          =   165
            Left            =   9210
            TabIndex        =   9
            Top             =   855
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
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
            Left            =   135
            TabIndex        =   23
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Al"
            Height          =   195
            Left            =   4590
            TabIndex        =   22
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label5 
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
            Left            =   135
            TabIndex        =   21
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Lbl_Periodo_Nomina 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1170
            TabIndex        =   20
            Top             =   960
            Width           =   75
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPACTO NOMIPAQ"
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
         TabIndex        =   26
         Top             =   0
         Width           =   3660
      End
   End
End
Attribute VB_Name = "Frm_Adm_Importacion_Archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Renglon_Procesar As Integer 'Indica el renglon actual a procesar para el collapse general del grid de soliictudes pendientes
Dim Collapsing As Boolean       'Indica si se esta haciendo un collpase all en el grid de productos servicios
Public Operacion As String                     'Define la opcion para los procesos
Dim Tipo_Nomina_Empresa As String
Dim Periodo_Inicio As Date
Dim Periodo_Termino As Date

Private Sub Btn_Guardar_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String
On Error GoTo HANDLER

'Verifica que existan registros para el archivo de NomiPAQ
If Grid_Importacion_Archivo.Rows > 0 Then
    Cmd_Exportar.CancelError = True
    Cmd_Exportar.DialogTitle = "Seleccione la ruta destino para el archivo de Movimientos"
    Cmd_Exportar.Flags = cdlOFNHideReadOnly
    Cmd_Exportar.Filter = "Archivos de Movimientos(*.txt)|*.txt"
    Cmd_Exportar.FilterIndex = 2
    Cmd_Exportar.FileName = "MovimientoDYH" & Format(Now, "ddMMyy") & ".txt"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    'valida que el archivo no exista
    If Len(Dir$(Ruta_Exportacion)) > 0 Then
        If MsgBox("El archivo ya existe, Desea Sobreescribirlo ? ", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    Call Generar_Archivo(Ruta_Exportacion)
Else
    MsgBox "No existe información para guardar", vbInformation + vbOKOnly, Me.Caption
End If
Exit Sub

HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Importar_Archivo_Click()
    Dim Cont As Integer                 'Se utiliza para recorrer los check
    Dim Opciones As Boolean             'Verifica que al menos un check este con valor=1
    Opciones = False
    'valida que las opciones de generar archivo sean correctas
    If Cmb_Adm_Importacion_Empresa.ListIndex > -1 Then
        If DateDiff("d", Dtp_Adm_Importacion_Fecha_Inicio.Value, Dtp_Adm_Importacion_Fecha_Termino.Value) >= 0 Then
            Generar_Lista
        Else
            MsgBox "Rango de fechas no valido", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        MsgBox "Seleccione la empresa", vbInformation + vbOKOnly, Me.Caption
        Cmb_Adm_Importacion_Empresa.SetFocus
    End If

End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Generar_Lista
    'DESCRIPCIÓN:           Genera la lista de incidencias que se importaran a Nomipaq
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           19 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Generar_Lista()
Dim Rs_Consulta_Cat_Empleados As rdoResultset               'Información de los empleados
Dim Rs_Consulta_Informacion_Tmp As rdoResultset             'Información de los parametros
Dim Rs_Consulta_Informacion_Tmp_Detalles As rdoResultset    'Información de los parametros
Dim No_Tarjeta As String                                    'No de Tarjeta del empleado
Dim Empleado_ID_tmp As String                               'Identificador del empleado consultado
Dim Nombre_Empleado As String                               'Nombre del empleado consultado
Dim Horas_Tiples_Generadas As Boolean                       'Si se han generado las horas triple
Dim Horas_Dobles_Sumas As Double                           'Suma de las horas dobles para evaluarlas contra las triples
Dim Horas_Laboradas As Double
Dim Dias_Vacaciones As Double                              'Obtiene los dias de vacaciones
Dim Cont_Dias As Integer                                    'Contador para incrementar la fecha
Dim Dia_Vacacion As Date                                    'Se utiliza para avanzar en los dias de vacacion
Dim Tipo_Movimiento As String                               'Informacion del tipo de movimiento, Percepcion, Deduccion, Falta
Dim Columna As Integer
Dim Dias_Habiles As Integer
Dim Fecha_Habil As Date

On Error GoTo HANDLER

PrgBar_Importacion_Archivo.Min = 0
PrgBar_Importacion_Archivo.Value = 0
Grid_Importacion_Archivo.Rows = 0
Grid_Importacion_Archivo.Cols = 8
'Informacion para la barra de progreso
Mi_SQL = "SELECT count(CE.Empleado_ID) as Empleados"
Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE"
Mi_SQL = Mi_SQL & " WHERE CE.Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Empleados
    If Not .EOF Then
        'obtiene la informacion para configurar el progress bar
        If Val(.rdoColumns("Empleados")) > 0 Then
            PrgBar_Importacion_Archivo.Max = Val(.rdoColumns("Empleados"))
        End If
        .Close
    End If
End With
Set Rs_Consulta_Cat_Empleados = Nothing
'Informacion para la lista
Mi_SQL = "SELECT CE.Empleado_ID,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre, CE.Nomipaq_ID, CE.No_Tarjeta, CE.Turno_ID, CE.Trabaja_Domingos,"
Mi_SQL = Mi_SQL & " CT.Hora_Inicio, CT.Hora_Termino, cast((casT(datediff(n,ISNULL(CT.Hora_Inicio,0),ISNULL(CT.Hora_Termino,0)) as Decimal(18,2))/60)as decimal(18,2)) as Horas_Turno"
Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE, Cat_Turnos CT"
Mi_SQL = Mi_SQL & " WHERE CE.Turno_ID = CT.Turno_ID"
Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
Mi_SQL = Mi_SQL & " ORDER BY CE.Apellido_Paterno,CE.Empleado_ID"
Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Empleados
    If Not .EOF Then
        Me.MousePointer = 11
        'obtiene los dias habiles del periodo seleccionado
        Fecha_Habil = Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy")
        Dias_Habiles = 0
        'Obtiene los dias del periodo seleccionado
        For Cont_Dias = 1 To DateDiff("d", Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy"), _
                            Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy")) + 1
            'Verifica que no sea dia Feriado
            If Conectar_Ayudante.Es_Dia_No_Laboral(Fecha_Habil) = False Then
                'Verifica que no sea Sabado o Domingo
                If (Weekday(Fecha_Habil) <> vbSunday) And (Weekday(Fecha_Habil) <> vbSaturday) Then
                    Dias_Habiles = Dias_Habiles + 1
                End If
            End If
            Fecha_Habil = DateAdd("d", 1, Fecha_Habil)
        Next
        'Agrega el encabezado
        Grid_Importacion_Archivo.AddItem "" & Chr(9) & "ID Nomipaq" & Chr(9) & "Empleado_ID" & Chr(9) & "Nombre" & Chr(9) & "Dias Periodo" & Chr(9) & "Dias Trab."
        PrgBar_Importacion_Archivo.Visible = True
        PrgBar_Importacion_Archivo.Value = 0
        While Not .EOF
            Debug.Print Empleado_ID_tmp
            If Empleado_ID_tmp <> .rdoColumns("Empleado_ID") Then
                'Agrega el encabezado del empleado al grid
                No_Tarjeta = .rdoColumns("Nomipaq_ID")
                Empleado_ID_tmp = .rdoColumns("Empleado_ID")
                Nombre_Empleado = .rdoColumns("Nombre")
                Horas_Tiples_Generadas = False
                'Obtiene las horas aprobadas del turno
                Mi_SQL = "SELECT ISNULL(SUM(AA.Horas_Aprobadas),0) as Horas_Aprobadas"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA"
                Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
                Mi_SQL = Mi_SQL & " AND AA.Fecha > = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
                Mi_SQL = Mi_SQL & " AND AA.Fecha < = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
                Mi_SQL = Mi_SQL & " AND ISNULL(AA.Horas_Aprobadas,0)>0"
                Horas_Laboradas = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Horas_Aprobadas")
                If Val(.rdoColumns("Horas_Turno")) > 0 Then
                    Horas_Laboradas = Round(Horas_Laboradas / Val(.rdoColumns("Horas_Turno")), 2)
                Else
                    Horas_Laboradas = 0
                End If
                Grid_Importacion_Archivo.AddItem "-" & Chr(9) & .rdoColumns("Nomipaq_ID") & Chr(9) & _
                    .rdoColumns("Empleado_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & _
                    Dias_Habiles & Chr(9) & _
                    Horas_Laboradas & Chr(9) & "" & Chr(9) & _
                    ""
                For Columna = 1 To Grid_Importacion_Archivo.Cols - 1
                    Grid_Importacion_Archivo.Col = Columna
                    Grid_Importacion_Archivo.Row = Grid_Importacion_Archivo.Rows - 1
                    Grid_Importacion_Archivo.CellBackColor = &H8000000F
                Next Columna
                Grid_Importacion_Archivo.AddItem " " & Chr(9) & .rdoColumns("Nomipaq_ID") & Chr(9) & .rdoColumns("Empleado_ID") & Chr(9) & "Movimiento" & Chr(9) & "Simbologia" & Chr(9) & "SubSimbologia" & Chr(9) & "Cantidad"
            End If
            
            'Busca cada una de las opciones seleccionadas
            'Obtiene las hras dobles
            'If Chk_Opciones(0).Value = 1 Then
            Mi_SQL = "SELECT ISNULL(SUM(AA.Horas_Extra),0) as Horas_Extra"
            Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA"
            Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
            Mi_SQL = Mi_SQL & " AND AA.Fecha > = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND AA.Fecha < = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND ISNULL(AA.Horas_Extra,0)>0"
            'Mi_SQL = Mi_SQL & " ORDER BY AA.Fecha"
            Set Rs_Consulta_Informacion_Tmp = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Informacion_Tmp
                If Not .EOF Then
                    Horas_Dobles_Sumas = 0
                    While Not .EOF
                        If Val(.rdoColumns("Horas_Extra")) > 0 Then
                            Horas_Dobles_Sumas = .rdoColumns("Horas_Extra")
                            If Horas_Dobles_Sumas > Horas_Triples Then
                                Horas_Tiples_Generadas = Horas_Dobles_Sumas - Horas_Dobles
                                Grid_Importacion_Archivo.AddItem " " & Chr(9) & "ID Nomipaq" & Chr(9) & "Empleado_ID" & Chr(9) & "Hrs. Extra Doble" & Chr(9) & "HE2" & Chr(9) & "" & Chr(9) & Horas_Dobles
                                If Chk_Opciones(1).Value = 1 Then
                                    Horas_Tiples_Generadas = True
                                    Grid_Importacion_Archivo.AddItem " " & Chr(9) & "ID Nomipaq" & Chr(9) & "Empleado_ID" & Chr(9) & "Hrs. Extra Triple" & Chr(9) & "HE2" & Chr(9) & "" & Chr(9) & Horas_Tiples_Generadas
                                End If
                            Else
                                Grid_Importacion_Archivo.AddItem " " & Chr(9) & "ID Nomipaq" & Chr(9) & "Empleado_ID" & Chr(9) & "Hrs. Extra Doble" & Chr(9) & "HE2" & Chr(9) & "" & Chr(9) & Horas_Dobles_Sumas
                            End If
                        End If
                        .MoveNext
                    Wend
                    .Close
                End If
            End With
            Set Rs_Consulta_Informacion_Tmp = Nothing
            'Obtiene todas las incidencias
            Mi_SQL = "SELECT Count(Simbologia) as Cantidad, "
            'Mi_SQL = Mi_SQL & " ISNULL(Referencia,'') as Referencia, "
            Mi_SQL = Mi_SQL & " ISNULL(Tipo_Incidencia,'') as Tipo_Incidencia, "
            Mi_SQL = Mi_SQL & " Simbologia, Subsimbologia"
            Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA"
            Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
            Mi_SQL = Mi_SQL & " AND AA.Fecha > = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND AA.Fecha < = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " GROUP BY Tipo_Incidencia, Simbologia, Subsimbologia"
            Set Rs_Consulta_Informacion_Tmp = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Informacion_Tmp.EOF Then
                While Not Rs_Consulta_Informacion_Tmp.EOF
                        Grid_Importacion_Archivo.AddItem " " & Chr(9) & "ID Nomipaq" & Chr(9) & "Empleado_ID" & Chr(9) & "" & Chr(9) & Rs_Consulta_Informacion_Tmp.rdoColumns("Simbologia") & Chr(9) & Rs_Consulta_Informacion_Tmp.rdoColumns("SubSimbologia") & Chr(9) & Rs_Consulta_Informacion_Tmp.rdoColumns("Cantidad")
                    Rs_Consulta_Informacion_Tmp.MoveNext
                Wend
                Rs_Consulta_Informacion_Tmp.Close
            End If
            Set Rs_Consulta_Informacion_Tmp = Nothing
            'Calcula la informacion de ayuda de alimentacion y transporte
            PrgBar_Importacion_Archivo.Value = PrgBar_Importacion_Archivo.Value + 1
            .MoveNext
        Wend
    End If
End With
'Configuracion del grid
With Grid_Importacion_Archivo
    If .Rows > 1 Then .FixedRows = 1
        .ColWidth(0) = 600 'Signo
        .ColWidth(1) = 0    'ID NOI
        .ColWidth(2) = 0   'EMpleado ID
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(3) = 3500 'Referencia
        .ColAlignment(3) = flexAlignLeftCenter
        .ColWidth(4) = 1200 'Simbologia
        .ColWidth(5) = 1200  'Subsimbologia
        .ColAlignment(5) = flexAlignRightCenter
        .ColWidth(6) = 1000  'Cantidad
        .ColAlignment(6) = flexAlignRightCenter
        .ColWidth(7) = 1000  'Al
        .ColAlignment(7) = flexAlignRightCenter
End With
Collapsing = True
Call Collapse_Grid
Collapsing = False
If Grid_Importacion_Archivo.Rows > 1 Then
    PrgBar_Importacion_Archivo.Visible = False
Else
    MsgBox "No existe informacion con los parametros seleccionados", vbInformation + vbOKOnly, Me.Caption
End If
Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    PrgBar_Importacion_Archivo.Visible = False
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Inicializa()
    Select Case Operacion
        Case "Exportacion_Nomipaq":
            Dtp_Adm_Importacion_Fecha_Inicio.Value = Now
            Dtp_Adm_Importacion_Fecha_Termino.Value = Now
            Call Conectar_Ayudante.Llena_Combo_Item("Empresa_ID, Nombre", "Cat_Empresas", Cmb_Adm_Importacion_Empresa, 0, "Nombre")
    End Select
End Sub

Private Sub Btn_Limpiar_Click()
    Dim Cont As Integer                 'Contador para recorre los check
    Grid_Importacion_Archivo.Rows = 0
    Cmb_Adm_Importacion_Empresa.ListIndex = -1
    Dtp_Adm_Importacion_Fecha_Inicio.Value = Now
    Dtp_Adm_Importacion_Fecha_Termino.Value = Now
    For Cont = 0 To 6
        Chk_Opciones(Cont).Value = 0
    Next
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Generar_Archivo
    'DESCRIPCIÓN:           Genera el archivo para Nomipaq
    'PARÁMETROS :           Archivo:Ubucación y Nombre del Archivo que se generará
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           25 Junio 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Generar_Archivo(Archivo As String)
Dim Tipo_Movimiento As String
Dim Tipo_Falta As String
Dim Rs_Consulta_Cat_Empleados As rdoResultset               'Información de los empleados
Dim Rs_Consulta_Informacion_Tmp As rdoResultset             'Información de los parametros
Dim Rs_Consulta_Informacion_Tmp_Detalles As rdoResultset    'Información de los parametros
Dim No_Tarjeta As String                                    'No de Tarjeta del empleado
Dim Empleado_ID_tmp As String                               'Identificador del empleado consultado
Dim Nombre_Empleado As String                               'Nombre del empleado consultado
Dim Horas_Tiples_Generadas As Boolean                       'Si se han generado las horas triple
Dim Horas_Dobles_Sumas As Double                           'Suma de las horas dobles para evaluarlas contra las triples
Dim Horas_Laboradas As Double
Dim Dias_Vacaciones As Double                              'Obtiene los dias de vacaciones
Dim Cont_Dias As Integer                                    'Contador para incrementar la fecha
Dim Dia_Vacacion As Date                                    'Se utiliza para avanzar en los dias de vacacion
Dim Columna As Integer
Dim Dias_Habiles As Integer
Dim Fecha_Habil As Date
Dim Fecha_Aplicar As Date                                   'Fecha en que se aplicara la incidencia
Dim Tipo_Nomina_Empresa As String
Dim Fecha_Correcta As Boolean                               'Define si la fecha esta bien para aplicar

Dim Archivo_Abierto As Boolean      'Indica si el archivo de NomiPAQ esta abierto
Dim Cont_Fila As Integer            'Contador para recorrer las filas del grid
Dim Clave_NomiPAQ As String             'Indica la clave del empleado para el ENCABEZADO
Dim Periodo As Integer              'Periodo en que se capturara la informacion
On Error GoTo HANDLER

Mi_SQL = "SELECT count(CE.Empleado_ID) as Empleados"
Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE"
Mi_SQL = Mi_SQL & " WHERE CE.Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Empleados
    If Not .EOF Then
        'obtiene la informacion para configurar el progress bar
        If Val(.rdoColumns("Empleados")) > 0 Then
            Prg_Guardar.Max = Val(.rdoColumns("Empleados"))
        End If
        .Close
    End If
End With
Set Rs_Consulta_Cat_Empleados = Nothing
'Informacion para la lista
Mi_SQL = "SELECT CE.Empleado_ID,(CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre, CE.Nomipaq_ID, CE.No_Tarjeta, CE.Turno_ID, CE.Trabaja_Domingos,"
Mi_SQL = Mi_SQL & " CT.Hora_Inicio, CT.Hora_Termino, cast((casT(datediff(n,ISNULL(CT.Hora_Inicio,0),ISNULL(CT.Hora_Termino,0)) as Decimal(18,2))/60)as decimal(18,2)) as Horas_Turno"
Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE, Cat_Turnos CT"
Mi_SQL = Mi_SQL & " WHERE CE.Turno_ID = CT.Turno_ID"
Mi_SQL = Mi_SQL & " AND CE.Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
Mi_SQL = Mi_SQL & " AND CE.Estatus ='A'"
Mi_SQL = Mi_SQL & " ORDER BY CE.Apellido_Paterno,CE.Empleado_ID"
Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Cat_Empleados
    If Not .EOF Then
        Mi_SQL = "SELECT Tipo_Nomina FROM Cat_Empresas WHERE Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
        Tipo_Nomina_Empresa = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Tipo_Nomina")
        Clave_NomiPAQ = ""
        'Abre el archivo para su llenado
        Open Archivo For Output As #1
        Archivo_Abierto = True
        Me.MousePointer = 11
        'obtiene los dias habiles del periodo seleccionado
        Fecha_Habil = Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy")
        Dias_Habiles = 0
        'Agrega el encabezado
        Prg_Guardar.Visible = True
        Prg_Guardar.Value = 0
        While Not .EOF
            'Agrega el encabezado por empleado
            If Trim(Clave_NomiPAQ) <> Trim(.rdoColumns("Nomipaq_ID")) Then
                Print #1, "E"; Spc(7); Conectar_Ayudante.Agregar_Espacios(.rdoColumns("Nomipaq_ID"), 107 - (Len("E") + Len(.rdoColumns("Nomipaq_ID")) + 7))
                Clave_NomiPAQ = Trim(.rdoColumns("Nomipaq_ID"))
            End If
            'Obtiene el periodo
            'Periodo = Obtiene_Periodo_Quincena(.TextMatrix(Cont_Fila, 4))
            Mi_SQL = "SELECT ISNULL(SUM(AA.Horas_Extra),0) as Horas_Extra, AA.Fecha"
            Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA"
            Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
            Mi_SQL = Mi_SQL & " AND AA.Fecha > = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND AA.Fecha < = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND ISNULL(AA.Horas_Extra,0)>0"
            Mi_SQL = Mi_SQL & " GROUP BY AA.Fecha"
            Mi_SQL = Mi_SQL & " ORDER BY AA.Fecha"
            Set Rs_Consulta_Informacion_Tmp = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Informacion_Tmp
                If Not .EOF Then
                    Horas_Dobles_Sumas = 0
                    While Not .EOF
                        Fecha_Aplicar = .rdoColumns("Fecha")
                        If Val(.rdoColumns("Horas_Extra")) > 0 Then
                            Horas_Dobles_Sumas = .rdoColumns("Horas_Extra")
                            If Horas_Dobles_Sumas > Horas_Triples Then
                                Horas_Tiples_Generadas = Horas_Dobles_Sumas - Horas_Dobles
                                'Agrega las horas dobles al archivo
                                Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                    Conectar_Ayudante.Alinea_Derecha(CStr(Fecha_Aplicar), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                    Trim(PDF_Horas_Dobles); _
                                    Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(PDF_Horas_Dobles)))); Spc(1); _
                                    Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                    Year(Fecha_Aplicar); "                           "
                                    
                                Horas_Tiples_Generadas = True
                                Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                    Conectar_Ayudante.Alinea_Derecha(CStr(Fecha_Aplicar), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                    Trim(PDF_Horas_Triples); _
                                    Conectar_Ayudante.Alinea_Derecha(Format(Horas_Tiples_Generadas, "#.00"), 49 - (Len(Trim(PDF_Horas_Triples)))); Spc(1); _
                                    Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                    Year(Fecha_Aplicar); "                           "
                            Else
                                Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                    Conectar_Ayudante.Alinea_Derecha(CStr(Fecha_Aplicar), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                    Trim(PDF_Horas_Dobles); _
                                    Conectar_Ayudante.Alinea_Derecha(Format(Horas_Dobles_Sumas, "#.00"), 49 - (Len(Trim(PDF_Horas_Dobles)))); Spc(1); _
                                    Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                    Year(Fecha_Aplicar); "                           "
                            End If
                        End If
                        .MoveNext
                    Wend
                    .Close
                End If
            End With
            Set Rs_Consulta_Informacion_Tmp = Nothing
            'Obtiene todas las incidencias
            Mi_SQL = "SELECT "
            'Mi_SQL = Mi_SQL & " ISNULL(Referencia,'') as Referencia, "
            Mi_SQL = Mi_SQL & " ISNULL(Tipo_Incidencia,'') as Tipo_Incidencia, "
            Mi_SQL = Mi_SQL & " Simbologia, Subsimbologia, AA.Fecha"
            Mi_SQL = Mi_SQL & " FROM Adm_Asistencias AA"
            Mi_SQL = Mi_SQL & " WHERE AA.Empleado_ID = '" & .rdoColumns("Empleado_ID") & "'"
            Mi_SQL = Mi_SQL & " AND AA.Fecha > = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " AND AA.Fecha < = " & Par_Fecha & Format(Dtp_Adm_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha
            Mi_SQL = Mi_SQL & " GROUP BY AA.Fecha, Tipo_Incidencia, Simbologia, Subsimbologia"
            Set Rs_Consulta_Informacion_Tmp = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Informacion_Tmp.EOF Then
                While Not Rs_Consulta_Informacion_Tmp.EOF
                    Fecha_Aplicar = DateAdd("d", 15, Rs_Consulta_Informacion_Tmp.rdoColumns("Fecha"))
                    'Agrega las incidencias a NOI
                    Tipo_Movimiento = ""
                    Tipo_Falta = ""
                    Select Case Rs_Consulta_Informacion_Tmp.rdoColumns("Simbologia")
                        Case "VN"
                            Clave_NomiPAQ = PDF_Vacaciones
                            Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                Trim(Clave_NomiPAQ); _
                                Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                Year(Fecha_Aplicar); "                           "
                                
                        Case "PE":
                            'Definir si es con goce o sin goce
                            Clave_NomiPAQ = PDF_Permiso_CG
                            Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                Trim(Clave_NomiPAQ); _
                                Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                Year(Fecha_Aplicar); "                           "

                        Case "II":
                            Select Case Rs_Consulta_Informacion_Tmp.rdoColumns("SubSimbologia")
                                Case "EG"
                                    Clave_NomiPAQ = PDF_Enfermedad_General
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "

                                Case "MA"
                                    Clave_NomiPAQ = PDF_Maternidad
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                                Case "RT"
                                    Clave_NomiPAQ = PDF_Riesgo_Trabajo
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
'
                            End Select

                        Case "ID"
                            Select Case Rs_Consulta_Informacion_Tmp.rdoColumns("SubSimbologia")
                                Case "VA"
                                    Clave_NomiPAQ = PDF_Vacaciones
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                                Case "AL"
                                    Clave_NomiPAQ = PDF_Alumbramiento
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                                Case "DE"
                                    Clave_NomiPAQ = PDF_Defuncion
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                                Case "MO"
                                    Clave_NomiPAQ = PDF_Matrimonio
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                            End Select

                        Case "FJ"
                            Clave_NomiPAQ = PDF_Falta_Justificada
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                        Case "FI"
                            Clave_NomiPAQ = PDF_Falta_InJustificada
                                    Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                        Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                        Trim(Clave_NomiPAQ); _
                                        Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                        Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                        Year(Fecha_Aplicar); "                           "
                        Case Else
                            Clave_NomiPAQ = NOI_ID(Rs_Consulta_Informacion_Tmp.rdoColumns("Simbologia"))
                            Print #1, "D " & Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)); _
                                Conectar_Ayudante.Alinea_Derecha(CStr(Periodo), 26 - Len(Conectar_Ayudante.Agregar_Espacios(Tipo_Nomina, 22 - Len(Tipo_Nomina)))); Spc(1); "2"; Spc(19); _
                                Trim(Clave_NomiPAQ); _
                                Conectar_Ayudante.Alinea_Derecha(Format("1", "#.00"), 49 - (Len(Trim(Clave_NomiPAQ)))); Spc(1); _
                                Format(Fecha_Aplicar, "dd/MM/yyyy"); Spc(5); _
                                Year(Fecha_Aplicar); "                           "

                    End Select
                    Rs_Consulta_Informacion_Tmp.MoveNext
                Wend
                Rs_Consulta_Informacion_Tmp.Close
            End If
            Prg_Guardar.Value = Prg_Guardar.Value + 1
            .MoveNext
        Wend
    End If
End With
Close #1
Archivo_Abierto = False
Prg_Guardar.Visible = False
Me.MousePointer = 0
MsgBox "El archivo se ha generado exitosamente en:" & vbCrLf & Archivo, vbInformation + vbOKOnly, Me.Caption
Exit Sub
HANDLER:
    If Archivo_Abierto Then Close #1
    Me.MousePointer = 0
    PrgBar_Importacion_Archivo.Visible = False
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
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
    If Grid_Importacion_Archivo.Rows > 0 Then
        Grid_Importacion_Archivo.FixedRows = 1
        For Renglon_Procesar = 1 To Grid_Importacion_Archivo.Rows - 1
            If Grid_Importacion_Archivo.TextMatrix(Renglon_Procesar, 0) = "-" Then
                Grid_Importacion_Archivo.Col = 0
                Call Grid_Importacion_Archivo_Click
            End If
        Next Renglon_Procesar
    End If
End Sub

Private Sub Cmb_Adm_Importacion_Empresa_Click()
'Obtiene si la empresa es de nomina quincenal o semanal
If Cmb_Adm_Importacion_Empresa.ListIndex > -1 Then
    calcula_semana_laboral
End If
End Sub

Private Sub Dtp_Adm_Importacion_Fecha_Inicio_Change()
    calcula_semana_laboral
End Sub

Private Sub Dtp_Adm_Importacion_Fecha_Inicio_Click()
    calcula_semana_laboral
End Sub

Private Sub Dtp_Adm_Importacion_Fecha_Inicio_KeyPress(KeyAscii As Integer)
    calcula_semana_laboral
End Sub

Private Sub Grid_Importacion_Archivo_Click()
Dim Renglon As Integer     'Indica que renglon se esta consulltando
Dim Fila As Integer        'Contador de filas
If Grid_Importacion_Archivo.Rows > 1 Then
    If Grid_Importacion_Archivo.Col <= 1 And Grid_Importacion_Archivo.Row > 0 Then
        'And Grid_Importacion_Archivo.TextMatrix(Renglon, 1) = "-"
        If Collapsing = False Then
            Renglon = Grid_Importacion_Archivo.MouseRow
        Else
            Renglon = Renglon_Procesar
        End If
        If Renglon < 1 Then Exit Sub
        
        While Renglon > 0 And Trim(Grid_Importacion_Archivo.TextMatrix(Renglon, 0)) = ""
            Renglon = Renglon - 1
        Wend
        If Grid_Importacion_Archivo.TextMatrix(Renglon, 0) = "-" Then
            Grid_Importacion_Archivo.TextMatrix(Renglon, 0) = "+"
        Else
            Grid_Importacion_Archivo.TextMatrix(Renglon, 0) = "-"
        End If
        
        Renglon = Renglon + 1
        If Renglon < Grid_Importacion_Archivo.Rows Then
            If Grid_Importacion_Archivo.RowHeight(Renglon) = 0 Then
                Do While Trim(Grid_Importacion_Archivo.TextMatrix(Renglon, 0)) = ""
                    Grid_Importacion_Archivo.RowHeight(Renglon) = -1
                    Renglon = Renglon + 1
                    If Renglon >= Grid_Importacion_Archivo.Rows Then Exit Do
                Loop
            Else
                Do While Trim(Grid_Importacion_Archivo.TextMatrix(Renglon, 0)) = ""
                    Grid_Importacion_Archivo.RowHeight(Renglon) = 0
                    Renglon = Renglon + 1
                    If Renglon >= Grid_Importacion_Archivo.Rows Then Exit Do
                Loop
            End If
        End If
        Grid_Importacion_Archivo.Col = 0
    End If
End If

End Sub

Private Sub calcula_semana_laboral()
    If Cmb_Adm_Importacion_Empresa.ListIndex > -1 Then
        Mi_SQL = "SELECT Tipo_Nomina FROM Cat_Empresas WHERE Empresa_ID = '" & Format(Cmb_Adm_Importacion_Empresa.ItemData(Cmb_Adm_Importacion_Empresa.ListIndex), "00000") & "'"
        Tipo_Nomina_Empresa = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Tipo_Nomina")
        Select Case Tipo_Nomina_Empresa
            Case "QUINCENAL"
                Periodo_Inicio = DateAdd("d", 15, Dtp_Adm_Importacion_Fecha_Inicio.Value)
                Dtp_Adm_Importacion_Fecha_Termino.Value = DateAdd("d", 14, Dtp_Adm_Importacion_Fecha_Inicio.Value)
                Periodo_Termino = DateAdd("d", 15, Dtp_Adm_Importacion_Fecha_Termino.Value)
            Case "SEMANAL"
                Periodo_Inicio = DateAdd("d", 7, Dtp_Adm_Importacion_Fecha_Inicio.Value)
                Dtp_Adm_Importacion_Fecha_Termino.Value = DateAdd("d", 6, Dtp_Adm_Importacion_Fecha_Inicio.Value)
                Periodo_Termino = DateAdd("d", 7, Dtp_Adm_Importacion_Fecha_Termino.Value)
            Case Else
                MsgBox "Para la empresa seleccionada no esta definido el tipo de nomina", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
        End Select
        'Lbl_Periodo_Nomina.Caption = "PERIODO A APLICAR: " & Format(Periodo_Inicio, "dd MMM yyyy") & " AL " & Format(Periodo_Termino, "dd MMM yyyy")
    End If
End Sub

Private Function NOI_ID(Simbologia As String) As String
Dim Rs_Consulta_Cat_Tipos_Faltas As rdoResultset
    'Consulta el ID de por clave
    Mi_SQL = "SELECT Codigo_NOI From Cat_Tipos_Faltas"
    Mi_SQL = Mi_SQL & " WHERE SImbologia = '" & Simbologia & "'"
    Set Rs_Consulta_Cat_Tipos_Faltas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Tipos_Faltas.EOF Then
        NOI_ID = Rs_Consulta_Cat_Tipos_Faltas.rdoColumns("Codigo_NOI")
    End If
    Set Rs_Consulta_Cat_Tipos_Faltas = Nothing
End Function

