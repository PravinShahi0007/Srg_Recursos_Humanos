VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Historico_Vacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico de Vacaciones"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleMode       =   0  'User
   ScaleWidth      =   7950
   Begin VB.PictureBox Pic_Reportes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   -120
      ScaleHeight     =   8895
      ScaleWidth      =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.TextBox Txt_No_Tarjeta_Historico_Vacaciones 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   1080
         Width           =   6015
      End
      Begin VB.ComboBox Cmb_Historico_Vacaciones_Empleados 
         Height          =   315
         ItemData        =   "Frn_Adm_Historico_Vacaciones.frx":0000
         Left            =   1320
         List            =   "Frn_Adm_Historico_Vacaciones.frx":000A
         TabIndex        =   12
         Top             =   1440
         Width           =   5970
      End
      Begin VB.CommandButton Btn_Imprimir 
         Caption         =   "Imprimir"
         Height          =   570
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "A"
         Top             =   6840
         Width           =   1305
      End
      Begin VB.Frame Fra_Adm_Historico_Vacaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4800
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   7305
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Historico_Vacaciones 
            Height          =   4320
            Left            =   75
            TabIndex        =   10
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   7620
            _Version        =   393216
            Rows            =   0
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Btn_Exportar 
         Caption         =   "Exportar Excel"
         Height          =   555
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "A"
         Top             =   6840
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "C"
         Top             =   6840
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6840
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker Dt_Rpt_Ingresos_De 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   124321793
         CurrentDate     =   42373
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   480
         Top             =   6840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker Dt_Rpt_Ingresos_A 
         Height          =   315
         Left            =   5400
         TabIndex        =   8
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   124321793
         CurrentDate     =   42373
      End
      Begin VB.Label Lbl_No_Tarjeta 
         BackColor       =   &H8000000E&
         Caption         =   "No. Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Empleado"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Lbl_Hostorico_Vacaciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "HISTORICO VACACIONES"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   0
         Width           =   4725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Ingresos de:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "A:"
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frm_Adm_Historico_Vacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
Dt_Rpt_Ingresos_De.Value = "01/01/2000"
Dt_Rpt_Ingresos_A.Value = Now
Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, Nombre+' '+ Apellido_Paterno+' ' + Apellido_Materno as Nombre", "Cat_Empleados WHERE Estatus='A'", Cmb_Historico_Vacaciones_Empleados, 0, "Nombre", "", False, "")
' Call Conectar_Ayudante.Llena_Combo_Item("Turno_ID, Nombre", "Cat_Turnos", Cmb_Filtro_Empleado_Turno, 0, "Turno_ID", "", False, "")

End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Historico_Vacaciones
    'DESCRIPCIÓN:           Consulta los empleados y su histórico de vacaciones
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Historico_Vacaciones()
Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
    Grid_Adm_Historico_Vacaciones.Rows = 0
Dim Total_Años As Integer
Dim Dias_Totales As Integer
'Total_Años = Calculos_Años
Dim Columnas As String
Dim Columnas_Excel As String
Dim Columnas_Impresión As String
    Dim Dia_Actual  As String
    Dia_Actual = Now
    'Consulta los datos generales del usuario
    Mi_SQL = "select Cat_Empleados.Empleado_ID, No_Tarjeta, Cat_Empleados.Nombre+' '+ Apellido_Paterno+' '+ Apellido_Materno as Nomre_del_empleado, "
    Mi_SQL = Mi_SQL & " Cat_Puestos.Nombre as Puesto, Cat_Departamentos.Nombre as Departamento, '' as grupo, (Fecha_Ingreso+(365.5*2)) as Fecha_1,  Fecha_Ingreso, "
    Mi_SQL = Mi_SQL & " Cat_Empleados.Tipo_Empleado, DATEDIFF(DAY, Fecha_Ingreso, '" & Format(Dia_Actual, "mm/dd/yyyy") & "') as Dias_Trabajados, "
    Mi_SQL = Mi_SQL & Year(Now) & "-YEAR(Fecha_Ingreso) as Años, Tipo_Empleado"
'    Mi_SQL = Mi_SQL & ",ISNULL( SUM (Adm_Movimientos_Vacaciones.Dias), 0) as Total_Tomados "
    Mi_SQL = Mi_SQL & " ,ISNULL(SUM(Adm_Movimientos_Asistencias.Dias_Permiso), 0) AS Total_Tomados "
'    Mi_SQL = Mi_SQL & " From Cat_Empleados left outer JOIN Adm_Movimientos_Vacaciones on Cat_Empleados.Empleado_ID = Adm_Movimientos_Vacaciones.Empleado_Id, "
    Mi_SQL = Mi_SQL & " From Cat_Empleados left outer JOIN Adm_Movimientos_Asistencias on Cat_Empleados.Empleado_ID = Adm_Movimientos_Asistencias.Empleado_Id, "
    Mi_SQL = Mi_SQL & " Cat_Puestos, Cat_Departamentos "
    Mi_SQL = Mi_SQL & " Where Cat_Empleados.Puesto_ID = Cat_Puestos.Puesto_ID "
    Mi_SQL = Mi_SQL & " AND Cat_Departamentos.Departamento_ID = Cat_Empleados.Departamento_ID "
'    Mi_SQL = Mi_SQL & " AND (Fecha_Ingreso between '" & Format(Dt_Rpt_Ingresos_De.Value, "mm/dd/yyyy") & "' and '" & Format(Dt_Rpt_Ingresos_A.Value, "mm/dd/yyyy") & "') "
    Mi_SQL = Mi_SQL & "AND (Fecha_Inicio >= '" & Format(Dt_Rpt_Ingresos_De.Value, "mm/dd/yyyy") & "' AND  Fecha_Termino <= '" & Format(Dt_Rpt_Ingresos_A.Value, "mm/dd/yyyy") & "')"
    If Cmb_Historico_Vacaciones_Empleados.Text <> "" And Cmb_Historico_Vacaciones_Empleados.Text <> "TODOS" Then
        Mi_SQL = Mi_SQL & " AND Cat_Empleados.Empleado_ID = '" & Format(Cmb_Historico_Vacaciones_Empleados.ItemData(Cmb_Historico_Vacaciones_Empleados.ListIndex), "00000") & "'"
    End If
'    Mi_SQL = Mi_SQL & "  and No_Tarjeta >= 50 and No_Tarjeta < 60 "
    Mi_SQL = Mi_SQL & " and (Adm_Movimientos_Asistencias.Tipo_Falta_ID = 4 or Tipo_Falta_ID = 10) "
    Mi_SQL = Mi_SQL & " group By Cat_Empleados.Empleado_ID, No_Tarjeta, Cat_Empleados.Nombre,Apellido_Paterno,Apellido_Materno,"
     Mi_SQL = Mi_SQL & " Cat_Puestos.Nombre , Cat_Departamentos.Nombre, Fecha_Ingreso, Tipo_Empleado "
    Mi_SQL = Mi_SQL & " order BY No_Tarjeta"
    
    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Empleados
        If Not .EOF Then
         Dim Tipo_Empleado_Aux As String
        Tipo_Empleado_Aux = .rdoColumns("Tipo_Empleado")
        If Trim(UCase(.rdoColumns("Tipo_Empleado"))) <> "CONFIANZA" And Trim(UCase(.rdoColumns("Tipo_Empleado"))) <> "SINDICALIZADO" Then
            Tipo_Empleado_Aux = "otros"
        End If
        Total_Años = .rdoColumns("Años")
'            Columnas = "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Puesto" & Chr(9) & "Departamento" & Chr(9) & "Fecha" & Chr(9) & "Ingreso" & Chr(9) & "Tipo_Empleado" & Chr(9) & "Dias Trabajados" & Chr(9) & "Empleado_ID"
            Columnas = "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Tipo Empleado" & Chr(9) & "Empleado_ID"
            Columnas_Excel = "No Tarjeta|" & Chr(9) & "Nombre|" & Chr(9) & "Puesto|" & Chr(9) & "Departamento|" & Chr(9) & "Fecha|" & Chr(9) & "Ingreso|" & Chr(9) & "Tipo_Empleado|" & Chr(9) & "Dias Trabajados|"
            Columnas_Impresión = "No Tarjeta"
            If Total_Años > 0 Then
                For I = 1 To Total_Años
'                Columnas = Columnas & Chr(9) & "Días año " & I & Chr(9) & " Restante"
                Columnas_Excel = Columnas_Excel & Chr(9) & "Días año " & I & Chr(9) & "| Restante|"
                Next I
                Columnas = Columnas & Chr(9) & " Dias Totales" & Chr(9) & " Dias Tomados " & Chr(9) & " Saldo "
                Columnas_Excel = Columnas_Excel & Chr(9) & "Días Totales|" & Chr(9) & " Dias Tomados|" & Chr(9) & " Saldo|"
                Columnas_Impresión = Columnas_Impresión & Chr(9) & " D. Totales" & Chr(9) & " D. Tomados" & Chr(9) & " Saldo"
            End If
            Call Encabezado_Reporte("HISTÓRICO DE VACACIONES", DateAdd("s", 1, Dt_Rpt_Ingresos_De.Value), DateAdd("s", 1, Dt_Rpt_Ingresos_A.Value))
            
            Grid_Adm_Historico_Vacaciones.AddItem Columnas
            
            Dim Cadena_Columnas_Desglose_Vacaciones As String
            Cadena_Columnas_Desglose_Vacaciones = Columnas_Desglose_Vacaciones
            If Trim(Cadena_Columnas_Desglose_Vacaciones) <> "" Then
            Columnas_Excel = Columnas_Excel + Formato_Meses_Columnas_Vacaciones(Cadena_Columnas_Desglose_Vacaciones)
            End If
            Print #1, Columnas_Impresión
            Print #2, Columnas_Excel
            
            While Not .EOF
            Dim Numero_Días As Double
            Dim Dias_Restantes As Double
                Dim Cadena_Datos As String
                Dim Cadena_Datos_Excel As String
                Dim Cadena_Datos_Impresión As String
'                Cadena_Datos = .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nomre_del_empleado") & Chr(9) & .rdoColumns("Puesto") & Chr(9) & .rdoColumns("Departamento") & Chr(9) & .rdoColumns("Fecha_1") & Chr(9) & .rdoColumns("Fecha_Ingreso") & Chr(9) & .rdoColumns("Tipo_Empleado") & Chr(9) & .rdoColumns("Dias_Trabajados") & Chr(9) & .rdoColumns("Empleado_ID")
                Cadena_Datos = .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nomre_del_empleado") & Chr(9) & .rdoColumns("Tipo_Empleado") & Chr(9) & .rdoColumns("Empleado_ID")
                Cadena_Datos_Excel = .rdoColumns("No_Tarjeta") & "|" & Chr(9) & .rdoColumns("Nomre_del_empleado") & "|" & Chr(9) & .rdoColumns("Puesto") & "|" & Chr(9) & .rdoColumns("Departamento") & "|" & Chr(9) & .rdoColumns("Fecha_1") & "|" & Chr(9) & .rdoColumns("Fecha_Ingreso") & "|" & Chr(9) & .rdoColumns("Tipo_Empleado") & "|" & Chr(9) & .rdoColumns("Dias_Trabajados") & "|"
                Cadena_Datos_Impresión = "     " & .rdoColumns("No_Tarjeta")
                Dias_Restantes = Val(.rdoColumns("Dias_Trabajados"))
                Dias_Totales = 0
                For I = 1 To Total_Años
                'Calcular_Numero_Dias
                Dim Ye As Integer
'                Ye = Obtener_Valor_Ye(I)
                Ye = Calcular_Días(I, Tipo_Empleado_Aux)
                Dim Dias_Trabajados As Integer
                Numero_Días = Calcular_Numero_Dias(Dias_Restantes, Ye)
                Dias_Restantes = Dias_Restantes - 365
'                If Numero_Días > 0 Then
                Dias_Totales = Dias_Totales + Numero_Días
'                End If
                
'                Cadena_Datos = Cadena_Datos & Chr(9) & Numero_Días & Chr(9) & Dias_Restantes
                Cadena_Datos_Excel = Cadena_Datos_Excel & Chr(9) & Numero_Días & "|" & Chr(9) & Dias_Restantes & "|"
                Next I
                Dim Saldo As Double
                Saldo = Dias_Totales - Val(.rdoColumns("Total_Tomados"))
                Cadena_Datos = Cadena_Datos & Chr(9) & Dias_Totales & Chr(9) & .rdoColumns("Total_Tomados") & Chr(9) & Saldo
                Cadena_Datos_Excel = Cadena_Datos_Excel & Chr(9) & Dias_Totales & "|" & Chr(9) & .rdoColumns("Total_Tomados") & "|" & Chr(9) & Saldo & "|"
                Cadena_Datos_Impresión = Cadena_Datos_Impresión & "     " & Chr(9) & "     " & Dias_Totales & Chr(9) & "             " & .rdoColumns("Total_Tomados") & "      " & Chr(9) & Saldo
                Grid_Adm_Historico_Vacaciones.AddItem Cadena_Datos
                'Datos_Cadena_Vacaciones
                If Trim(Cadena_Columnas_Desglose_Vacaciones) <> "" Then
                Dim Cadena_Datos_Vacaciones As String
                Cadena_Datos_Vacaciones = Datos_Desglose_Vacaciones(.rdoColumns("Empleado_Id"), Cadena_Columnas_Desglose_Vacaciones)
                Cadena_Datos_Excel = Cadena_Datos_Excel & Cadena_Datos_Vacaciones
                End If
                 Print #1, Cadena_Datos_Impresión
                Print #2, Cadena_Datos_Excel
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Adm_Historico_Vacaciones.FixedRows = 1
            Grid_Adm_Historico_Vacaciones.ColWidth(0) = 1000     'No_Tarjeta
            Grid_Adm_Historico_Vacaciones.ColWidth(1) = 3000   'Nombre
            Grid_Adm_Historico_Vacaciones.ColWidth(2) = 1000  'Tipo_Empleado
            Grid_Adm_Historico_Vacaciones.ColWidth(3) = 0  'Empleado_ID
            Grid_Adm_Historico_Vacaciones.ColWidth(4) = 1000  'Dias_Totales
            Grid_Adm_Historico_Vacaciones.ColWidth(5) = 1000  'Tomados
            Grid_Adm_Historico_Vacaciones.ColWidth(6) = 1000  'Saldo
            .Close
        End If
        Call Finalizar_Reporte
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Adm_Empleados = Nothing
    

End Sub

Private Sub Btn_Buscar_Click()
'Genera_Reporte
Me.MousePointer = 11
Consulta_Historico_Vacaciones
Me.MousePointer = 0
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Calculos_Años
    'DESCRIPCIÓN:           Reimprime el grid con de acuerdo a las fechas
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Function Calculos_Años() As Integer
'Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
'    Dim Dia_Actual  As String
'    Dia_Actual = Now
'    Dim Total_Años  As Integer
'    Calculos_Años = 0
'    'Consulta los datos generales del usuario
''    Mi_SQL = "select TOP 1 Fecha_Ingreso, DATEDIFF (YEAR,  Fecha_Ingreso, '" & Format(Dia_Actual, "dd/mm/yyyy") & "') as Total_Años from Cat_Empleados"
'    Mi_SQL = "select Top 1 Fecha_Ingreso, " & Year(Dia_Actual) & "-YEAR(Fecha_Ingreso) AS Total_Años from Cat_Empleados"
'    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Ingreso ASC "
'    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'    With Rs_Consulta_Adm_Empleados
'        If Not .EOF Then
'            Calculos_Años = .rdoColumns("Total_Años")
'        End If
'    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Obtener_Valor_Ye
    'DESCRIPCIÓN:           Obtiene el valor de la BD de acuerdo al año
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Function Obtener_Valor_Ye(ByVal año As Integer) As Integer
Dim Rs_Consulta_Referencia As rdoResultset       'Informacion de los registros
    'Consulta los datos generales del usuario
    Mi_SQL = "select Top 1 Referencia_Id, Año, Valor from Ope_Referencias_Reporte_Vacaciones"
    Mi_SQL = Mi_SQL & " where Año = " & año
    Set Rs_Consulta_Referencia = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Referencia
        If Not .EOF Then
            Obtener_Valor_Ye = .rdoColumns("Valor")
        End If
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Calcular_Numero_Dias
    'DESCRIPCIÓN:           Obtiene el valor de la BD de acuerdo al año
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Calcular_Numero_Dias(ByVal Dias_Trabajados As Double, ByVal Ye As Integer) As Integer
If Dias_Trabajados - 365 > 0 Then
    Calcular_Numero_Dias = Ye
Else
    If (Dias_Trabajados * ((1 * Ye) / 365)) < 0 Then
        Calcular_Numero_Dias = 0
    Else
        Calcular_Numero_Dias = Dias_Trabajados * ((1 * Ye) / 365)
    End If
End If
End Function

Private Sub Btn_Exportar_Click()
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String

On Error GoTo HANDLER
    Cmd_Exportar.CancelError = True
    Cmd_Exportar.DialogTitle = "Seleccione el directorio"
    Cmd_Exportar.Flags = cdlOFNHideReadOnly
    Cmd_Exportar.Filter = "Archivos de Excel(*.xls)|*.xls"
    Cmd_Exportar.FilterIndex = 2
    Cmd_Exportar.FileName = Reporte & ".xls"
    Cmd_Exportar.ShowSave
    Ruta_Exportacion = Cmd_Exportar.FileName
    Nombre_Archivo = Cmd_Exportar.FileTitle
    If Cmd_Exportar.FileName <> "" And Nombre_Archivo <> "" Then
        Call Exportar_Excel_Bien(Ruta_Temporal & Reporte & "xls.txt", Ruta_Exportacion)
    End If
    'Display name of selected file
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
'    Lbl_Progreso_Exportacion.Caption = "Exportando ..."
'    Lbl_Progreso_Exportacion.Visible = True
'    Prbar_Exportacion.Visible = True
'    Prbar_Exportacion.Value = 0
'    Prbar_Exportacion.Min = 0
    'Nuevo objeto Excel
    Set obj_Excel = CreateObject("Excel.Application")
    With obj_Excel
        'Agrega un libro
        .Workbooks.Add
        ' Obtiene el número de líneas del Csv con la función split
        Lineas = Split(Contenido, vbCrLf)
'        Prbar_Exportacion.Max = UBound(Lineas) + 1
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
'            Prbar_Exportacion.Value = Prbar_Exportacion.Value + 1
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
'    Lbl_Progreso_Exportacion.Caption = "Guardando ..."
    ' Guarda el documento Xls
    obj_Excel.ActiveWorkbook.SaveAs _
        FileName:=Ruta, _
        Password:="", _
        WriteResPassword:="", _
        ReadOnlyRecommended:=False, _
        CreateBackup:=False
    'obj_Excel.ActiveWorkbook.Close False
'    Lbl_Progreso_Exportacion.Visible = False
'    Prbar_Exportacion.Visible = False
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
'    Lbl_Progreso_Exportacion.Visible = False
'    Prbar_Exportacion.Visible = False
End Sub



Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    
    Open Ruta_Temporal & "Reporte_Historico_Vacaciones.txt" For Output As #1
    Open Ruta_Temporal & Opcion & "xls.txt" For Output As #2 'Reporte a xls
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
Private Sub Finalizar_Reporte()
    Close #1, #2
End Sub

Private Sub Imprimir()
Dim linea As String 'Obtiene el texto a imprimir
Dim X As Printer
Dim contar_linea As Integer
Dim Foto_Empleado As New StdPicture
Dim No_Tarjeta As String
Dim Cont_Fila As Integer
Dim Cordenada_Y_Imagen  As Double
Dim Cont_Saltos As Integer
Dim Mi_SQL As String
'Dim Rs_Conssultar_Foto As rdoResultset
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
    
    Open Ruta_Temporal & "Reporte_Historico_Vacaciones.txt" For Input As #1
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

Private Sub Btn_Imprimir_Click()
Imprimir
End Sub

Private Sub Btn_Salir_Click()
Unload Me
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Calcular_Días
    'DESCRIPCIÓN:           Calcula los dias de vacaciones de acuerdo al año y al empleado
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    15 Marzo 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Calcular_Días(ByVal Años As String, ByVal Tipo_Empleado As String) As Integer

Dim Rs_Consulta_Dias_Vacaciones As rdoResultset       'Informacion de los registros
     Dim Diaasss As String
    Mi_SQL = "select * from Ope_Referencias_Reporte_Vacaciones "
    Set Rs_Consulta_Dias_Vacaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Dias_Vacaciones
        While Not .EOF
        If Años > 10 And (Val(Años) Mod 5 = 0) Then
            If Trim(.rdoColumns("Tipo_Empleado")) = Trim(Tipo_Empleado) Then
'                Dim Diaasss As String
                Diaasss = .rdoColumns("Añosmas10")
                Calcular_Días = Diaasss
            End If
        Else
            If Trim(.rdoColumns("Tipo_Empleado")) = Trim(Tipo_Empleado) Then
               
                Diaasss = .rdoColumns("Año_" & Años)
                Calcular_Días = Diaasss
            End If
        End If
        .MoveNext
        Wend
            .Close
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Dias_Vacaciones = Nothing


End Function


Function Columnas_Desglose_Vacaciones() As String
Columnas_Desglose_Vacaciones = ""
Dim Fecha_Va As Date
Dim Mes_Año As String
Dim Rs_Consulta_Dias_Vacaciones As rdoResultset       'Informacion de los registros
'    Mi_SQL = "select Fecha_Inicio, Fecha_Fin from Adm_Movimientos_Vacaciones order By Fecha_Inicio"
    Mi_SQL = "select * from Adm_Movimientos_Asistencias where (Tipo_Falta_ID = 4 or Tipo_Falta_ID = 10)"
    Mi_SQL = Mi_SQL & "and (Fecha_Inicio >= '" & Format(Dt_Rpt_Ingresos_De.Value, "mm/dd/yyyy") & "' and Fecha_Termino <= '" & Format(Dt_Rpt_Ingresos_A.Value, "mm/dd/yyyy") & "' ) order By Fecha_Inicio"
    
    Set Rs_Consulta_Dias_Vacaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Dias_Vacaciones
        If Not .EOF Then
            While Not .EOF
'                Mes_Año = Month(.rdoColumns("Fecha_Inicio")) & "-" & Year(.rdoColumns("Fecha_Inicio"))
                 Fecha_Va = Format(.rdoColumns("Fecha_Inicio"), "mm/dd/yyyy")
                    While Fecha_Va <= Format(.rdoColumns("Fecha_Termino"), "mm/dd/yyyy")
                        If Len(Trim(Month(Fecha_Va))) = 1 Then
                        Mes_Año = "0" & Month(Fecha_Va) & "-" & Year(Fecha_Va)
                        Else
                    Mes_Año = Month(Fecha_Va) & "-" & Year(Fecha_Va)
                    End If
                    If InStr(Columnas_Desglose_Vacaciones, Mes_Año) = 0 Then
                    Columnas_Desglose_Vacaciones = Columnas_Desglose_Vacaciones & Mes_Año & "|"
                    End If
                    Fecha_Va = DateAdd("d", 1, Fecha_Va)
                    Wend
            .MoveNext
            Wend
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Dias_Vacaciones = Nothing


End Function

Function Formato_Meses_Columnas_Vacaciones(ByVal Cadena) As String
Formato_Meses_Columnas_Vacaciones = Cadena
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "01-", "Enero ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "02-", "Febrero ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "03-", "Marzo ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "04-", "Abril ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "05-", "Mayo ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "06-", "Junio ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "07-", "Julio ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "08-", "Agosto ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "09-", "Septiembre ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "10-", "Octubre ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "11-", "Noviembre ")
Formato_Meses_Columnas_Vacaciones = Replace(Formato_Meses_Columnas_Vacaciones, "12-", "Diciembre ")
End Function

Function Datos_Desglose_Vacaciones(ByVal Empleado_ID, ByVal Cadena) As String
Dim Regresar As String
Dim Contador_Dias As Integer
Regresar = ""
Dim Cadenita() As String
Cadenita = Split(Cadena, "|")
Dim Fecha_Comparar As Date
For I = 0 To UBound(Cadenita) - 1
Dim String_Fecha As String
String_Fecha = Fecha_Inicial_Formato_Meses(Cadenita(I))
Fecha_Comparar = Format(String_Fecha, "mm/dd/yyyy")
Dim Rs_Consulta_Dias_Vacaciones As rdoResultset       'Informacion de los registros
'    Mi_SQL = "SELECT Empleado_Id, Fecha_Inicio ,Fecha_Fin "
'    Mi_SQL = Mi_SQL & " From Adm_Movimientos_Vacaciones "
'    Mi_SQL = Mi_SQL & " Where Empleado_ID = " & Empleado_ID
'    Mi_SQL = Mi_SQL & "AND ('" & Fecha_Comparar & "' BETWEEN Fecha_Inicio AND Fecha_Fin )"
'    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Inicio"
Mi_SQL = "SELECT Empleado_Id, Fecha_Inicio, Fecha_Termino  From Adm_Movimientos_Asistencias  "
Mi_SQL = Mi_SQL & " Where Empleado_ID = " & Empleado_ID
Mi_SQL = Mi_SQL & " AND (Month('" & Fecha_Comparar & "') = month(Fecha_Inicio) or  Month('" & Fecha_Comparar & "') = Month(Fecha_Termino))"
Mi_SQL = Mi_SQL & " AND (year('" & Fecha_Comparar & "') = year(Fecha_Inicio) or  year('" & Fecha_Comparar & "') = Year(Fecha_Termino))"
Mi_SQL = Mi_SQL & " AND (Tipo_Falta_ID = 4 OR Tipo_Falta_ID = 10) ORDER BY Fecha_Inicio"
    
    Set Rs_Consulta_Dias_Vacaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Dias_Vacaciones
        If Not .EOF Then
        Dim Fecha_Va As Date
        Contador_Dias = 0
            While Not .EOF
            Fecha_Va = Format(.rdoColumns("Fecha_Inicio"), "mm/dd/yyyy")
                While Fecha_Va <= Format(.rdoColumns("Fecha_Termino"), "mm/dd/yyyy")
                    If Month(Fecha_Comparar) = Month(Fecha_Va) And Year(Fecha_Comparar) = Year(Fecha_Va) Then
'                        If (Weekday(Fecha_Va) <> 1) Then
                        Contador_Dias = Contador_Dias + 1
'                        End If
                    End If
                    Fecha_Va = DateAdd("d", 1, Fecha_Va)
                Wend
            .MoveNext
            Wend
            .Close
            
        Regresar = Regresar & Contador_Dias & "|"
        Else
        Regresar = Regresar & "0|"
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Dias_Vacaciones = Nothing
Next I
Datos_Desglose_Vacaciones = Regresar
End Function


Function Fecha_Inicial_Formato_Meses(ByVal Cadena) As String
Fecha_Inicial_Formato_Meses = Cadena
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "01-", "01-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "02-", "02-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "03-", "03-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "04-", "04-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "05-", "05-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "06-", "06-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "07-", "07-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "08-", "08-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "09-", "09-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "10-", "10-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "11-", "11-1-")
Fecha_Inicial_Formato_Meses = Replace(Fecha_Inicial_Formato_Meses, "12-", "12-1-")
End Function


Private Sub Txt_No_Tarjeta_Historico_Vacaciones_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Departamento As rdoResultset
Dim No_Tarjeta As String
 If KeyAscii = 13 Then
        No_Tarjeta = Format(Txt_No_Tarjeta_Historico_Vacaciones.Text, "00000")
           Mi_SQL = "select Empleado_ID, ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') + ' ' "
           Mi_SQL = Mi_SQL & "+ ISNULL(Nombre, '') as Nombre From Cat_Empleados "
           Mi_SQL = Mi_SQL & "WHERE Estatus = 'A' "
           If Trim(No_Tarjeta) = "" Then
           Mi_SQL = Mi_SQL & "and No_Tarjeta like '%%' "
           Else
            Mi_SQL = Mi_SQL & " and No_Tarjeta = " & No_Tarjeta & " "
           End If
           
           Mi_SQL = Mi_SQL & "ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
'
            Set Rs_Empleados_Departamento = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            Cmb_Historico_Vacaciones_Empleados.Clear
            If Trim(No_Tarjeta) = "" Then
            Cmb_Historico_Vacaciones_Empleados.AddItem "TODOS"

            Cmb_Historico_Vacaciones_Empleados.ItemData(Cmb_Historico_Vacaciones_Empleados.NewIndex) = 0
        End If
            While Not Rs_Empleados_Departamento.EOF
                Cmb_Historico_Vacaciones_Empleados.AddItem Rs_Empleados_Departamento.rdoColumns("Nombre")
                Cmb_Historico_Vacaciones_Empleados.ItemData(Cmb_Historico_Vacaciones_Empleados.NewIndex) = Rs_Empleados_Departamento.rdoColumns("Empleado_Id")
                Rs_Empleados_Departamento.MoveNext
            Wend
            Rs_Empleados_Departamento.Close
            If Cmb_Historico_Vacaciones_Empleados.ListCount > 0 Then
                Cmb_Historico_Vacaciones_Empleados.ListIndex = 0
            End If
End If
End Sub
