VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Adm_Control_Calzado_Empleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de calzado"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleMode       =   0  'User
   ScaleWidth      =   12583.42
   Begin VB.PictureBox Pic_Ope_Programacion_Cursos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.CommandButton Btn_Responsiva 
         Caption         =   "Imprimir Responsiva"
         Height          =   555
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "A"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox Txt_No_Tarjeta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Txt_Nombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox Txt_Empleado_ID 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "C"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Guardar 
         Caption         =   "Guardar"
         Height          =   555
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "A"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Adm_Control_Calzado 
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
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Control_Calzado 
            Height          =   1920
            Left            =   75
            TabIndex        =   2
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   3387
            _Version        =   393216
            Rows            =   0
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin MSComCtl2.DTPicker Dt_Adm_Control_Calzado_Fecha 
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   1440
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108265473
         CurrentDate     =   42373
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha de entrega de calzado"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "No. Tarjeta"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Empleado ID"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Lbl_Notificaciones_Aniversarios 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTROL DE CALZADO"
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
         Left            =   1500
         TabIndex        =   6
         Top             =   15
         Width           =   4425
      End
   End
End
Attribute VB_Name = "Frm_Adm_Control_Calzado_Empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
Consulta_Empleados_Calzado ""
Dt_Adm_Control_Calzado_Fecha.Value = Now
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Empleados_Calzado
    'DESCRIPCIÓN:           Consulta los empleados pendientes a recibir calzado
    'PARÁMETROS :           Nombre: Indica el nombre del empleado
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           29 Febrero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Empleados_Calzado(Nombre As String)
Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
Dim Meses As Integer
Meses = Consultar_Meses_Calzado()
    Grid_Adm_Control_Calzado.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT No_Tarjeta, Nombre, Apellido_Paterno, Apellido_Materno, Empleado_ID, CONVERT(DATE, Fecha_Ultima_Entrega_Calzado) AS Fecha_Ultima_Entrega_Calzado"
    Mi_SQL = Mi_SQL & " from Cat_Empleados "
    Mi_SQL = Mi_SQL & " Where (Fecha_Ultima_Entrega_Calzado Is Null "
    Mi_SQL = Mi_SQL & " OR DATEADD (MONTH, " & Meses & ", Fecha_Ultima_Entrega_Calzado) <= GETDATE())"
    Mi_SQL = Mi_SQL & " AND Estatus = 'A' "
    Mi_SQL = Mi_SQL & " AND (No_Tarjeta LIKE '" & Nombre & "' OR Nombre LIKE '%" & Nombre & "%' OR Apellido_Paterno LIKE '%" & Nombre & "%' OR Apellido_Materno LIKE '%" & Nombre & "%' )"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Ultima_Entrega_Calzado"
    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Empleados
        If Not .EOF Then
            Grid_Adm_Control_Calzado.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "A. Paterno" & Chr(9) & "A. Materno" & Chr(9) & "Fecha" & Chr(9) & "Empleado_ID"
            While Not .EOF
                Grid_Adm_Control_Calzado.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Apellido_Paterno") & Chr(9) & .rdoColumns("Apellido_Materno") & Chr(9) & .rdoColumns("Fecha_Ultima_Entrega_Calzado") & Chr(9) & .rdoColumns("Empleado_ID")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Adm_Control_Calzado.FixedRows = 1
            Grid_Adm_Control_Calzado.ColWidth(0) = 800     'No_Tarjeta
            Grid_Adm_Control_Calzado.ColWidth(1) = 2000   'Nombre
            Grid_Adm_Control_Calzado.ColWidth(2) = 1500   'A paterno
            Grid_Adm_Control_Calzado.ColWidth(3) = 1500  'A. Materno
            Grid_Adm_Control_Calzado.ColWidth(4) = 1000  'Fecha
            Grid_Adm_Control_Calzado.ColWidth(5) = 0  'Empleado_ID
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Adm_Empleados = Nothing

End Sub


'*******************************************************************************
'NOMBRE_FUNCION:  Consultar_Meses_Calzado
'DESCRIPCION:     Consulta los parámetros para saber los meses para el cambio de calzado
'PARAMETROS :
'CREO       :     Ana Laura Huichapa Ramírez
'FECHA_CREO :     29-Febrero-2016
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Function Consultar_Meses_Calzado() As Integer
Dim Rs_Consulta_Cat_Parametros As rdoResultset

On Error GoTo HANDLER
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Parametros
        If Not .EOF Then
           
             If Not IsNull(.rdoColumns("Meses_Cambio_Calzado")) Then
            Consultar_Meses_Calzado = .rdoColumns("Meses_Cambio_Calzado")
            Else
             Consultar_Meses_Calzado = 0
            End If
            
        End If
    End With
    Set Rs_Consulta_Cat_Parametros = Nothing
Exit Function
HANDLER:
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Function

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el No_Tarjeta, Nombre, Apellido Paterno o Apellido Materno para buscar los empleados")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Empleados_Calzado Nombre
End Sub

Private Sub Btn_Enviar_Correos_Click()

End Sub

Private Sub Btn_Guardar_Click()
Modificar_Fecha_Entrega
Btn_Responsiva.Visible = True
End Sub

Private Sub Btn_Responsiva_Click()
Dim Hoora As Date
Hoora = Format$(Now, "d-mmmm-yy h:mm:ss")
Dim hora As String
hora = Replace(Hoora, " ", "")
hora = Replace(hora, ":", "_")
hora = Replace(hora, ".", "")
hora = Replace(hora, "/", "")

Crea_PDF "Rpt_Carta_Responsiva", "Entrega_Calzado_" & Txt_Nombre.Text & "_" & hora
End Sub

Private Sub Btn_Salir_Click()
Unload Me
End Sub

Private Sub Grid_Adm_Control_Calzado_Click()
Txt_Empleado_Id.Text = Grid_Adm_Control_Calzado.TextMatrix(Grid_Adm_Control_Calzado.RowSel, 5)
Txt_No_Tarjeta.Text = Grid_Adm_Control_Calzado.TextMatrix(Grid_Adm_Control_Calzado.RowSel, 0)
Txt_Nombre.Text = Grid_Adm_Control_Calzado.TextMatrix(Grid_Adm_Control_Calzado.RowSel, 1) & " " & Grid_Adm_Control_Calzado.TextMatrix(Grid_Adm_Control_Calzado.RowSel, 2) & " " & Grid_Adm_Control_Calzado.TextMatrix(Grid_Adm_Control_Calzado.RowSel, 3)
Btn_Responsiva.Visible = False
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Institucion
    'DESCRIPCIÓN:           Modifica el registro de la Institución
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    21 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Fecha_Entrega()
Dim Rs_Modificacion_Cat_Empleados As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Empleado_ID ='" & Trim(Txt_Empleado_Id) & "'"
    Set Rs_Modificacion_Cat_Empleados = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Empleados
        .Edit
            .rdoColumns("Fecha_Ultima_Entrega_Calzado") = Dt_Adm_Control_Calzado_Fecha.Value
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
        
    End With
    Set Rs_Modificacion_Cat_Empleados = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "Control de calzado actualizado para este empleado", vbInformation + vbOKOnly, Me.Caption
   Consulta_Empleados_Calzado ""
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Sub Crea_PDF(Reporte_Rpt As String, Nombre As String)
Dim crxApplication As New CRAXDDRT.Application
Dim crxReport As CRAXDDRT.Report
Dim crxDatabase As CRAXDDRT.Database
Dim crxDatabaseTables As CRAXDDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
Dim crxSections As CRAXDDRT.Sections
Dim crxSection As CRAXDDRT.Section
Dim crxSubreport As CRAXDDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Ruta_RPT As String
Dim Ruta_Salida
    
On Error GoTo HANDLER
    'Asigna el formato de la factura a la variable
    Ruta_RPT = App.Path & "\Reportes\" & Reporte_Rpt & ".rpt"
    Ruta_Salida = App.Path & "\Reportes_Cursos_Capacitaciones\" & Nombre & ".doc"

     Set crxReport = crxApplication.OpenReport(Ruta_RPT)
           
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo "SRG_Recursos_Humanos", Conectar_Ayudante.Base, Conectar_Ayudante.Usuario_Conexion, Conectar_Ayudante.Password
            End Select
        Next
    End With
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
    Dim Fecha As Date
    Dim parametro As String
        Select Case crParamDef.ParameterFieldName
        'Cursos_Tomados_Por_Empleado
            Case "Nombre_Empleado"
                If Trim(Txt_Nombre.Text) <> "" Then
                    crParamDef.AddCurrentValue (Txt_Nombre.Text)
                End If

            
        End Select
    Next
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
   crxReport.ExportOptions.DiskFileName = Ruta_Salida


    crxReport.ExportOptions.FormatType = crEFTWordForWindows
'crxReport.ExportOptions.FormatType = crEFTExcel97
    crxReport.ExportOptions.PDFExportAllPages = True
    'Oculta el progreso de la exportacion
    crxReport.DisplayProgressDialog = False
    'Genera la exportación del documento
    crxReport.Export (False)
    'Destruye el documento
    Set crxReport = Nothing
    ShellExecute Me.hwnd, "open", Ruta_Salida, "", "", 4
    Exit Sub
HANDLER:
    Printer.EndDoc
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub



