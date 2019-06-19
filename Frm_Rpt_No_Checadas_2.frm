VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Rpt_No_Checadas_2 
   Caption         =   "Reporte No Checadas"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   8160
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox Pic_Box_Fondo 
      BackColor       =   &H8000000E&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   4320
         TabIndex        =   13
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton Btn_Exportar_Excel 
         Caption         =   "Exportar Excel"
         Height          =   615
         Left            =   2520
         TabIndex        =   12
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Frame Fra_Grid_No_Checadas 
         BackColor       =   &H8000000E&
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
         Height          =   4335
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   7455
         Begin MSFlexGridLib.MSFlexGrid Grid_No_Checadas 
            Height          =   3855
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.ComboBox Cmb_Empleados 
         Height          =   315
         ItemData        =   "Frm_Rpt_No_Checadas_2.frx":0000
         Left            =   2040
         List            =   "Frm_Rpt_No_Checadas_2.frx":0002
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1560
         Width           =   5535
      End
      Begin VB.TextBox Txt_No_Tarjeta 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker Dt_Fecha_Fin 
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54657025
         CurrentDate     =   42508
      End
      Begin MSComCtl2.DTPicker Dt_Fecha_Inicio 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54657025
         CurrentDate     =   42508
      End
      Begin VB.CheckBox Chk_Rango_Fechas 
         BackColor       =   &H8000000E&
         Caption         =   "Rangos de Fechas"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Btn_PDF 
         Caption         =   "Exportar PDF"
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   6480
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog Cmd_Exportar 
         Left            =   240
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Empleado"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "No. Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "REPORTE NO CHECADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Frm_Rpt_No_Checadas_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicializa()
Me.Width = 8140
Me.Height = 8040
Cmb_Empleados.Clear
Cmb_Empleados.AddItem "TODOS"
Cmb_Empleados.ItemData(Cmb_Empleados.NewIndex) = 0

Mi_SQL = "select Empleado_ID, ISNULL(Nombre, '') + ' ' + ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') as Nombre"
Mi_SQL = Mi_SQL & " From Cat_Empleados where Estatus = 'A' ORDER bY Nombre, Apellido_Paterno, Apellido_Materno"
Set Rs_Empleados_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            While Not Rs_Empleados_Consulta.EOF
            DoEvents
                Cmb_Empleados.AddItem Rs_Empleados_Consulta.rdoColumns("Nombre")
                Cmb_Empleados.ItemData(Cmb_Empleados.NewIndex) = Rs_Empleados_Consulta.rdoColumns("Empleado_Id")
                Rs_Empleados_Consulta.MoveNext
            Wend
            Rs_Empleados_Consulta.Close
            If Cmb_Empleados.ListCount > 0 Then
                Cmb_Empleados.ListIndex = 0
            End If
End Sub

Private Sub Btn_Exportar_Excel_Click()
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

Private Sub Btn_PDF_Click()
Dim Nombre As String
Dim Nombre_RPT As String
Dim Hoora As Date
Hoora = Format$(Now, "d-mmmm-yy h:mm:ss")
Dim hora As String
hora = Replace(Hoora, " ", "")
hora = Replace(hora, ":", "_")
hora = Replace(hora, ".", "")
hora = Replace(hora, "/", "")
                Nombre_RPT = "Rpt_No_Checadas"
                Nombre = "Rpt_No_Checadas_" & hora
Crea_PDF Nombre_RPT, Nombre

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
    Ruta_Salida = App.Path & "\Reportes_Cursos_Capacitaciones\" & Nombre & ".pdf"

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
    Dim Fecha_enviar As String
    Dim parametro As Integer
    Dim parametro_enviar As String
        Select Case crParamDef.ParameterFieldName
'        Cursos_Tomados_Por_Empleado
            Case "No_Tarjeta_No_Checadas"
                If Cmb_Empleados.ListIndex = 0 Then
                    parametro = 0
                Else
                    Dim No_Tarjeta As String
                    No_Tarjeta = Obtener_No_Tarjeta(Format(Cmb_Empleados.ItemData(Cmb_Empleados.ListIndex), "00000"))
                    parametro = No_Tarjeta
                End If
                parametro_enviar = parametro
                 crParamDef.AddCurrentValue (parametro_enviar)

            Case "Fecha_Inicio_No_Checadas"
                If Chk_Rango_Fechas.Value = 1 Then
                   Fecha = Format(Dt_Fecha_Inicio.Value, "MM/dd/yyyy")
                Else
                    Fecha = Format("01/01/1990", "MM/dd/yyyy")
                End If
                Fecha_enviar = Fecha
                crParamDef.AddCurrentValue ("'" + Fecha_enviar + "'")

            Case "Fecha_Fin_No_Checadas"
                If Chk_Rango_Fechas.Value = 1 Then
                   Fecha = Format(Dt_Fecha_Fin.Value, "MM/dd/yyyy")
                Else
                   Fecha = Format("12/31/2100", "MM/dd/yyyy")
                End If
                Fecha_enviar = Fecha
                crParamDef.AddCurrentValue ("'" + Fecha_enviar + "'")
                
        End Select
    Next
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
   crxReport.ExportOptions.DiskFileName = Ruta_Salida

   

    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat
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


Private Sub Command2_Click()
Me.MousePointer = 11
Consulta_No_Checadas
Me.MousePointer = 0
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_No_Checadas
    'DESCRIPCIÓN:           Consulta los días que no han sido checados por el empleado
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           10 Mayo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Sub Consulta_No_Checadas()
Dim No_Tarjeta As String
If Cmb_Empleados.ListIndex > 0 Then
No_Tarjeta = Obtener_No_Tarjeta(Format(Cmb_Empleados.ItemData(Cmb_Empleados.ListIndex), "00000"))
End If
Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
    Grid_No_Checadas.Clear
    Grid_No_Checadas.Rows = 0
    Grid_No_Checadas.Cols = 5

    Dim Dia_Actual  As String
    Dia_Actual = Now
    
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT  CE.Empleado_ID, CE.Apellido_Paterno + ' ' + CE.Apellido_Materno + ' ' + CE.Nombre AS Empleado, CE.No_Tarjeta, CE.Tipo_Empleado, "
    Mi_SQL = Mi_SQL & " CD.Nombre AS Departamento, CD.Clave, CONVERT(date, Adm_Asistencias_Detalles.Fecha) AS Fecha "
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados AS CE INNER JOIN Cat_Departamentos AS CD ON CE.Departamento_ID = CD.Departamento_ID INNER JOIN"
    Mi_SQL = Mi_SQL & " Adm_Asistencias_Detalles ON CE.Empleado_ID = Adm_Asistencias_Detalles.Empleado_ID "
    Mi_SQL = Mi_SQL & " WHERE     (CE.Estatus = 'A') "
    If Chk_Rango_Fechas Then
    Mi_SQL = Mi_SQL & "AND  Adm_Asistencias_Detalles.Fecha between '" & Format(Dt_Fecha_Inicio, "MM/dd/yyyy") & "' and  '" & Format(Dt_Fecha_Fin, "MM/dd/yyyy") & "' "
    End If
    Mi_SQL = Mi_SQL & " AND (Adm_Asistencias_Detalles.Validada = 'N') "
    If Cmb_Empleados.ListIndex > 0 Then
    Mi_SQL = Mi_SQL & " and CE.No_Tarjeta = " & No_Tarjeta & " "
    End If
    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Empleados
        If Not .EOF Then
             
            Call Encabezado_Reporte("REPORTE NO CHECADAS", DateAdd("s", 1, Dt_Fecha_Inicio.Value), DateAdd("s", 1, Dt_Fecha_Fin.Value))
            
            Grid_No_Checadas.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "Tipo Empleado" & Chr(9) & "Departamento" & Chr(9)
            
            Print #2, "No Tarjeta|" & "Nombre|" & "Fecha|" & "Tipo_Empleado|" & "Departamento|"
            
            While Not .EOF
'                Cadena_Datos_Excel = Cadena_Datos_Excel & Chr(9) & Numero_Días & "|" & Chr(9) & Dias_Restantes & "|"
'                Dim hora As Date
'                hora = CDate(.rdoColumns("Hora"))
                Grid_No_Checadas.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Empleado") & Chr(9) & .rdoColumns("Fecha") & Chr(9) & .rdoColumns("Tipo_Empleado") & Chr(9) & .rdoColumns("Departamento") & Chr(9)
                
                Print #2, .rdoColumns("No_Tarjeta") & "|" & Chr(9) & .rdoColumns("Empleado") & "|" & Chr(9) & .rdoColumns("Fecha") & "|" & Chr(9) & .rdoColumns("Tipo_Empleado") & "|" & Chr(9) & .rdoColumns("Departamento") & "|"
            .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_No_Checadas.FixedRows = 1
            Grid_No_Checadas.ColWidth(0) = 1000     'No_Tarjeta
            Grid_No_Checadas.ColWidth(1) = 3500   'Nombre
            Grid_No_Checadas.ColWidth(2) = 1000  'Fecha
            Grid_No_Checadas.ColWidth(3) = 1200  'Tipo_Empleado
            Grid_No_Checadas.ColWidth(4) = 1200  'Departamento
        End If
        Call Finalizar_Reporte
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Adm_Empleados = Nothing
End Sub

Private Sub Encabezado_Reporte(Titulo As String, Optional Fecha_Inicial As Date, Optional Fecha_Termino As Date, Optional Solo_mes As Boolean)
    
    Open Ruta_Temporal & "Reporte_No_Checadas.txt" For Output As #1
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

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Txt_No_Tarjeta_KeyPress(KeyAscii As Integer)
Dim Rs_Empleados_Consulta As rdoResultset
Dim No_Tarjeta As String
 If KeyAscii = 13 Then
        No_Tarjeta = Format(Txt_No_Tarjeta.Text, "00000")
        Mi_SQL = "select Empleado_ID, ISNULL(Apellido_Paterno, '') + ' ' + ISNULL(Apellido_Materno, '') + ' ' "
        Mi_SQL = Mi_SQL & "+ ISNULL(Nombre, '') as Nombre From Cat_Empleados "
        Mi_SQL = Mi_SQL & "WHERE Estatus = 'A' "
        If Trim(No_Tarjeta) = "" Then
            Mi_SQL = Mi_SQL & "and No_Tarjeta like '%%' "
        Else
            Mi_SQL = Mi_SQL & " and No_Tarjeta = " & No_Tarjeta & " "
        End If
        Mi_SQL = Mi_SQL & "ORDER BY Nombre, Apellido_Paterno, Apellido_Materno"
        Set Rs_Empleados_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        Cmb_Empleados.Clear
        Cmb_Empleados.AddItem "TODOS"
        Cmb_Empleados.ItemData(Cmb_Empleados.NewIndex) = 0
        While Not Rs_Empleados_Consulta.EOF
            If Trim(Rs_Empleados_Consulta.rdoColumns("Nombre")) <> "" Then
                Cmb_Empleados.AddItem Rs_Empleados_Consulta.rdoColumns("Nombre")
                Cmb_Empleados.ItemData(Cmb_Empleados.NewIndex) = Rs_Empleados_Consulta.rdoColumns("Empleado_ID")
                Rs_Empleados_Consulta.MoveNext
            End If
        Wend
            Rs_Empleados_Consulta.Close
            If Cmb_Empleados.ListCount > 0 Then
                If No_Tarjeta = "" Then
                    Cmb_Empleados.ListIndex = 0
                Else
                    If Cmb_Empleados.ListCount = 2 Then
                        Cmb_Empleados.ListIndex = 1
                    Else
                        Cmb_Empleados.ListIndex = 0
                    End If
                End If
            End If
        End If

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


Private Function Obtener_No_Tarjeta(ByVal Empleado_ID As String)
Mi_SQL = "select Top 1 No_Tarjeta "
Mi_SQL = Mi_SQL & " From Cat_Empleados where Empleado_ID = '" & Empleado_ID & "'"
Set Rs_Empleados_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Empleados_Consulta.EOF Then
                Obtener_No_Tarjeta = Rs_Empleados_Consulta.rdoColumns("No_Tarjeta")
            End If
            Rs_Empleados_Consulta.Close
End Function
