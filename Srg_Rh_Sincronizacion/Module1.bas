Attribute VB_Name = "Module1"
Public Catalogo As String                   'Indicar que formulario se va a abrir
Public Er As rdoError                       'Especifica que tipo de error se esta cometiendo
Public Conexion_Base As New rdoConnection   'Se utiliza para la conexion a la base de datos
Public Usuario As String                    'Se utiliza para guardar el Usuario y es utilizada en diferentes procesos
Public Usuario_ID As String
Public Nombre_Usuario As String             'Obtiene el nombre del usuario
Public Rol_ID As String                     'Obtiene el rol que tiene asignado el usuario
Public Rol As String                        'Obtiene el rol que tiene asignado el usuario
Public Empleado_Supervisor_ID As String
Public Punto As Boolean                     'Se utiliza para v�lidar que sea solo un punto decimal
Public Ciclos As Integer                    'Se utiliza para v�lidar el tiempo de espera de la pantalla de presentaci�n
Public Conectar_Ayudante As Ayudante        'Es utilizada para ligar a la ayuda
Public Par_Fecha As String
Public Mi_SQL As String
Public Base_Datos As String
Public Dias_Credito As Integer
Public Server As String
Public Database As String
Public User_Password As String
Public User_Conexion As String
Public Movimiento_Factura As String
Public Abrir_Movimiento As Boolean
Public Empresa As String
Public RFC As String
Public Direccion As String
Public CP As String
Public Telefono As String
Public Ciudad_Edo As String
Public Saldo As Double
Public Fecha_Saldo As String
Public Cliente_Pendiente As Boolean
Public Cliente_Temporal As Boolean
Public Pedido_Rechazado As Boolean
Public Tipo_Consulta_Mapa As String         'Se utiliza para el manejo del mapa de ubicaciones
Public Encontro_Turno As Boolean            'Valida si en la consulta de turnos encuentra turno abierto
Public Ubicacion_Usuario As String          'Almacena la ubicaci�n del usuario
'Public Application As New CRAXDRT.Application
'Public Report As New CRAXDRT.Report
'***************PARAMETROS GENERALES********************
Public PG_Leyenda_Cheques As String                 '#  Almacena la leyenda por default de los cheques
Public PG_Formato_Cuenta As String                  '#  Almacena el formato de las cuentas
Public PG_Retencion_ISR As String                   '#  Almacena la Retencion del ISR
Public PG_Retencion_IVA As String                   '#  Almacena la Retencion del IVA
Public PG_Retencion_Fletes As String                '#  Almacena la Retencion de Fletes
Public PG_Impuesto_Cedular As String                '#  Almacena el Impuesto Cedular
Public PG_Comision_Completa As Double               '#  Almacena el porcentaje de la comision completa
Public PG_Comision_Promocion As Double              '#  Almacena el porcentaje de la comision promocion
Public PG_Cliente_Factura_Global As String          '#  Almacena el ID del cliente para las facturas globales
Public Serie_Factura_ZUMA As String                 'Serie de factura para la empresa ZUMA�
Public Antiguedad_Anticipos_Proveedores As Integer  'Antig�edad para validar el alta de anticipos de proveedores
'Seguridad
Public Dias_Caducidad_Contrase�as As Integer        'Almacena el par�metro de caducidad de las cuentas de usuario
Public Longitud_Minima_Password As Integer          'Almacena el par�metro de longitud m�nima de caracteres para el password
Public Intentos_Sesion_Fallidos As Integer          'Almacena el par�metro de cantidad de intentos fallidos permitidos
Public Historico_Password As Integer                'Almacena el par�metro de hist�rico de password que no pueden ser usados
Public Tipo_Validacion As Boolean                   'Almacena el tipo de utilidad que tenda la ventana de loguin
'***************FIN PARAMETROS GENERALES****************
'Obtiene el directorio ya sea windows o winnt
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Manejo de archivos
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

'******************************variables factura electronica***************
Public Ruta_Certificado As String           'Almacena la ruta del certificado
Public Ruta_Llave_Privada As String         'Almacena la ruta de la llave privada
Public Password_Llave As String             'Almacena el password de la llave privada
Public Mensaje_Error As String
'******************************************************************************
'Recuros Humanos
Public Edad_Minima_Contratacion As Integer
Public Horas_Dobles As Integer                      'Parametro Horas
Public Horas_Triples As Integer
Public Dias_Falta As Integer
Public Periodo_Retardos_Dias As Integer
Public Tipo_Nomina  As String
Public Minutos_Tolerancia As Integer
Public Minutos_Comida As Integer
Public Dias_Aviso_Contrato_Eventual As Integer      'Dias antes de vencimiento de contrato
Public PDF_Horas_Dobles As String
Public PDF_Horas_Triples As String
Public PDF_Enfermedad_General As String
Public PDF_Maternidad As String
Public PDF_Riesgo_Trabajo As String
Public PDF_Vacaciones As String
Public PDF_Alumbramiento As String
Public PDF_Defuncion As String
Public PDF_Matrimonio As String
Public PDF_Falta_Justificada As String
Public PDF_Falta_InJustificada As String
Public PDF_Permiso_Temporal As String
Public PDF_Retardo As String
Public PDF_Ayuda_Transporte As String
Public PDF_Ayuda_Comida As String
Public PDF_Dia_Doble As String
Public PDF_Permiso_CG As String
Public PDF_Permiso_SG As String
'**************Parametros para envio de correos
Public Email_Sistema As String
Public Hora_Importacion As String
Public Hora_Importacion_Dia As String
Public Email_validacion As String
Public Email_Administrador As String
Public Email_Notificacion As String
Public Proceso As String
Public Servidor_SMTP As String
Public Puerto_SMTP As Integer
'************ Formato vista previa
Public Alto_Carta_Cms, Ancho_Carta_Cms As Double
Public Alto_MCarta_Cms, Ancho_MCarta_Cms As Double
'*********************Redimensionar imagenes
Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4
Public Const CF_BITMAP = 2
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'********************************
'Ruta temporal de windows
Public Ruta_Temporal As String                      'Especifica la carpeta de temporales de windows
'Obtiene el directorio de archivos temporales
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Obtiene el nombre de la maquina
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Api para saber si una carpeta existe
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
'Abrir archivos
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3
'Funciones para esperar a que termine un proceso de otro programa
Public Const INFINITE = &HFFFF
Public Const SYNCHRONIZE = &H100000
Public Declare Sub WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDA As Long, ByVal bIH As Integer, ByVal dwPID As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32.dll" (ByVal hObject As Long)

'*******************************************************************************
'NOMBRE_FUNCION: Actualiza_Turnos_Programacion
'DESCRIPCION: Actualiza los turnos con los que fueron programados para los empleados
'PARAMETROS :
'CREO       : Sergio Ulises Dur�n Hern�ndez
'FECHA_CREO : 16-Abril-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Actualiza_Turnos_Programacion()
Dim Rs_Consulta_Cambio_Turno As rdoResultset
Dim Rs_Actualiza_Turno_Empleado As rdoResultset

On Error GoTo ErrorHandler
    MDIFrm_Apl_Principal.MousePointer = 11
    'Consulta los empleados que se les har� el cambio de turno a la fecha
    Mi_SQL = "SELECT * FROM Adm_Cambios_Turnos"
    Mi_SQL = Mi_SQL & " WHERE Fecha_Cambio='" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Estatus='PENDIENTE'"
    Set Rs_Consulta_Cambio_Turno = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    While Not Rs_Consulta_Cambio_Turno.EOF
        'Consulta el empleado que se le cambia el turno
        Mi_SQL = "SELECT Empleado_ID,Turno_ID,Usuario_Modifico,Fecha_Modifico FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Rs_Consulta_Cambio_Turno.rdoColumns("Empleado_ID") & "'"
        Set Rs_Actualiza_Turno_Empleado = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        If Not Rs_Actualiza_Turno_Empleado.EOF Then
            Rs_Actualiza_Turno_Empleado.Edit
                Rs_Actualiza_Turno_Empleado.rdoColumns("Turno_ID") = Rs_Consulta_Cambio_Turno.rdoColumns("Turno_Nuevo_ID")
                Rs_Actualiza_Turno_Empleado.rdoColumns("Usuario_Modifico") = Nombre_Usuario
                Rs_Actualiza_Turno_Empleado.rdoColumns("Fecha_Modifico") = Now
            Rs_Actualiza_Turno_Empleado.Update
        End If
        Rs_Actualiza_Turno_Empleado.Close
        Rs_Consulta_Cambio_Turno.Edit
            Rs_Consulta_Cambio_Turno.rdoColumns("Estatus") = "CAMBIADO"
        Rs_Consulta_Cambio_Turno.Update
        Rs_Consulta_Cambio_Turno.MoveNext
    Wend
    Rs_Consulta_Cambio_Turno.Close
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
ErrorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Public Sub SHCopyFile(ByVal from_file As String, ByVal to_file As String)
Dim sh_op As SHFILEOPSTRUCT
    With sh_op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = from_file & vbNullChar & vbNullChar
        .pTo = to_file & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation sh_op
End Sub

Public Function Mensaje(ByVal cadena_mensaje As String, Optional ByVal Tipo As Integer = 1, Optional Titulo As String = "") As Integer
Dim Titulo_Mensaje  As String   '#  Almacena el titulo del msj
    
    Texto = "!!! Atenci�n !!!" & vbCrLf & vbCrLf & cadena_mensaje
    Titulo_Mensaje = Empresa
    If Titulo <> "" Then Titulo_Mensaje = Titulo
    Select Case Tipo
        Case 1: MsgBox Texto, vbInformation, Titulo_Mensaje
        Case 2: MsgBox Texto, vbCritical, Titulo_Mensaje
        Case 3: Mensaje = MsgBox(cadena_mensaje, vbQuestion + vbYesNo, "Confirmacion de Proceso")
    End Select
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCI�N: Alinea_Derecha
    'DESCRIPCI�N: Alinea_Derecha los n�mero a la izquierda del documento
    'PAR�METROS:
    '             1. Numero:
    '             2. Longitud:
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACI�N
'*******************************************************************************
'
Public Function Alinea_Derecha(Numero As String, Longitud As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim Caracteres_Ciclo As Integer     'Cuenta el numero de caracteres de la cadena

    Nuevo = Numero
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For Caracteres_Ciclo = 1 To Longitud - Len(Numero)
        Nuevo = " " & Nuevo
    Next Caracteres_Ciclo
    Alinea_Derecha_Derecha = Nuevo
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Parametros_Generales
'DESCRIPCION: Consulta los parametros generales del sistema
'CREO       :
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Consulta_Parametros_Generales()
Dim Mi_SQL As String
Dim Rs_Consulta_Cat_Parametros As rdoResultset  '#  consulta el catalogo de parametros
    
    'Limpio los parametros
    PG_Leyenda_Cheques = ""
    PG_Formato_Cuenta = ""
    PG_Impuesto_Cedular = ""
    
    PG_Retencion_IVA = ""
    PG_Retencion_ISR = ""
    PG_Cliente_Factura_Global = ""
    Serie_Factura_ZUMA = ""
    Antiguedad_Anticipos_Proveedores = 0
    PG_Comision_Completa = 0
    PG_Comision_Promocion = 0
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Parametros
        'Seguridad
        If Not IsNull(.rdoColumns("Dias_Caducidad_Contrase�a")) Then Dias_Caducidad_Contrase�as = .rdoColumns("Dias_Caducidad_Contrase�a")
        If Not IsNull(.rdoColumns("Longitud_Minima_Password")) Then Longitud_Minima_Password = .rdoColumns("Longitud_Minima_Password")
        If Not IsNull(.rdoColumns("Intentos_Sesion_Fallidos")) Then Intentos_Sesion_Fallidos = .rdoColumns("Intentos_Sesion_Fallidos")
        If Not IsNull(.rdoColumns("Historico_Password")) Then Historico_Password = .rdoColumns("Historico_Password")
    End With
    Rs_Consulta_Cat_Parametros.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCI�N: Actualiza saldo banco
    'DESCRIPCI�N: Actualiza el saldo del banco del cual se esta afectando el movimiento
    'PAR�METROS:
    '               1. Banco Movimiento : es el nombre del banco que se esta afectando en el movimiento
    '               2. Fecha movimiento :Se envia la fecha del movimiento que se eta afectando
    'CREO: Joel Romero
    'FECHA_CREO:
    'MODIFICO:     Jorge Razo
    'FECHA_MODIFICO : 10-Agosto-2006
    'CAUSA_MODIFICACI�N :Estandarizacion de codigo
'*******************************************************************************
Public Sub Actualiza_Saldo(Banco_Movimiento As String, Fecha_Movimiento As Date)
Dim Rs_Movimiento As rdoResultset       'Manejador de datos de los movimientos
Dim Rs_Modifica_Cat_Bancos As rdoResultset  '#  Modifica el saldo del banco
Dim Rs_Consulta_Fecha As rdoResultset   'Manejador de datos de consulta de fecha
Dim Rs_Saldo As rdoResultset            'Manejador de consulta de saldo
Dim Fecha As String                     'Almacena la fecha del movimiento en curso
Dim Fecha_Saldo As String               'Almacena la ultima fecha del saldo
Dim Banco As String                     'Almacena el Id del banco que se afectara con la actualizacion
Dim Saldo As Double                     'Almacena el saldo del banco en cada movimiento afectado
Dim Consecutivo As Double

On Error GoTo MuestraError
    'Obtiene el ultimo movimiento de acuerdo a la fecha
    Fecha = Format(DateAdd("d", -1, Format(Fecha_Movimiento, "MM/dd/yyyy")), "MM/dd/yyyy")
    Banco = Banco_Movimiento
    Bandera = 0
    Saldo = 0
    'Obtiene la ultima fecha de los movimientos registrados del banco
    'hasta la fecha que se envia en el parametro fecha
    Mi_SQL = "SELECT MAX(Fecha)"
    Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Fecha<='" & Fecha & "'"
    Mi_SQL = Mi_SQL & " AND Banco_ID='" & Banco & "'"
    Mi_SQL = Mi_SQL & " AND Estatus='A'"
    Set Rs_Consulta_Fecha = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not IsNull(Rs_Consulta_Fecha(0)) Then
        Fecha_Saldo = Format(Rs_Consulta_Fecha(0), "MM/dd/yyyy")
        'Consulta el saldo de la fecha obtenida en la consulta anterior
        Mi_SQL = "SELECT Cantidad"
        Mi_SQL = Mi_SQL & " FROM Adm_Movimientos"
        Mi_SQL = Mi_SQL & " WHERE Fecha='" & Fecha_Saldo & "'"
        Mi_SQL = Mi_SQL & " AND Banco_ID='" & Banco & "'"
        Mi_SQL = Mi_SQL & " AND Estatus='A'"
        Mi_SQL = Mi_SQL & " ORDER BY Fecha ASC,No_Movimiento"
        Set Rs_Saldo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        If Not Rs_Saldo.EOF And Not IsNull(Rs_Saldo.rdoColumns("Cantidad")) Then
            Saldo = Rs_Saldo.rdoColumns("Cantidad")
        End If
        Rs_Saldo.Close
    Else
        Saldo = 0
        Bandera = 1
        Fecha_Saldo = Fecha
    End If
    Rs_Consulta_Fecha.Close
    'Actualiza saldos
    Mi_SQL = "SELECT * FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Fecha>='" & Format(Fecha_Saldo, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Banco_ID='" & Banco & "'"
    Mi_SQL = Mi_SQL & " AND (Tipo='I' OR Tipo='E')"
    Mi_SQL = Mi_SQL & " AND Estatus='A'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Movimiento"
    Set Rs_Movimiento = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Saldo = 0 And Bandera = 1 And Not Rs_Movimiento.EOF Then
        Saldo = Rs_Movimiento!Cantidad
        If Rs_Movimiento!Tipo = "E" Then
            Saldo = Saldo * (-1)
        End If
    End If
    While Not Rs_Movimiento.EOF
        Rs_Movimiento.Edit
            Rs_Movimiento!Cantidad = Saldo
        Rs_Movimiento.Update
        Rs_Movimiento.MoveNext
        If Not Rs_Movimiento.EOF Then
            If Rs_Movimiento!Tipo = "E" Then
                Saldo = Saldo - Rs_Movimiento!Cantidad
            Else
                Saldo = Saldo + Rs_Movimiento!Cantidad
            End If
        End If
    Wend
    Rs_Movimiento.Close
    'Consulta el saldo del banco
    Mi_SQL = "SELECT Banco_ID,Saldo,Fecha_Modifico FROM Cat_Bancos"
    Mi_SQL = Mi_SQL & " WHERE Banco_ID='" & Banco & "'"
    Set Rs_Modifica_Cat_Bancos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Modifica_Cat_Bancos.EOF Then
        Rs_Modifica_Cat_Bancos.Edit
            Rs_Modifica_Cat_Bancos.rdoColumns("Saldo") = Saldo
            Rs_Modifica_Cat_Bancos.rdoColumns("Fecha_Modifico") = Format(Now, "MM/dd/yyyy")
        Rs_Modifica_Cat_Bancos.Update
    End If
    Rs_Modifica_Cat_Bancos.Close
    Exit Sub
MuestraError:
    MsgBox Err.Description
End Sub



Public Function Calcula_Edad(Fecha_Nacimiento As Date, Optional ByRef Edad_Anios As Double) As String
Dim Calculo As Double
Dim anios As Double
    Calcula_Edad = ""
    Calculo = 0
    'Calcula los a�os cumplidos
    Edad_Anios = 0
    Calculo = DateDiff("M", Fecha_Nacimiento, Now)
    anios = Fix(Calculo / 12)
    Edad_Anios = anios
    Meses = Calculo Mod 12
    Calcula_Edad = anios & " a�os" & " y " & Meses & " mes(es)"
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Obtiene_Ruta_Temporal
'DESCRIPCION: Regresa el nombre de la carpeta de archivos temporales
'PARAMETROS :
'CREO       : Ya�ez Rodriguez Diego Neftali
'FECHA_CREO : 21-Mayo-2008
'MODIFICO   : Sergio Ulises Dur�n Hern�ndez
'FECHA_MODIFICO: 10-Mayo-2012
'CAUSA_MODIFICO: Se agreg� otra funci�n para obtener la ruta de mis documentos
'*******************************************************************************
Public Function Obtiene_Ruta_Temporal() As String
Dim Buffer As String, Size As Long
    
'    Const MAX_PATH = 260
'    'Inicializamos la cadena donde se cargar� la ruta
'    Buffer = String(MAX_PATH, 0)
'    'Recuperamos la trayectoria
'    Size = GetTempPath(Len(Buffer) - 1, Buffer)
'    If Size <> 0 Then
'        Obtiene_Ruta_Temporal = Left(Buffer, Size)
'    Else
'        Obtiene_Ruta_Temporal = "C:\"
'    End If
'    If Mid(Obtiene_Ruta_Temporal, Len(Obtiene_Ruta_Temporal), 1) = "\" Then
'        Obtiene_Ruta_Temporal = Mid(Obtiene_Ruta_Temporal, 1, Len(Obtiene_Ruta_Temporal) - 1)
'    End If
    
Dim Wscript As Object   'Variable para usar WSH
    'Creamos la referencia para usar Windows Scripting Host
    Set Wscript = CreateObject("WScript.Shell")
    Obtiene_Ruta_Temporal = Wscript.SpecialFolders("MyDocuments")
    If Not Wscript Is Nothing Then
       Set Wscript = Nothing
    End If
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Exportar_Excel
'DESCRIPCION: Genera el reporte en archivo de excel
'PARAMETROS : Ruta- Ruta donde se guardara el archivo
'             Nombre_Archivo- Nombre del archivo
'CREO       : Ya�ez Rodriguez Diego Neftali
'FECHA_CREO : 22-Abril-2008
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
'Public Sub Exportar_Excel(Archivo_Exportar As String, Ruta As String, Prg_Exportacion As ProgressBar, Lbl_Informacion As Label, Frm As Form)
'Dim obj_Excel As Object
'Dim Fila As Integer, Columna As Integer
'Dim Contenido As String, Lineas As Variant
'Dim Datos As Variant, MC As Integer
'Dim Encabezado As Boolean
'Dim Fila_Encabezado As Integer
'
'On Error GoTo HANDLER
'    MDIFrm_Apl_Principal.MousePointer = 11
'    'Lee el contenido del reporte
'    Open Archivo_Exportar For Input As #1
'    Contenido = Input$(LOF(1), #1)
'    Close
'    Lbl_Informacion.Caption = "Exportando ..."
'    Lbl_Informacion.Visible = True
'    Prg_Exportacion.Visible = True
'    Prg_Exportacion.Value = 0
'    Prg_Exportacion.Min = 0
'    'Nuevo objeto Excel
'    Set obj_Excel = CreateObject("Excel.Application")
'    With obj_Excel
'        'Agrega un libro
'        .Workbooks.Add
'        ' Obtiene el n�mero de l�neas del Csv con la funci�n split
'        Lineas = Split(Contenido, vbCrLf)
'        Prg_Exportacion.Max = UBound(Lineas) + 1
'        For Fila = 0 To UBound(Lineas)
'            'Encabezado = False
'            'Separa los datos de la linea
'            Datos = Split(Lineas(Fila), "|")
'            'Recorre los datos de esta fila que corresponden a cada campo
'            For Columna = 0 To UBound(Datos)
'                ' Agrega el dato a la celda de la hoja activa
'                .ActiveSheet.Cells(Fila + 1, Columna + 1) = Datos(Columna)
'                If Trim(Mid(Datos(Columna), 1, 1)) = "." Then
'                    Encabezado = True
'                    Fila_Encabezado = Fila
'                End If
'                If Encabezado Then
'                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Borders.LineStyle = 1
'                    '.ActiveSheet.Cells(Fila + 1, Columna + 1).Borders.Weight = 0
'                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Borders.Color = RGB(0, 0, 0)
'                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Font.FontStyle = "Bold"
'                    .ActiveSheet.Cells(Fila_Encabezado + 1, Columna + 1).Font.Size = 12
'                End If
'            Next
'            If MC < Columna Then
'               MC = Columna
'            End If
'            Prg_Exportacion.Value = Prg_Exportacion.Value + 1
'        Next
'        'Selecciona toda la hoja
'        .ActiveSheet.UsedRange.Select
'        'Autoajusta las columnas
'        '.Selection.Columns.AutoFit
'        'Selecciona el encabezado
'    End With
'    ' Aplica atributos a la fuente a la selecci�n anterior ( los encabezados )
'    With obj_Excel.Selection.Font
'        '.Name = "Verdana"
'        '.FontStyle = "Bold"
'        '.Size = 14
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        '.Underline = xlUnderlineStyleNone
'    End With
'    'se ajusta el documento a una hoja
'    'With obj_Excel.Selection
'    Lbl_Informacion.Caption = "Guardando ..."
'    ' Guarda el documento Xls
'    obj_Excel.ActiveWorkbook.SaveAs _
'        FileName:=Ruta, _
'        Password:="", _
'        WriteResPassword:="", _
'        ReadOnlyRecommended:=False, _
'        CreateBackup:=False
'    'obj_Excel.ActiveWorkbook.Close False
'    Lbl_Informacion.Visible = False
'    Prg_Exportacion.Visible = False
'    MDIFrm_Apl_Principal.MousePointer = 0
'     'Cierra el archivo y elimina la variable
'     If MsgBox("�Desea abrir el archivo?", vbQuestion + vbYesNo, Frm.Caption) = vbYes Then
'        obj_Excel.Visible = True
'     Else
'        obj_Excel.Quit
'        MsgBox "Reportes exportado", vbInformation + vbOKOnly, Frm.Caption
'    End If
'    'obj_Excel.Quit
'    Set obj_Excel = Nothing
'    Exit Sub
'Exit Sub
' ' Error
'HANDLER:
'    MDIFrm_Apl_Principal.MousePointer = 0
'    MsgBox Err.Description
'    On Error Resume Next
'    If Not obj_Excel Is Nothing Then
'        obj_Excel.Quit
'        Set obj_Excel = Nothing
'    End If
'    Lbl_Informacion.Visible = False
'    Prg_Exportacion.Visible = False
'End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCI�N: Valida_Importacion_Realizada
    'DESCRIPCI�N:          Valida si ya se ha realizado la importacion de las asistencias
    'PAR�METROS:           Fecha: Fecha a verificar
    '                      Nombre_Archivo: Nombre del archivo
    'CREO:                 Ya�ez Rodriguez Diego Neftali
    'FECHA_CREO:           22-Abril-2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACI�N
'*******************************************************************************
Public Function Valida_Importacion_Realizada(Fecha As Date) As Boolean
Dim Rs_Consulta_Bitacora As rdoResultset
Dim Mi_SQL_Bitacora As String

    Valida_Importacion_Realizada = False
    Mi_SQL_Bitacora = "SELECT * FROM Adm_Bitacora_Importacion"
    Mi_SQL_Bitacora = Mi_SQL_Bitacora & " WHERE Fecha = " & Par_Fecha & Format(Fecha, "MM/dd/yyyy") & Par_Fecha
    Set Rs_Consulta_Bitacora = Conectar_Ayudante.Recordset_Consultar(Mi_SQL_Bitacora)
    If Not Rs_Consulta_Bitacora.EOF Then
        Valida_Importacion_Realizada = True
    End If
    Set Rs_Consulta_Bitacora = Nothing
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCI�N: Obtiene_Periodo_Quincena
    'DESCRIPCI�N:          Obtiene el no de periodo de la fecha ingresada
    'PAR�METROS:           Fecha: Fecha a evaluar para el calculo del periodo
    'CREO:                 Ya�ez Rodriguez Diego Neftali
    'FECHA_CREO:           25 Junio 2009
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACI�N
'*******************************************************************************
Public Function Obtiene_Periodo_Quincena(Fecha As Date) As Integer
Dim Dia As Integer
Dim Mes As Integer
Dim Quincena As Integer

    Mes = Month(Fecha)
    Obtiene_Periodo_Quincena = Mes * 2
    Dia = Day(Fecha)
    If Dia <= 15 Then Quincena = Obtiene_Periodo_Quincena - 1
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCI�N: Crea_PDF_Empleado_Expediente
'DESCRIPCI�N:           Crea el archivo PDF de la solicitud de cotizaciones
'PAR�METROS:
'                       1. Ruta_Directorio, ruta donde se almacenara el pdf de la solicitud
'                       2. Nombre_Archivo_Solicitud, es el nombre que se dara al archvio
'CREO:                  Ya�ez Rodriguez Diego Neftali
'FECHA_CREO:
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACI�N
'*******************************************************************************
Public Function Crea_PDF_Empleado_Expediente(Ruta_Directorio As String, Nombre_Archivo As String, Empleado_ID As String, Ruta_Imagen As String) As Boolean
Dim crxApplication As New CRAXDRT.Application
Dim crxReport As CRAXDRT.Report
Dim crxDatabase As CRAXDRT.Database
Dim crxDatabaseTables As CRAXDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crxSubreport As CRAXDRT.Report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Ruta_Aplicacion As String

On Error GoTo HANDLER
    Crea_PDF_Empleado_Expediente = False
    'Elimina el archivo si llegara a existir
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Directorio & "\" & Nombre_Archivo & ".pdf", "ARCHIVO") = True Then
        Kill Ruta_Directorio & "\" & Nombre_Archivo & ".pdf"
    End If
    Ruta_Aplicacion = App.Path
    If Mid(Ruta_Aplicacion, Len(Ruta_Aplicacion), 1) = "\" Then
        Ruta_Aplicacion = Mid(Ruta_Aplicacion, 1, Len(Ruta_Aplicacion) - 1)
    End If
    Set crxReport = crxApplication.OpenReport(Ruta_Aplicacion & "\Reportes\Rpt_Empleado_Perfil.rpt")
    Identificador = "despues de abrir facturas crystal"
    
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                    Identificador = Identificador & " " & "Antes de ODBC"
                    .Database.Tables(Cuenta_Tablas).SetLogOnInfo Database, Database, User_Conexion, User_Password
                    Identificador = Identificador & " " & "DESPUES de ODBC"
            End Select
        Next
    End With
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Empleado_ID"
                crParamDef.AddCurrentValue (Empleado_ID)
            Case "Ruta_Imagen"
                crParamDef.AddCurrentValue (Ruta_Imagen)
        End Select
    Next
    
    Frm_Ver_Reportes.Crv_Reporte.DisplayBorder = False
    Frm_Ver_Reportes.Crv_Reporte.DisplayTabs = False
    Frm_Ver_Reportes.Crv_Reporte.EnableDrillDown = False
    Frm_Ver_Reportes.Crv_Reporte.EnableRefreshButton = False
    Frm_Ver_Reportes.Crv_Reporte.ReportSource = crxReport
    Frm_Ver_Reportes.Crv_Reporte.ViewReport
    Frm_Ver_Reportes.Crv_Reporte.Zoom 100
    
'    'Asigna los datos de exportaci�n
'    crxReport.ExportOptions.DestinationType = crEDTDiskFile
'
'    'MsgBox Ruta_Directorio & "\" & Nombre_Archivo & ".pdf"
'    crxReport.ExportOptions.DiskFileName = Ruta_Directorio & "\" & Nombre_Archivo & ".pdf"
'    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat
'    crxReport.ExportOptions.PDFExportAllPages = True
'    'Oculta el progreso de la exportacion
'    crxReport.DisplayProgressDialog = False
'
'    'Genera la exportaci�n del documento
'    crxReport.Export (False)
'    'crxReport.PrintOut
'    'Destruye el documento
'    Set crxReport = Nothing
    Crea_PDF_Empleado_Expediente = True
Exit Function
HANDLER:
    Crea_PDF_OV = False
    MsgBox Err.Description
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Parametros
'DESCRIPCION: Consulta los par�metros que tiene el sistema asignados
'PARAMETROS :
'CREO       : Yazmin Delgado G�mez
'FECHA_CREO : 15-Octubre-2007
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Consulta_Parametros()
Dim Rs_Consulta_Cat_Parametros As rdoResultset 'Consulta los par�metros que tiene asignado el sistema

    '#  Dimensiones de la hoja Carta
    Alto_Carta_Cms = 27.94
    Ancho_Carta_Cms = 21.59
    '#  Dimensiones media carta
    Alto_MCarta_Cms = 15.59
    Ancho_MCarta_Cms = 21.59
    'Consulta los par�metros del sistema
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Parametros.EOF Then
        With Rs_Consulta_Cat_Parametros
            Edad_Minima_Contratacion = .rdoColumns("Edad_Minima_Contratacion")
            Horas_Dobles = .rdoColumns("Horas_Dobles")
            Horas_Triples = .rdoColumns("Horas_Triples")
            Dias_Falta = .rdoColumns("Dias_Falta")
            Periodo_Retardos_Dias = .rdoColumns("Periodo_Retardos_Dias")
            'Tipo_Nomina = .rdoColumns("Tipo_Nomina")
            Minutos_Tolerancia = .rdoColumns("Minutos_Tolerancia")
            Email_Sistema = .rdoColumns("Email_Sistema")
            'Email_validacion = .rdoColumns("Email_validaci�n")
            Email_Administrador = .rdoColumns("Email_Administrador")
            Email_Notificacion = .rdoColumns("Email_Notificacion")
            Hora_Importacion = Format(.rdoColumns("Hora_Importacion"), "HH:mm")
            Hora_Importacion_Dia = Format(.rdoColumns("Hora_Importacion_Dia"), "HH:mm")
            Servidor_SMTP = .rdoColumns("Servidor_SMTP")
            Puerto_SMTP = Val(.rdoColumns("Puerto_SMTP"))
            PDF_Enfermedad_General = Trim(.rdoColumns("PDF_Enfermedad_General"))
            PDF_Maternidad = Trim(.rdoColumns("PDF_Maternidad"))
            PDF_Riesgo_Trabajo = Trim(.rdoColumns("PDF_Riesgo_Trabajo"))
            PDF_Vacaciones = Trim(.rdoColumns("PDF_Vacaciones"))
            PDF_Alumbramiento = Trim(.rdoColumns("PDF_Alumbramiento"))
            PDF_Defuncion = Trim(.rdoColumns("PDF_Defuncion"))
            PDF_Matrimonio = Trim(.rdoColumns("PDF_Matrimonio"))
            PDF_Falta_Justificada = Trim(.rdoColumns("PDF_Falta_Justificada"))
            PDF_Permiso_Temporal = Trim(.rdoColumns("PDF_Permiso_Temporal"))
            PDF_Horas_Dobles = Trim(.rdoColumns("PDF_Horas_Dobles"))
            PDF_Horas_Triples = Trim(.rdoColumns("PDF_Horas_Triples"))
            PDF_Falta_InJustificada = Trim(.rdoColumns("PDF_Falta_InJustificada"))
            PDF_Permiso_CG = Trim(.rdoColumns("PDF_Permiso_Goce"))
            PDF_Permiso_SG = Trim(.rdoColumns("PDF_Permiso_Sin_Goce"))
            Dias_Aviso_Contrato_Eventual = Val(.rdoColumns("Aviso_Contratacion"))
        End With
    End If
    Rs_Consulta_Cat_Parametros.Close
    Set Rs_Consulta_Cat_Parametros = Nothing
End Sub

