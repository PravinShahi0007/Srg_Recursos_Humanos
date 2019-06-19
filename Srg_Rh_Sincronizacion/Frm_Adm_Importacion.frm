VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Importacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker Dtp_Fecha 
      Height          =   285
      Left            =   585
      TabIndex        =   2
      Top             =   75
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      Format          =   115539969
      CurrentDate     =   41792
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Importacion 
      Height          =   375
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin VB.Timer Tmr_Intervalo 
      Left            =   15
      Top             =   0
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   405
      Left            =   0
      OleObjectBlob   =   "Frm_Adm_Importacion.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Frm_Adm_Importacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As String                     'Define la opcion para los procesos
Dim Bandera As Boolean      'Bandera que se habilita si hubo algun error durante la sincronización

Private Const APP_SYSTRAY_ID = 999 'unique identifier
Private Const NOTIFYICON_VERSION = &H3
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
'manejo de eventos del raton
'Constantes para los botones y el mouse (mensajes)
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Api SetForegroundWindow Para traer la ventana al frente
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

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
    'Conexion_Base.BeginTrans
    For Cont_Fila = 1 To Grid_Importacion.Rows - 1
        If Grid_Importacion.TextMatrix(Cont_Fila, 9) = "S" Then
            'Verifica si el registro ya se ha generado para actualizarlo, si no lo da de alta
            Mi_SQL = "SELECT * FROM Adm_Asistencias_Detalles "
            Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion.TextMatrix(Cont_Fila, 10)) & "'"
            Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
            Set Rs_Modifica_Adm_Asistencias_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Modifica_Adm_Asistencias_Detalles.EOF Then
                'Cambia sólo si no está validada la asistencia
                If Rs_Modifica_Adm_Asistencias_Detalles.rdoColumns("Validada") = "N" Then
                    Mi_SQL = "UPDATE Adm_Asistencias_Detalles"
                    Mi_SQL = Mi_SQL & " SET Hora_Entrada='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Salida='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 6), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Comida_Entrada='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 4), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Hora_Comida_Salida='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 5), "HH:mm:ss") & "'"
                    Mi_SQL = Mi_SQL & " , Horas_Laboradas='" & Val(Grid_Importacion.TextMatrix(Cont_Fila, 7)) & "'"
                    Mi_SQL = Mi_SQL & " , Fecha_Importacion=GETDATE()"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Grid_Importacion.TextMatrix(Cont_Fila, 10)) & "'"
                    Mi_SQL = Mi_SQL & " AND Fecha='" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
                    Conexion_Base.Execute Mi_SQL
                End If
            Else
                Mi_SQL = "INSERT INTO Adm_Asistencias_Detalles(Empleado_ID,No_Tarjeta,Fecha,Hora_Entrada,Hora_Salida,Hora_Comida_Entrada,Hora_Comida_Salida,Horas_Laboradas,Validada,Fecha_Importacion)"
                Mi_SQL = Mi_SQL & " VALUES('" & Trim(Grid_Importacion.TextMatrix(Cont_Fila, 10)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Trim(Grid_Importacion.TextMatrix(Cont_Fila, 1)) & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 0), "MM/dd/yyyy") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 3), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 6), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 4), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Format(Grid_Importacion.TextMatrix(Cont_Fila, 5), "HH:mm:ss") & "'"
                Mi_SQL = Mi_SQL & " , '" & Val(Grid_Importacion.TextMatrix(Cont_Fila, 7)) & "'"
                Mi_SQL = Mi_SQL & " , 'N'"
                Mi_SQL = Mi_SQL & " , GETDATE())"
                Conexion_Base.Execute Mi_SQL
            End If
            Rs_Modifica_Adm_Asistencias_Detalles.Close
        End If
    Next
    'Conexion_Base.CommitTrans
Exit Sub
HANDLER:
    Conexion_Base.RollbackTrans
    Tmr_Intervalo.Enabled = True
End Sub

Private Sub Obtiene_Informacion_Checadas()
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
Dim Rs_Consulta_Dispositivos_Empresa As rdoResultset
Dim Rs_Consulta_Informacion_Dispositivo As rdoResultset
Dim Fecha_Checada As String
Dim Fecha_Inicio_Descarga As Date

On Error GoTo HANDLER
    'Configura el dispositivo
    IP = ""
    Puerto = 0
    Maquina = 0
    CZKEM1.Disconnect
    'Consulta la fecha donde comenzará la descarga de información
    Mi_SQL = "SELECT TOP 1 * FROM Adm_Asistencias_Registro_Checadores"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha DESC,Hora DESC,No_Movimiento DESC"
    Set Rs_Consulta_Informacion_Dispositivo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Informacion_Dispositivo.EOF Then
        If Not IsNull(Rs_Consulta_Informacion_Dispositivo.rdoColumns("Fecha")) Then
            Fecha_Inicio_Descarga = Format(Rs_Consulta_Informacion_Dispositivo.rdoColumns("Fecha"), "MM/dd/yyyy")
        Else
            Fecha_Inicio_Descarga = Format(Now, "MM/dd/yyyy")
        End If
    Else
        Fecha_Inicio_Descarga = Format(Now, "MM/dd/yyyy")
    End If
    Rs_Consulta_Informacion_Dispositivo.Close
    
    'Asigna la fecha de descarga de información
    Dtp_Fecha.Value = Fecha_Inicio_Descarga
    
    'Consulta los checadores de la empresa seleccionada
    Mi_SQL = "SELECT Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Empresas_Equipos_Identificacion.Equipo_ID,Cat_Equipos_Identificadores.No_Equipo,Cat_Equipos_Identificadores.Direccion_IP,Cat_Equipos_Identificadores.Puerto_IP,Cat_Equipos_Identificadores.Descripcion"
    Mi_SQL = Mi_SQL & " FROM Cat_Empresas_Equipos_Identificacion,Cat_Equipos_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empresas_Equipos_Identificacion.Equipo_ID=Cat_Equipos_Identificadores.Equipo_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empresas_Equipos_Identificacion.Empresa_ID='00001'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empresas_Equipos_Identificacion.Empresa_ID,Cat_Equipos_Identificadores.No_Equipo"
    Set Rs_Consulta_Dispositivos_Empresa = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Dispositivos_Empresa.EOF Then
        With Rs_Consulta_Dispositivos_Empresa
            While Not .EOF
                'De acuerdo a los checadores de las empresas inicia la extraccion de informacion
                IP = .rdoColumns("Direccion_IP")
                Puerto = .rdoColumns("Puerto_IP")
                Maquina = .rdoColumns("No_Equipo")
                Descripcion_Equipo = .rdoColumns("Descripcion")
                Checador_ID = Rs_Consulta_Dispositivos_Empresa.rdoColumns("Equipo_ID")
                'Inicia la recoleccion de datos
                If CZKEM1.Connect_Net(IP, Puerto) Then
                Else
                    'No se conectó se va al siguiente checador
                    GoTo SIGUIENTE
                End If
                'Abre los logs en los checadores para extraer la informacion
                res = CZKEM1.GetDeviceStatus(Maquina, 6, dwvalue)
                If res Then
                    If dwvalue = 0 Then
                        'No encontró registros
                        GoTo SIGUIENTE1
                    End If
                End If
                If CZKEM1.ReadGeneralLogData(Maquina) Then
                    CZKEM1.ReadAllUserID Maquina
                    Me.Refresh
                    While CZKEM1.SSR_GetGeneralLogData(Maquina, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode)
                        If CStr(dwInOutMode) > 1 Then
                            dwInOutMode = 0
                        End If
                        Me.Refresh
                        If IsNumeric(dwEnrollNumber) Then
                            cad = Trim(Str(dwEnrollNumber))
                            Fecha_Checada = Format(dwYear, "0000") & "-" + Format(dwMonth, "00") & "-" & Format(dwDay, "00") & " " & Format(dwHour, "00") & ":" & Format(dwMinute, "00") & ":" & Format(dwSecond, "00")
                            If dwEnrollNumber > 0 Then
                                If DateDiff("d", Fecha_Inicio_Descarga, CDate(Fecha_Checada)) >= 0 And DateDiff("d", Format(Now, "MM/dd/yyyy"), CDate(Fecha_Checada)) <= 0 Then
                                    aux = IIf(IsNull(Fecha_Checada), "", Fecha_Checada)
                                    If Len(aux) = 19 Then
                                        aux = cad & Chr(9) & Fecha_Checada
                                        aux = aux & Chr(9) & CStr(dwVerifyMode)
                                        aux = aux & Chr(9) & CStr(dwInOutMode)
                                        aux = aux & Chr(9) & CStr(dwVerifyMode)
                                        aux = aux & Chr(9) & CStr(dwInOutMode)
                                        aux = aux & Chr(9) & Checador_ID
                                        aux = ""
                                        Cuenta = Cuenta + 1
                                        'Verifica si el registro existe, para no duplicar información
                                        Mi_SQL = "SELECT * FROM Adm_Asistencias_Registro_Checadores"
                                        Mi_SQL = Mi_SQL & " WHERE No_Tarjeta='" & Trim(Str(dwEnrollNumber)) & "'"
                                        Mi_SQL = Mi_SQL & " AND Fecha='" & Format(CDate(Fecha_Checada), "MM/dd/yyyy") & "'"
                                        Mi_SQL = Mi_SQL & " AND Hora='" & "12/30/1899 " & Format(CDate(Fecha_Checada), "HH:mm") & "'"
                                        Mi_SQL = Mi_SQL & " AND No_Equipo='" & Maquina & "'"
                                        Mi_SQL = Mi_SQL & " AND Empresa_ID='00001'"
                                        Set Rs_Consulta_Adm_Asistencias_Registro_Checadores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                        If Rs_Consulta_Adm_Asistencias_Registro_Checadores.EOF Then
                                            'Guarda el registro de las checadas en la base de datos
                                            Mi_SQL = "INSERT INTO Adm_Asistencias_Registro_Checadores(No_Tarjeta,Fecha,Hora,Fecha_Importacion"
                                            Mi_SQL = Mi_SQL & " ,No_Equipo,Equipo_ID,Empresa_ID,E_S,IP,Verificacion)"
                                            Mi_SQL = Mi_SQL & " VALUES('" & Trim(Str(dwEnrollNumber)) & "'"
                                            Mi_SQL = Mi_SQL & " , '" & Format(CDate(Fecha_Checada), "MM/dd/yyyy") & "'"
                                            Mi_SQL = Mi_SQL & " , '12/30/1899 " & Format(CDate(Fecha_Checada), "HH:mm") & "'"
                                            'Mi_SQL = Mi_SQL & " , '" & Format(Now, "MM/dd/yyyy") & "'"
                                            Mi_SQL = Mi_SQL & " , GETDATE()"
                                            Mi_SQL = Mi_SQL & " , " & Maquina & ""
                                            Mi_SQL = Mi_SQL & " , '" & Checador_ID & "'"
                                            Mi_SQL = Mi_SQL & " , '00001'"
                                            Mi_SQL = Mi_SQL & " , '" & CStr(dwInOutMode) & "'"
                                            Mi_SQL = Mi_SQL & " , '" & IP & "'"
                                            Mi_SQL = Mi_SQL & " , '" & CStr(dwVerifyMode) & "')"
                                            Conexion_Base.Execute (Mi_SQL)
                                        End If
                                    End If
                                End If
                                DoEvents
                            End If
                        End If
                    Wend
                End If
SIGUIENTE:
SIGUIENTE1:
            CZKEM1.Disconnect
            .MoveNext
            Wend
        End With
    End If
    Rs_Consulta_Dispositivos_Empresa.Close
    Me.Refresh
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Close #1
    Tmr_Intervalo.Enabled = True
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Depurar_Lista
'DESCRIPCION: Depura la lista de información para obtener la hora de entrada y salida
'PARAMETROS : Fecha_Registro- Fecha de Generación de Checadas
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 02-Junio-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Depurar_Lista(Fecha_Registro As Date)
Dim Rs_Consulta_Cat_Empleados As rdoResultset   'informacion del turno del empleado
Dim Rs_Consulta_Cat_Clientes As rdoResultset
Dim Rs_Consulta_Checada As rdoResultset
Dim Cont_Fila As Integer                'Recorre el grid de Grid_Importacion_Keri_System
Dim Cont_Fila_2 As Integer              'Recorre el grid de Grid_Importacion
Dim Cont_Fila_3 As Integer              'Recorre el grid de Grid_Importacion para buscar las horas intermedias
Dim Encontrado As Boolean               'Indica si se ha encontrado el registro en el grid
Dim Cont_Columna As Integer              'Recorre el grid de Grid_Importacion
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

On Error GoTo HANDLER1
    'Llena el grid de acuerdo a la consulta
    Grid_Importacion.Rows = 0
    Grid_Importacion.Cols = 12
    Me.MousePointer = 11
    Me.Refresh
    Hora_Entrada = ""
    Hora_Salida = ""
    Hora_Comida = ""
    Hora_Comida2 = ""
    No_Tarjeta = ""
    Nombre = ""
    Checador = ""
    'Consulta los turnos que son entrada y salida en la misma fecha
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC,Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Turno_ID=Cat_Turnos.Turno_ID"
    'Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='00001'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Fecha_Registro, "MM/dd/yyyy") & "' AND '" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno>=0"
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,AARC.Hora"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        With Rs_Consulta_Cat_Clientes
            Me.Refresh
            Encontrado = False
            Fila_Encontrado = 0
            'LLenado de la informacion
            Empleado_ID = ""
            While Not .EOF
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
                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
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
                        Grid_Importacion.AddItem Format(Fecha, "dd/MMM/yyyy") _
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
                        Me.Refresh
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
                    Grid_Importacion.AddItem Format(Fecha, "dd/MMM/yyyy") _
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
                     Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    'Consulta los turnos que son entrada y salida de dias diferentes
    Mi_SQL = "SELECT DISTINCT AARC.Hora,AARC.Fecha,ISNULL(CE.Apellido_Paterno,'') AS Apellido_Paterno,ISNULL(CE.Apellido_Materno,'') AS Apellido_Materno"
    Mi_SQL = Mi_SQL & " ,ISNULL(CE.Nombre,'') AS Nombre,CE.No_Tarjeta,CE.Empleado_ID,AARC.Equipo_ID,CE.Turno_ID"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE,Adm_Asistencias_Registro_Checadores AARC,Cat_Turnos"
    Mi_SQL = Mi_SQL & " WHERE CE.No_Tarjeta=AARC.No_Tarjeta"
    Mi_SQL = Mi_SQL & " AND CE.Turno_ID=Cat_Turnos.Turno_ID"
    'Mi_SQL = Mi_SQL & " AND CE.Empresa_ID='00001'"
    Mi_SQL = Mi_SQL & " AND AARC.Fecha BETWEEN '" & Format(Fecha_Registro, "MM/dd/yyyy") & "' AND '" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cat_Turnos.Horas_Turno<0"
    Mi_SQL = Mi_SQL & " ORDER BY AARC.Fecha,CE.No_Tarjeta,AARC.Hora"
    Set Rs_Consulta_Cat_Clientes = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Clientes.EOF Then
        With Rs_Consulta_Cat_Clientes
            Me.Refresh
            Encontrado = False
            Fila_Encontrado = 0
            'LLenado de la informacion
            Empleado_ID = ""
            While Not .EOF
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
                    Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                        Empleado_ID = ""
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = .rdoColumns("Hora")
                    Case Else
'                        Hora_Comida = Hora_Comida2
'                        Hora_Comida2 = Hora_Salida
'                        Hora_Salida = Format(.rdoColumns("Hora"), "HH:mm:ss")
                End Select
                'Consulta la hora de entrada del día
                Mi_SQL = "SELECT TOP 1 Adm_Asistencias_Registro_Checadores.No_Tarjeta,Adm_Asistencias_Registro_Checadores.Hora"
                Mi_SQL = Mi_SQL & " FROM Adm_Asistencias_Registro_Checadores,Cat_Turnos"
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='00001'"
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
                Mi_SQL = Mi_SQL & " WHERE Adm_Asistencias_Registro_Checadores.Empresa_ID='00001'"
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
                    'If Empleado_ID <> .rdoColumns("Empleado_ID") Then
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
                        Grid_Importacion.AddItem Format(Fecha, "dd/MMM/yyyy") _
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
                        Me.Refresh
                    'End If
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
                    Grid_Importacion.AddItem Format(Fecha, "dd/MMM/yyyy") _
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
                     Me.Refresh
                End If
            Wend
        End With
    End If
    Rs_Consulta_Cat_Clientes.Close
    Me.Refresh
    With Grid_Importacion
        If Grid_Importacion.Rows > 1 Then
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
        Grid_Importacion.Col = 1
        Grid_Importacion.Sort = flexSortGenericAscending
    End With
    Me.MousePointer = 0
    Me.Refresh
Exit Sub
HANDLER1:
    Tmr_Intervalo.Enabled = True
End Sub

Private Sub Form_Load()
Set Conectar_Ayudante = New Ayudante
    'Valida si no esta abierta ya la aplicacion
    If App.PrevInstance = False Then
        Tmr_Intervalo.Enabled = False           'Deshabilita el contador de tiempo
        Call Conectar_Ayudante.Conexion
        Consulta_Parametros                     'Realiza la consulta de los parametros del sistema
        Dtp_Fecha.Value = Format(Now, "MM/dd/yyyy")
        Call ShellTrayAdd
        Call ShellTrayModifyTip(1)
        Tmr_Intervalo.Interval = ((Val(Mid(Hora_Importacion, 1, 2)) * 3600) + (Val(Mid(Hora_Importacion, 4)) * 60)) * 1000
        Tmr_Intervalo.Enabled = True            'Habilita el contador de tiempo
        Me.Hide
    Else
        End
    End If
End Sub

Private Sub ShellTrayModifyTip(nIconIndex As Long)
Dim nid As NOTIFYICONDATA

   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = Frm_Adm_Importacion.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO Or NIF_ICON
      .dwState = NIS_SHAREDICON
      .hIcon = Frm_Adm_Importacion.Icon
      '.uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .hIcon = Frm_Adm_Importacion.Icon
      'InfoTitle is the balloon tip title, and szInfo is the message displayed.
      'Terminating both with vbNullChar prevents the display of the unused padding in the
      'strings defined as fixed-length in NOTIFYICONDATA.
      .szInfoTitle = "SRG " & vbNullChar
      .szInfo = "" & vbNullChar
   End With
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
End Sub

Private Function IsShellVersion(ByVal version As Long) As Boolean
Dim nBufferSize As Long
Dim nUnused As Long
Dim lpBuffer As Long
Dim nVerMajor As Integer
Dim bBuffer() As Byte
   
    Const sDLLFile As String = "shell32.dll"
    nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
    If nBufferSize > 0 Then
        ReDim bBuffer(nBufferSize - 1) As Byte
        Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
        If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
            CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
            IsShellVersion = nVerMajor >= version
        End If
    End If
End Function

Private Sub SetShellVersion()
   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select
End Sub

Private Sub ShellTrayAdd()
Dim nid As NOTIFYICONDATA
   
    If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
    'set up the type members
    With nid
        .cbSize = NOTIFYICONDATA_SIZE
        .hwnd = Frm_Adm_Importacion.hwnd
        .uID = APP_SYSTRAY_ID
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .dwState = NIS_SHAREDICON
        .hIcon = Frm_Adm_Importacion.Icon
        .uCallbackMessage = WM_MOUSEMOVE
        'szTip is the tooltip shown when the mouse hovers over the systray icon.
        'Terminate it since the strings are fixed-length in NOTIFYICONDATA
        .szTip = "SRG " & App.Major & "." & App.Minor & "." & App.Revision & vbNullChar
        .uTimeoutAndVersion = NOTIFYICON_VERSION
    End With
    'add the icon ...
    Call Shell_NotifyIcon(NIM_ADD, nid)
    '... and inform the system of the NOTIFYICON version in use
    Call Shell_NotifyIcon(NIM_SETVERSION, nid)
End Sub

Private Sub Tmr_Intervalo_Timer()
On Error GoTo ErrorHandler

    Tmr_Intervalo.Enabled = False
    'Descarga las checadas del aparato
    Call Obtiene_Informacion_Checadas
    'Actualiza las asistencias del día
    Call Depurar_Lista(Format(Dtp_Fecha.Value, "MM/dd/yyyy"))
    'Guarda los cambios
    Alta_Importacion_Asistencias
    Tmr_Intervalo.Enabled = True
    End
Exit Sub
ErrorHandler:
    Tmr_Intervalo.Enabled = True
End Sub

