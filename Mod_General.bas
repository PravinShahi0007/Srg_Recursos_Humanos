Attribute VB_Name = "Mod_General"
Public Mi_SQL As String                             'Obtiene los valores de la consulta
Public Catalogo As String                           'Indicar que formulario se va a abrir
Public Rdo_Error As rdoError                        'Especifica que tipo de error se esta cometiendo
Public Conexion_Base As New rdoConnection           'Se utiliza para la conexion a la base de datos de Facturas
Public Conexion_Base_CxC As New rdoConnection       'Se utiliza para la conexion a la base de datos de CxC
Public Conexion_CxC As Boolean                      'Se utiliza poara verificar si esta conectado o no
Public Conexion_Base_Inv As New rdoConnection       'Se utiliza para la conexion a la base de datos de Inventario
Public Conexion_Inv As Boolean                      'Se utiliza poara verificar si esta conectado o no
Public Nombre_Usuario As String                     'Obtiene el nombre del usuario
Public Ciclos As Integer                            'Se utiliza para válidar el tiempo de espera de la pantalla de presentación
Public Conectar_Ayudante As Ayudante                'Es utilizada para ligar a la ayuda
Public Conectar_Ayudante_CxC As Ayudante_CxC        'Es utilizada para ligar a la ayuda
Public Conectar_Ayudante_Inv As Ayudante_Inv        'Es utilizada para ligar a la ayuda
Public Par_Fecha As String
Public Base_Datos As String                         'Indica el nombre de la base de datos a conectarse
Public Server As String                             'Indica el nombre del servidor en donde se encuentra la base de datos
Public Rol_ID As String                             'Obtiene el rol que tiene asignado el usuario
Public Usuario_ID As String                         'Alamacena el ID del usuario registrado
Public Intentos_Fallidos As Integer                 'Indica el número de intentos fallidos que puede tener un usuario para deshabilitar la cuenta
Public Bloqueo_Por_No_Utilizar As Integer           'Almacena el parametyro de diferencia de dias
Public Bloqueo_Por_Expiración_Password As Integer   'Almacena el parametyro de diferencia de dias
Public Rol_Administrador As String                  'Almacena el rol del admisnitrador para hacer comparativas de seguridad
Public Tipo_Validacion As String                    'Almacena el tipo de utilidad que tenda la ventana de loguin
Public Nombre_Base_Datos As String                  'Indica el nombre de la base de datos a conectarse
Public Usuario_Autorizo As String                   'Almacena el usuario de las autorizaciones

Public Empresa As String                            'Almacena el nombre de la empresa
Public Empresa_Abreviatura As String                'Almacena el nombre corto de la empresa
Public RFC As String                                'Almacena el RFC de la empresa
Public Expedida_En As String                        'Almacena la expedición de las facturas
Public Direccion As String                          'Almacena la direccion de la empresa
Public Colonia As String                            'Almacena la colonia de la empresa
Public Numero_Exterior As String                    'Almacena el no exterior de la empresa
Public Numero_Interior As String                    'Almacena el no interior de la empresa
Public CP As String                                 'Almacena el codigo postal de la empresa
Public Ciudad As String                             'Almacena la ciudad de la empresa
Public Estado As String                             'Almacena el estado de la empresa
Public Pais As String                               'Almacena el pais de la empresa
Public GLN As String                                'Almacena el codigo gln de la empresa
Public Codigo_Proveedor As String                   'almacena el codigo del proveedor
Public Password As String                           'Guarda la contraseña para el acceso al servidor
Public Usuario_Conexion As String                   'Guarda el nombre válido del usuario para el acceso a la BD

Public Ruta_Certificado As String                   'Almacena la ruta del certificado
Public Ruta_Llave_Privada As String                 'Almacena la ruta de la llave privada
Public Password_Llave As String                     'Almacena el password de la llave privada
Public RFC_Empresa As String                        'Almacena el RFC de la empresa
Public RFC_Mostrador As String                      'Almacena el RFC generico
Public Ruta_Pdfs As String                          'Almacena la ruta de los pdfs
Public Ruta_Xmls As String                          'Almacena la ruta de los xmls
Public Ruta_PreFacturas As String                   'Almacena la ruta de la prefacturas
Public Tasa_Impuesto_IVA As Integer                 'Almacena la tasa de impuesto del iva
Public Texto_Factura As String                      'Almacena el texto a imprimir en la factura

Public Ruta_Bimbo_Entrada As String                 'Almacena la ruta de la carpeta de entrada de bimbo
Public Ruta_Bimbo_Salida As String                  'Almacena la ruta de la carpeta de salida de bimbo
Public Servidor_Bimbo As String                     'Almacena el servidor de conexión de bimbo
Public Puerto_Bimbo As Integer                      'Almacena el puerto de conexión de bimbo
Public Usuario_Bimbo As String                      'Almacena el nombre de usuario de bimbo
Public Password_Bimbo As String                     'Almacena el password del usuario de bimbo
Public Buzon_Bimbo As String                        'Almacena el buzon de bimbo
Public Cliente_ID_Bimbo As String                   'Almacena el id de cliente bimbo
Public Cliente_RFC_Bimbo As String                  'Almacena el rfc del cliente bimbo
Public Cliente_ID_Soriana As String                 'Almacena el id de cliente soriana
Public Cliente_RFC_Soriana As String                'Almacena el rfc del cliente soriana
Public WebService_Url_Soriana As String             'Almacena la url del webservice de soriana
Public Ruta_Soriana As String                       'Almacena la ruta de los dcumentos de soriana
Public Mensaje_Error As String                      'Almacena el mensaje de error

'Obtiene el directorio ya sea windows o winnt
Public Declare Function GetWindowsDirectory Lib "kernel32" _
      Alias "GetWindowsDirectoryA" ( _
      ByVal lpBuffer As String, _
      ByVal nSize As Long) As Long

'Api para saber si una carpeta existe
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'Variables para manejo de la ventana de seleccion de directorio
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
' Constantes
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260 ' Para Buffer de caracteres del path
' Funcion Api CoTaskMemFree
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
' Funcion Api CoTaskMemFree lstrcat
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
' Funcion Api SHBrowseForFolder
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
' Funcion Api SHGetPathFromIDList
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long

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
    'NOMBRE DE LA FUNCIÓN: Aviso_Termino_Folios
    'DESCRIPCIÓN: Realiza la consulta del parametro para avisar que se terminan los folios
    'PARÁMETROS :
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 24/Octubre/2009 9:35am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Aviso_Termino_Folios()
Dim Rs_Consulta_Parametro_Folio As rdoResultset 'Variable para el manejo de la tabla
Dim Rs_Consulta_Parametros As rdoResultset      'Variable para el manejo de la tabla
Dim Rs_Consulta_Factura_Folio As rdoResultset   'Variable para el manejo de la tabla
Dim Termino_Folios As Double                    'Almacena el parametro de los folios
Dim Folio_Final As Double                       'Almacena el folio final activo
Dim Folio_Factura As Double                     'Almacena el folio de la factura

On Error GoTo errorHandler
    
    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Realiza la consulta del parametro
    Mi_SQL = "SELECT Dias_Aviso_Termina_Folios FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Parametros
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Dias_Aviso_Termina_Folios")) Then
                Termino_Folios = Val(.rdoColumns("Dias_Aviso_Termina_Folios"))
            Else
                Termino_Folios = 0
            End If
        Else
            Termino_Folios = 0
        End If
    End With
    Rs_Consulta_Parametros.Close
    
    If Termino_Folios > 0 Then
        'Consulta los parametros del folio final
        Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Serie = ''"
        Mi_SQL = Mi_SQL & " AND Estatus <> 'CANCELADO'"
        Set Rs_Consulta_Parametro_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametro_Folio
            If Not .EOF Then
                Folio_Final = Val(.rdoColumns("Folio_Final"))
            Else
                Folio_Final = 0
            End If
        End With
        Rs_Consulta_Parametro_Folio.Close
        
        'Consulta el folio final de la factura
        Mi_SQL = "SELECT ISNULL(MAX(IdFactura),0) AS IdFactura FROM FacturaE"
        Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Factura_Folio
            If Not .EOF Then
                Folio_Factura = Val(.rdoColumns("IdFactura"))
            Else
                Folio_Factura = 0
            End If
        End With
        Rs_Consulta_Factura_Folio.Close
        
        'Realiza la validacion de los folios
        If (Folio_Final - Folio_Factura) <= Termino_Folios Then
            MDIFrm_Apl_Principal.MousePointer = 0
            MsgBox "Le quedan disponibles " & Folio_Final - Folio_Factura - 1 & " folios para las facturas electrónicas, favor de solicitar mas en la página del SAT."
        End If
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Aviso_Vigencia_Certificado
    'DESCRIPCIÓN: Realiza la consulta del parametro para avisar de la vigencia del certificado
    'PARÁMETROS : Fecha_Factura, fecha de la factura a valida, u opcional, para validar el aviso
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 24/Octubre/2009 10:15am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Aviso_Vigencia_Certificado(Optional Fecha_Factura As Date = "01/01/1900") As Boolean
Dim Rs_Consulta_Parametro_Folio As rdoResultset 'Variable para el manejo de la tabla
Dim Dias_Aviso_Expira_Vigencia As Integer       'Almacena el parametro de la vigencia
Dim Vigencia_Certificado_Desde As Date          'Almacena la vigencia desde
Dim Vigencia_Certificado_Hasta As Date          'Almacena la vigencia hasta

On Error GoTo errorHandler
    
    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Realiza la consulta del parametro
    Mi_SQL = "SELECT Dias_Aviso_Expira_Vigencia, Vigencia_Certificado_Desde, Vigencia_Certificado_Hasta FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Parametros
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Dias_Aviso_Expira_Vigencia")) Then
                Dias_Aviso_Expira_Vigencia = Val(.rdoColumns("Dias_Aviso_Expira_Vigencia"))
            Else
                Dias_Aviso_Expira_Vigencia = 0
            End If
            If Not IsNull(.rdoColumns("Vigencia_Certificado_Desde")) Then
                Vigencia_Certificado_Desde = .rdoColumns("Vigencia_Certificado_Desde")
            Else
                Vigencia_Certificado_Desde = "01/01/1900"
            End If
            If Not IsNull(.rdoColumns("Vigencia_Certificado_Hasta")) Then
                Vigencia_Certificado_Hasta = .rdoColumns("Vigencia_Certificado_Hasta")
            Else
                Vigencia_Certificado_Hasta = "01/01/1900"
            End If
        Else
            Dias_Aviso_Expira_Vigencia = 0
            Vigencia_Certificado_Desde = "01/01/1900"
            Vigencia_Certificado_Hasta = "01/01/1900"
        End If
    End With
    Rs_Consulta_Parametros.Close
        
    If Fecha_Factura = "01/01/1900" Then
        'Realiza la comparación de la vigencia
        Dias_Diferencia_Vigencia = DateDiff("d", Vigencia_Certificado_Hasta, Now) * -1
        If Dias_Diferencia_Vigencia = 0 Then
            Aviso_Vigencia_Certificado = True
            MsgBox "El certificado expira el día de hoy."
        Else
            If Dias_Diferencia_Vigencia < 0 Then
                Aviso_Vigencia_Certificado = False
                MsgBox "El ha expirado, favor de solicitar la renovación en la página del SAT."
            Else
                If Dias_Aviso_Expira_Vigencia >= Dias_Diferencia_Vigencia Then
                    Aviso_Vigencia_Certificado = True
                    MsgBox "El certificado esta a punto de expirar, le quedan " & Dias_Diferencia_Vigencia & " días."
                End If
            End If
        End If
    Else
        'Realiza la comparación de la vigencia
        If Format(Fecha_Factura, "yyyy/MM/dddd") >= Format(Vigencia_Certificado_Desde, "yyyy/MM/dd") And Format(Fecha_Factura, "yyyy/MM/dddd") <= Format(Vigencia_Certificado_Hasta, "yyyy/MM/dd") Then
            Aviso_Vigencia_Certificado = True
        Else
            Aviso_Vigencia_Certificado = False
        End If
    End If
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Function
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Function



'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Parametros_Factura
    'DESCRIPCIÓN: Consulta la tabla de los parametros de factura electronica
    'PARÁMETROS :
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 01/Sep/2009 1:00pm
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Parametros_Factura()
Dim Rs_Consulta_Parametros As rdoResultset  'Variable para el manejo de la tabla

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Consulta la tabla de parametros de factura
    Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Parametros
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Ruta_Certificado")) Then Ruta_Certificado = .rdoColumns("Ruta_Certificado")
            If Not IsNull(.rdoColumns("Ruta_Llave_Privada")) Then Ruta_Llave_Privada = .rdoColumns("Ruta_Llave_Privada")
            If Not IsNull(.rdoColumns("Password_Llave_Privada")) Then Password_Llave = MySecret_Decrypt(Trim(.rdoColumns("Password_Llave_Privada")), "hisa2009")
            If Not IsNull(.rdoColumns("RFC_Empresa")) Then RFC_Empresa = .rdoColumns("RFC_Empresa")
            If Not IsNull(.rdoColumns("RFC_Mostrador")) Then RFC_Mostrador = .rdoColumns("RFC_Mostrador")
            If Not IsNull(.rdoColumns("Ruta_Pdfs")) Then Ruta_Pdfs = .rdoColumns("Ruta_Pdfs")
            If Not IsNull(.rdoColumns("Ruta_Xmls")) Then Ruta_Xmls = .rdoColumns("Ruta_Xmls")
            If Not IsNull(.rdoColumns("Ruta_PreFactura")) Then Ruta_PreFacturas = .rdoColumns("Ruta_PreFactura")
            If Not IsNull(.rdoColumns("Ruta_Entrada_Bimbo")) Then Ruta_Bimbo_Entrada = .rdoColumns("Ruta_Entrada_Bimbo")
            If Not IsNull(.rdoColumns("Ruta_Salida_Bimbo")) Then Ruta_Bimbo_Salida = .rdoColumns("Ruta_Salida_Bimbo")
            If Not IsNull(.rdoColumns("Servidor_Bimbo")) Then Servidor_Bimbo = .rdoColumns("Servidor_Bimbo")
            If Not IsNull(.rdoColumns("Puerto_Bimbo")) Then Puerto_Bimbo = .rdoColumns("Puerto_Bimbo")
            If Not IsNull(.rdoColumns("Usuario_Bimbo")) Then Usuario_Bimbo = .rdoColumns("Usuario_Bimbo")
            If Not IsNull(.rdoColumns("Password_Bimbo")) Then Password_Bimbo = MySecret_Decrypt(Trim(.rdoColumns("Password_Bimbo")), "hisa2009")
            If Not IsNull(.rdoColumns("Buzon_Bimbo")) Then Buzon_Bimbo = .rdoColumns("Buzon_Bimbo")
            If Not IsNull(.rdoColumns("Cliente_ID_Bimbo")) Then Cliente_ID_Bimbo = .rdoColumns("Cliente_ID_Bimbo")
            If Not IsNull(.rdoColumns("Cliente_RFC_Bimbo")) Then Cliente_RFC_Bimbo = .rdoColumns("Cliente_RFC_Bimbo")
            If Not IsNull(.rdoColumns("Ruta_Soriana")) Then Ruta_Soriana = .rdoColumns("Ruta_Soriana")
            If Not IsNull(.rdoColumns("WebService_Url_Soriana")) Then WebService_Url_Soriana = .rdoColumns("WebService_Url_Soriana")
            If Not IsNull(.rdoColumns("Cliente_ID_Soriana")) Then Cliente_ID_Soriana = .rdoColumns("Cliente_ID_Soriana")
            If Not IsNull(.rdoColumns("Cliente_RFC_Soriana")) Then Cliente_RFC_Soriana = .rdoColumns("Cliente_RFC_Soriana")
        End If
    End With
    Rs_Consulta_Parametros.Close
    
    'Consulta la tabla de parametros de control
    Mi_SQL = "SELECT Impuesto, TextoFactura FROM Control"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Parametros
        If Not .EOF Then
            If Not IsNull(.rdoColumns("Impuesto")) Then Tasa_Impuesto_IVA = .rdoColumns("Impuesto")
            If Not IsNull(.rdoColumns("TextoFactura")) Then Texto_Factura = .rdoColumns("TextoFactura")
        End If
    End With
    Rs_Consulta_Parametros.Close
    
    MDIFrm_Apl_Principal.MousePointer = 0
Exit Sub
errorHandler:
    For Each Rdo_Error In rdoErrors
        MsgBox Rdo_Error.Description
    Next
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Envia_Facturas_Bimbo
    'DESCRIPCIÓN: Realiza el envio de las facturas a Bimbo mediante la interfaz
    'PARÁMETROS : No_Factura, numero de factura a enviar a Bimbo
    '             Mensaje, mensaje de error captado en el proceso
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 14/Sep/2009 11:30am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Envia_Xmls_Bimbo(No_Factura As String, ByRef Mensaje As String) As Integer
Dim Ejecucion As String         'Almacena la cadena de ejecución
Dim Ruta_Ejecucion As String    'Almacena la ruta en que se debe de ejecutar
Dim Ruta_Log_Actual As String   'Almacena la ruta del log del dia
Dim hProcess As Long            'Indica que se ejecuta el proceso de MS-DOS
Dim Archivo_Bat As String       'Almacena el nombre del archivo bat
Dim Lista_Archivos As Object    'Almacena el listado de archivos
Dim Archivo As Object           'Almacena el archivo actual
Dim Ultimo_Archivo As Object    'Almacena el ultimo archivo
Dim Linea As String             'Almacena la lectura de la Linea

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Valida que exista el archivo en la carpeta de salida
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Xmls & "\CFD_" & No_Factura & ".XML", "ARCHIVO") = True Then
        'Realiza la copia del archivo xml a la carpeta de salida
        FileCopy Ruta_Xmls & "\CFD_" & No_Factura & ".XML", Ruta_Bimbo_Salida & "\CFD_" & No_Factura & ".XML"

        'Valida que exista el archivo en la carpeta de salida
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Bimbo_Salida & "\CFD_" & No_Factura & ".XML", "ARCHIVO") = True Then
            'Asigna la ruta de ejecucion
            Ruta_Ejecucion = Mid(Ruta_Bimbo_Salida, 1, Len(Ruta_Bimbo_Salida) - 6)
            'Asigna las instrucciones de ejecucion
            Ejecucion = Ruta_Ejecucion & "SEDEB2BONLINE.EXE /SERVER:" & Servidor_Bimbo & " /PORT:" & Puerto_Bimbo & " /SSL /USER:" & Usuario_Bimbo & " /PWD:" & Password_Bimbo & " /PUT:""" & Ruta_Bimbo_Salida & "\*.XML"" /DESTINO:""" & Buzon_Bimbo & """ /REN:*.ENV"
            'Asigna el bat
            Archivo_Bat = "ENVIO_HARINERA_" & Nombre_Usuario & ".bat"
            Open "C:\" & Archivo_Bat For Output As #1
            Print #1, Ejecucion
            Close #1
            'Ejecuta el proceso
            hProcess = OpenProcess(SYNCHRONIZE, 0, Shell("C:\" & Archivo_Bat, vbHide))
            'Indica si se termino de procesar la información para poder continuar con las siguientes
            'ejecuciones
            If hProcess Then
                WaitForSingleObject hProcess, INFINITE
                CloseHandle hProcess
            End If
            
            'Elimina el archivo bat
            'Valida que exista el archivo
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta("C:\" & Archivo_Bat, "ARCHIVO") = True Then
                'Elimina el archivo
                Kill "C:\" & Archivo_Bat
            End If
            
            'Revisa en el log que se haya enviado el archivo
            Set Lista_Archivos = CreateObject("Scripting.FileSystemObject")
            For Each Archivo In Lista_Archivos.GetFolder(Ruta_Ejecucion & "LOGS").Files
                If Ultimo_Archivo Is Nothing Then Set Ultimo_Archivo = Archivo
                'Valida si son archivos viejos los copia a su carpeta correspondiente
                If DateDiff("d", Now, Archivo.DateLastModified) < 0 Or Archivo.DateLastModified = Ultimo_Archivo.DateLastModified Then
                    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy"), "CARPETA") = False Then
                        MkDir Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy")
                    End If
                    'Copia el archivo a la carpeta de respaldo
                    FileCopy Archivo, Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy") & "\" & Archivo.Name
                End If
                'Verifica cual es el ultimo archivo
                If Archivo.DateLastModified > Ultimo_Archivo.DateLastModified Then
                    Set Ultimo_Archivo = Archivo
                End If
            Next
            'Elimina los archivos antiguos
            For Each Archivo In Lista_Archivos.GetFolder(Ruta_Ejecucion & "LOGS").Files
                If Archivo <> Ultimo_Archivo Then
                    Kill Archivo
                End If
            Next
            'Lee el archivo actual
            Open Ultimo_Archivo For Input As #1
            Do While Not EOF(1)
                'Lee linea por linea
                Line Input #1, Linea
                'Valida los diferentes mensajes
                If Trim(Linea) <> "" Then
                    'Si no habia archivos que enviar
                    If InStr(1, Linea, "No hay mensajes para enviar") > 0 Then
                        Mensaje = "No hay mensajes para enviar."
                        Envia_Xmls_Bimbo = 0
                        Exit Do
                    Else
                        'Si se enviaron los archivo
                        If InStr(1, Linea, "Fichero enviado") > 0 Then
                            Mensaje = ""
                            Envia_Xmls_Bimbo = 1
                        Else 'Si hubo error
                            If InStr(1, Linea, "Error:") > 0 Then
                                Mensaje = Trim(Mid(Linea, InStr(1, Linea, "Error:") + 1))
                                Envia_Xmls_Bimbo = -1
                                Exit Do
                            End If
                        End If
                    End If
                End If
            Loop
            Close #1
            'Elimina el archivo actual
            Kill Ultimo_Archivo
            'Reinicia las variables
            Set Archivo = Nothing
            Set Ultimo_Archivo = Nothing
            Set Lista_Archivos = Nothing
        Else
            Mensaje = "No existe el archivo " & No_Factura & ".xml en la carpeta de salida, favor de verificarlo."
            Envia_Xmls_Bimbo = -1
        End If
    Else
        Mensaje = "No existe el archivo " & No_Factura & ".xml generado, favor de regenerar la factura."
        Envia_Xmls_Bimbo = -1
    End If
        
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Function
errorHandler:
    Mensaje = Err.Description
    Envia_Xmls_Bimbo = -1
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Envia_Xmls_Soriana
    'DESCRIPCIÓN: Realiza el envio de las facturas a Soriana mediante la interfaz
    'PARÁMETROS : No_Factura, numero de factura a enviar a Soriana
    '             Mensaje, mensaje de error captado en el proceso
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 26/Ene/2010 12:30pm
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Envia_Xmls_Soriana(No_Factura As String, ByRef Mensaje As String) As Integer
Dim Ejecucion As String         'Almacena la cadena de ejecución
Dim Ruta_Ejecucion As String    'Almacena la ruta en que se debe de ejecutar
Dim Ruta_Log_Actual As String   'Almacena la ruta del log del dia
Dim hProcess As Long            'Indica que se ejecuta el proceso de MS-DOS
Dim Archivo_Bat As String       'Almacena el nombre del archivo bat
Dim Lista_Archivos As Object    'Almacena el listado de archivos
Dim Archivo As Object           'Almacena el archivo actual
Dim Ultimo_Archivo As Object    'Almacena el ultimo archivo
Dim Linea As String             'Almacena la lectura de la Linea
Dim Xml_Notificacion As DOMDocument   'Almacena la notificacion
Dim Lista_Nodos As IXMLDOMNodeList  'Variable para recorrer los nodos
Dim Nodo As IXMLDOMNode             'Variable para obtener el nodo
Dim Nodo_1 As IXMLDOMNode             'Variable para obtener el nodo
Dim Nodo_Atributo As IXMLDOMAttribute      'Variable para obtener el nodo
Dim Encontro As Integer         'Almacena si encontro el nodo en alguna posición
Dim Mensaje_Notificacion As String  'Almacena el mensaje de notificacion de cada documento
Dim No_PreFacturas As Integer       'Almacena la cantidad de prefacturas

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Valida que exista el archivo en la carpeta de salida
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Xmls & "\CFD_" & No_Factura & ".XML", "ARCHIVO") = True Then
        'Realiza la copia del archivo xml a la carpeta de salida
        FileCopy Ruta_Xmls & "\CFD_" & No_Factura & ".XML", Ruta_Soriana & "\Salida\CFD_" & No_Factura & ".XML"

        'Valida que exista el archivo en la carpeta de salida
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Soriana & "\Salida\CFD_" & No_Factura & ".XML", "ARCHIVO") = True Then
            'Asigna la ruta de ejecucion
            Ruta_Ejecucion = Ruta_Soriana
            'Asigna las instrucciones de ejecucion
            Ejecucion = Ruta_Ejecucion & "\HISA_Soriana_Interfaz.exe"
            'Asigna el bat
            Archivo_Bat = "ENVIO_HARINERA_" & Nombre_Usuario & ".bat"
            Open "C:\" & Archivo_Bat For Output As #1
            Print #1, Ejecucion
            Close #1
            'Ejecuta el proceso
            hProcess = OpenProcess(SYNCHRONIZE, 0, Shell("C:\" & Archivo_Bat, vbHide))
            'Indica si se termino de procesar la información para poder continuar con las siguientes
            'ejecuciones
            If hProcess Then
                WaitForSingleObject hProcess, INFINITE
                CloseHandle hProcess
            End If
            
            'Elimina el archivo bat
            'Valida que exista el archivo
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta("C:\" & Archivo_Bat, "ARCHIVO") = True Then
                'Elimina el archivo
                Kill "C:\" & Archivo_Bat
            End If
            
            'Revisa en el log que se haya enviado el archivo
            Set Lista_Archivos = CreateObject("Scripting.FileSystemObject")
            For Each Archivo In Lista_Archivos.GetFolder(Ruta_Soriana & "\LOG").Files
                If Ultimo_Archivo Is Nothing Then Set Ultimo_Archivo = Archivo
                'Valida si son archivos viejos los copia a su carpeta correspondiente
                If DateDiff("d", Now, Archivo.DateLastModified) < 0 Or Archivo.DateLastModified = Ultimo_Archivo.DateLastModified Then
                    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Soriana & "\LOG\" & Format(Archivo.DateLastModified, "ddMMMyyyy"), "CARPETA") = False Then
                        MkDir Ruta_Soriana & "\LOG\" & Format(Archivo.DateLastModified, "ddMMMyyyy")
                    End If
                    'Copia el archivo a la carpeta de respaldo
                    FileCopy Archivo, Ruta_Soriana & "\LOG\" & Format(Archivo.DateLastModified, "ddMMMyyyy") & "\" & Archivo.Name
                End If
                'Verifica cual es el ultimo archivo
                If Archivo.DateLastModified > Ultimo_Archivo.DateLastModified Then
                    Set Ultimo_Archivo = Archivo
                End If
            Next
            'Elimina los archivos antiguos
            For Each Archivo In Lista_Archivos.GetFolder(Ruta_Soriana & "\LOG").Files
                If Archivo <> Ultimo_Archivo Then
                    Kill Archivo
                End If
            Next
            'Lee el archivo actual
            Open Ultimo_Archivo For Input As #1
            Do While Not EOF(1)
                'Lee linea por linea
                Line Input #1, Linea
                'Valida los diferentes mensajes
                If Trim(Linea) <> "" Then
                    'Si no habia archivos que enviar
                    If InStr(1, Linea, "No hay mensajes para enviar") > 0 Then
                        Mensaje = "No hay mensajes para enviar."
                        No_Notificaciones = 0
                        Exit Do
                    Else
                        'Si se enviaron los archivo
                        If InStr(1, Linea, "Archivos Enviados") > 0 Then
                            Mensaje = ""
                            No_Notificaciones = 1
                            Exit Do
                        Else 'Si hubo error
                            If InStr(1, Linea, "Error:") > 0 Then
                                Mensaje = Trim(Mid(Linea, InStr(1, Linea, "Error:") + 1))
                                No_Notificaciones = -1
                                Exit Do
                            End If
                        End If
                    End If
                End If
            Loop
            Close #1
            'Elimina el archivo actual
            Kill Ultimo_Archivo
            'Reinicia las variables
            Set Archivo = Nothing
            Set Ultimo_Archivo = Nothing
            Set Lista_Archivos = Nothing
            
            'Copia los archivo a la prefactura
            Set Lista_Archivos = CreateObject("Scripting.FileSystemObject")
            For Each Archivo In Lista_Archivos.GetFolder(Ruta_Soriana & "\Entrada").Files
                
                'Poner los datos en el analizador de XML
                Set Xml_Notificacion = New DOMDocument
                
                'Valida el xml
                Xml_Notificacion.resolveExternals = True
            
                'Para que valide el documento xml
                Xml_Notificacion.validateOnParse = True
            
                'Agiliza la carga del documento
                Xml_Notificacion.async = False
                
                'Carga el documento
                If Xml_Notificacion.Load(Trim(Archivo)) Then
                    'Comprobamos si se carga
                    If Xml_Notificacion.parseError.reason = "" Then
                        'Valida si es una prefactura con el nombre del nodo
                        If Xml_Notificacion.documentElement.nodeName = "AckErrorApplication" And Xml_Notificacion.lastChild.nodeName = "AckErrorApplication" Then
                            'Recorre el documento xml
                            Set Lista_Nodos = Xml_Notificacion.selectNodes("//")
                            For Each Nodo In Lista_Nodos
                                'Revisa el nodo de número de documento
                                If Nodo.nodeName = "ReferenceNumber" Then
                                    Set Nodo_Atributo = Nodo.Attributes(0)
                                    If Not Nodo_Atributo Is Nothing Then
                                        If Nodo.Attributes(0).Text = "IV" Then
                                            If Trim(Mensaje_Notificacion) = "" Then
                                                Mensaje_Notificacion = "Factura --> " & Nodo.Text
                                            Else
                                                Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Factura --> " & Nodo.Text
                                            End If
                                        End If
                                    End If
                                End If
                                'Revisa el nodo de mensajes
                                If Nodo.nodeName = "messageError" Then
                                    Set Nodo_1 = Nodo.selectSingleNode("errorDescription/text")
                                    If Not Nodo_1 Is Nothing Then
                                        If Trim(Mensaje_Notificacion) = "" Then
                                            Mensaje_Notificacion = "Notificación: " & Nodo_1.Text
                                        Else
                                            Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Notificación: " & Nodo_1.Text
                                        End If
                                    End If
                                End If
                            Next
                            Set Lista_Nodos = Nothing
                            Set Nodo = Nothing
                            
                            'Copia el archivo a la ruta de notificaciones
                            FileCopy Archivo, Ruta_Soriana & "\Notificaciones\" & Archivo.Name
                        End If
                    End If
                Else
                    'Abre el archivo como lectura
                    AckErrorApplication = False
                    Open Archivo For Input As #1
                    If Not EOF(1) Then
                        Do While Not EOF(1)
                            Line Input #1, Linea
                            'Reemplaza las comillas dobles
                            Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                            'Valida si es una archivo de notificacion
                            Encontro = InStr(1, Linea, "AckErrorApplication", 0)
                            If Encontro > 0 Then
                                AckErrorApplication = True
                            End If
                            'Si es un archivo de notificacion continua con el proceso
                            If AckErrorApplication = True Then
                                Encontro = InStr(1, Linea, "<ReferenceNumber type =  IV >", 0)
                                If Encontro > 0 Then
                                    Line Input #1, Linea
                                    'Reemplaza las comillas dobles
                                    Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                                    Encontro = InStr(1, Linea, ">", 0)
                                    If Encontro > 0 Then
                                        Linea = Mid(Linea, Encontro + 1)
                                        Encontro = InStr(1, Linea, "<", 0)
                                        If Encontro > 0 Then
                                            Linea = Mid(Linea, 1, Encontro - 1)
                                            If Trim(Mensaje_Notificacion) = "" Then
                                                Mensaje_Notificacion = "Factura --> " & Linea
                                            Else
                                                Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Factura --> " & Linea
                                            End If
                                            No_Notificaciones = No_Notificaciones + 1
                                        End If
                                    End If
                                End If
                                Encontro = InStr(1, Linea, "messageError sequence", 0)
                                If Encontro > 0 Then
                                    Line Input #1, Linea
                                    Line Input #1, Linea
                                    Line Input #1, Linea
                                    'Reemplaza las comillas dobles
                                    Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                                    Encontro = InStr(1, Linea, ">>", 0)
                                    If Encontro > 0 Then
                                        Linea = Mid(Linea, Encontro + 2)
                                        Encontro = InStr(1, Linea, "<", 0)
                                        If Encontro > 0 Then
                                            Linea = Mid(Linea, 1, Encontro - 1)
                                            Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Error --> " & Linea
                                        End If
                                    Else
                                        Encontro = InStr(1, Linea, ">", 0)
                                        If Encontro > 0 Then
                                            Linea = Mid(Linea, Encontro + 1)
                                            Encontro = InStr(1, Linea, "<", 0)
                                            If Encontro > 0 Then
                                                Linea = Mid(Linea, 1, Encontro - 1)
                                                Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Exito --> " & Linea
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Loop
                    End If
                    Close #1
                    'Si no es archivo de notificacion lo elimina
                    If AckErrorApplication = False Then GoTo 1
                    'Copia el archivo a la ruta de notificaciones
                    FileCopy Archivo, Ruta_Ejecucion & "Notificaciones\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 3) & ".xml"
                End If
                'Elimina el archivo
1:              Kill Archivo
                'Resetea la variable
                Set Xml_Notificacion = Nothing
            Next
            Set Archivo = Nothing
            
            Mensaje = "Cantidad Notificaciones --> " & No_Notificaciones & Chr(13) & Mensaje_Notificacion
            Envia_Xmls_Soriana = 1
        Else
            Mensaje = "No existe el archivo " & No_Factura & ".xml en la carpeta de salida, favor de verificarlo."
            Envia_Xmls_Soriana = -1
        End If
    Else
        Mensaje = "No existe el archivo " & No_Factura & ".xml generado, favor de regenerar la factura."
        Envia_Xmls_Soriana = -1
    End If
        
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Function
errorHandler:
    Mensaje = Err.Description
    Envia_Xmls_Soriana = -1
End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recibe_Xmls_Bimbo
    'DESCRIPCIÓN: Realiza la recepción de los xmls de Bimbo mediante la interfaz
    'PARÁMETROS : Mensaje, mensaje de error captado en el proceso
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 15/Sep/2009 10:40am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Recibe_Xmls_Bimbo(ByRef Mensaje As String) As Integer
Dim Ejecucion As String         'Almacena la cadena de ejecución
Dim Ruta_Ejecucion As String    'Almacena la ruta en que se debe de ejecutar
Dim Ruta_Log_Actual As String   'Almacena la ruta del log del dia
Dim hProcess As Long            'Indica que se ejecuta el proceso de MS-DOS
Dim Archivo_Bat As String       'Almacena el nombre del archivo bat
Dim Lista_Archivos As Object    'Almacena el listado de archivos
Dim Archivo As Object           'Almacena el archivo actual
Dim Ultimo_Archivo As Object    'Almacena el ultimo archivo
Dim Linea As String             'Almacena la lectura de la Linea
Dim No_Mensajes As Integer      'Almacena el numero de mensajes recibidos
Dim Xml_PreFactura As DOMDocument   'Almacena la prefactura o notificacion
Dim Encontro As Integer         'Almacena si encontro el nodo en alguna posición
Dim Mensaje_Notificacion As String  'Almacena el mensaje de notificacion de cada documento
Dim No_PreFacturas As Integer       'Almacena la cantidad de prefacturas
Dim No_Notificaciones As Integer    'Almacena la cantidad de notificaciones

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    
    'Asigna la ruta de ejecucion
    Ruta_Ejecucion = Mid(Ruta_Bimbo_Entrada, 1, Len(Ruta_Bimbo_Entrada) - 7)
    'Asigna las instrucciones de ejecucion
    Ejecucion = Ruta_Ejecucion & "SEDEB2BONLINE.EXE /SERVER:" & Servidor_Bimbo & " /PORT:" & Puerto_Bimbo & " /SSL /USER:" & Usuario_Bimbo & " /PWD:" & Password_Bimbo & " /GET:""" & Ruta_Bimbo_Entrada & "\"" /PARSEMIME"
    'Asigna el bat
    Archivo_Bat = "RECEPCION_HARINERA_" & Nombre_Usuario & ".bat"
    Open "C:\" & Archivo_Bat For Output As #1
    Print #1, Ejecucion
    Close #1
    'Ejecuta el proceso
    hProcess = OpenProcess(SYNCHRONIZE, 0, Shell("C:\" & Archivo_Bat, vbHide))
    'Indica si se termino de procesar la información para poder continuar con las siguientes
    'ejecuciones
    If hProcess Then
        WaitForSingleObject hProcess, INFINITE
        CloseHandle hProcess
    End If
    
    'Elimina el archivo bat
    'Valida que exista el archivo
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta("C:\" & Archivo_Bat, "ARCHIVO") = True Then
        'Elimina el archivo
        Kill "C:\" & Archivo_Bat
    End If
    
    'Copia los archivo a la prefactura
    Set Lista_Archivos = CreateObject("Scripting.FileSystemObject")
    For Each Archivo In Lista_Archivos.GetFolder(Ruta_Bimbo_Entrada).Files
        
        'Poner los datos en el analizador de XML
        Set Xml_PreFactura = New DOMDocument
        
        'Valida el xml
        Xml_PreFactura.resolveExternals = True
    
        'Para que valide el documento xml
        Xml_PreFactura.validateOnParse = True
    
        'Agiliza la carga del documento
        Xml_PreFactura.async = False
        
        'Carga el documento
        If Xml_PreFactura.Load(Trim(Archivo)) Then
            'Comprobamos si se carga
            If Xml_PreFactura.parseError.reason = "" Then
                'Valida si es una prefactura con el nombre del nodo
                If Xml_PreFactura.documentElement.nodeName = "PrefacturaBimbo" And Xml_PreFactura.lastChild.nodeName = "PrefacturaBimbo" Then
                    No_PreFacturas = No_PreFacturas + 1
                    'Copia el archivo a la ruta de prefacturas
                    FileCopy Archivo, Ruta_PreFacturas & "\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 3) & ".xml"
                Else
                    'Valida si es una prefactura con el nombre del nodo
                    If Xml_PreFactura.documentElement.nodeName = "AckErrorApplication" And Xml_PreFactura.lastChild.nodeName = "AckErrorApplication" Then
                        No_Notificaciones = No_Notificaciones + 1
                        'Copia el archivo a la ruta de notificaciones
                        FileCopy Archivo, Ruta_Ejecucion & "\Notificaciones\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 3) & ".xml"
                    End If
                End If
            End If
        Else
            'Abre el archivo como lectura
            AckErrorApplication = False
            Open Archivo For Input As #1
            If Not EOF(1) Then
                Do While Not EOF(1)
                    Line Input #1, Linea
                    'Reemplaza las comillas dobles
                    Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                    'Valida si es una archivo de notificacion
                    Encontro = InStr(1, Linea, "AckErrorApplication", 0)
                    If Encontro > 0 Then
                        AckErrorApplication = True
                    End If
                    'Si es un archivo de notificacion continua con el proceso
                    If AckErrorApplication = True Then
                        Encontro = InStr(1, Linea, "<ReferenceNumber type =  IV >", 0)
                        If Encontro > 0 Then
                            Line Input #1, Linea
                            'Reemplaza las comillas dobles
                            Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                            Encontro = InStr(1, Linea, ">", 0)
                            If Encontro > 0 Then
                                Linea = Mid(Linea, Encontro + 1)
                                Encontro = InStr(1, Linea, "<", 0)
                                If Encontro > 0 Then
                                    Linea = Mid(Linea, 1, Encontro - 1)
                                    If Trim(Mensaje_Notificacion) = "" Then
                                        Mensaje_Notificacion = "Factura --> " & Linea
                                    Else
                                        Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Factura --> " & Linea
                                    End If
                                    No_Notificaciones = No_Notificaciones + 1
                                End If
                            End If
                        End If
                        Encontro = InStr(1, Linea, "messageError sequence", 0)
                        If Encontro > 0 Then
                            Line Input #1, Linea
                            Line Input #1, Linea
                            Line Input #1, Linea
                            'Reemplaza las comillas dobles
                            Linea = Trim(Replace(Linea, """", " ", 1, , vbBinaryCompare))
                            Encontro = InStr(1, Linea, ">>", 0)
                            If Encontro > 0 Then
                                Linea = Mid(Linea, Encontro + 2)
                                Encontro = InStr(1, Linea, "<", 0)
                                If Encontro > 0 Then
                                    Linea = Mid(Linea, 1, Encontro - 1)
                                    Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Error --> " & Linea
                                End If
                            Else
                                Encontro = InStr(1, Linea, ">", 0)
                                If Encontro > 0 Then
                                    Linea = Mid(Linea, Encontro + 1)
                                    Encontro = InStr(1, Linea, "<", 0)
                                    If Encontro > 0 Then
                                        Linea = Mid(Linea, 1, Encontro - 1)
                                        Mensaje_Notificacion = Mensaje_Notificacion & Chr(13) & "Exito --> " & Linea
                                    End If
                                End If
                            End If
                        End If
                    End If
                Loop
            End If
            Close #1
            'Si no es archivo de notificacion lo elimina
            If AckErrorApplication = False Then GoTo 1
            'Copia el archivo a la ruta de notificaciones
            FileCopy Archivo, Ruta_Ejecucion & "Notificaciones\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 3) & ".xml"
        End If
        'Elimina el archivo
1:      Kill Archivo
        'Resetea la variable
        Set Xml_PreFactura = Nothing
    Next
    Set Archivo = Nothing
    
    'Revisa en el log que se haya enviado el archivo
    For Each Archivo In Lista_Archivos.GetFolder(Ruta_Ejecucion & "LOGS").Files
        If Ultimo_Archivo Is Nothing Then Set Ultimo_Archivo = Archivo
        'Valida si son archivos viejos los copia a su carpeta correspondiente
        If DateDiff("d", Now, Archivo.DateLastModified) < 0 Or Archivo.DateLastModified = Ultimo_Archivo.DateLastModified Then
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy"), "CARPETA") = False Then
                MkDir Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy")
            End If
            'Copia el archivo a la carpeta de respaldo
            FileCopy Archivo, Ruta_Ejecucion & "LOGS\" & Format(Archivo.DateLastModified, "ddMMMyyyy") & "\" & Archivo.Name
        End If
        'Verifica cual es el ultimo archivo
        If Archivo.DateLastModified > Ultimo_Archivo.DateLastModified Then
            Set Ultimo_Archivo = Archivo
        End If
    Next
    'Elimina los archivos antiguos
    For Each Archivo In Lista_Archivos.GetFolder(Ruta_Ejecucion & "LOGS").Files
        If Archivo <> Ultimo_Archivo Then
            Kill Archivo
        End If
    Next
    'Lee el archivo actual
    Open Ultimo_Archivo For Input As #1
    Do While Not EOF(1)
        'Lee linea por linea
        Line Input #1, Linea
        'Valida los diferentes mensajes
        If Trim(Linea) <> "" Then
            'Si no habia archivos que recibir
            If InStr(1, Linea, "Número de ficheros obtenidos:") > 0 Then
                If Val(Mid(Linea, InStr(1, Linea, "Número de ficheros obtenidos:", vbTextCompare) + 29)) = 0 Then
                    Mensaje = "No hay mensajes para recibir. Cantidad PreFacturas --> " & No_PreFacturas & "  Cantidad Notificaciones --> " & No_Notificaciones & Chr(13) & Mensaje_Notificacion
                    Recibe_Xmls_Bimbo = 0
                    Exit Do
                Else
                    No_Mensajes = Val(Mid(Linea, InStr(1, Linea, "Número de ficheros obtenidos:", vbTextCompare) + 29))
                End If
            Else
                'Si se enviaron los archivo
                If InStr(1, Linea, "Fin de procesar MIME") > 0 Then
                    Mensaje = "Archivos recibidos: " & No_Mensajes & ". Cantidad PreFacturas --> " & No_PreFacturas & "  Cantidad Notificaciones --> " & No_Notificaciones & Chr(13) & Mensaje_Notificacion
                    Recibe_Xmls_Bimbo = 1
                    Exit Do
                Else 'Si hubo error
                    If InStr(1, Linea, "Error:") > 0 Then
                        Mensaje = Trim(Mid(Linea, InStr(1, Linea, "Error:") + 1))
                        Recibe_Xmls_Bimbo = -1
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop
    Close #1
    'Elimina el archivo actual
    Kill Ultimo_Archivo
    'Reinicia las variables
    Set Archivo = Nothing
    Set Ultimo_Archivo = Nothing
    Set Lista_Archivos = Nothing

    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Function
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Mensaje = Err.Description
    Recibe_Xmls_Bimbo = -1
End Function


'*************************************************************************************
    'NOMBRE DE LA FUNCIÓN: Valida_Fechas
    'DESCRIPCIÓN: Valida que las fechas que proporciono el usuario sean validas para el
    '             sistema mandando un estatus de verdadero pero si no son validas
    '             entonces manda un valor de falso
    'PARÁMETROS :
    'CREO       : Yazmin A. Delgado Gómez
    'FECHA_CREO : 13-Diciembre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*************************************************************************************
Function Valida_Fechas(Fecha_Inicio As Date, Fecha_Final As Date) As Boolean
    If Year(Format(Fecha_Inicio, "yyyy/MM/dd")) < 1900 Or Year(Format(Fecha_Final, "yyyy/MM/dd")) > Year(Format(Now, "yyyy/MM/dd")) Then
        Valida_Fechas = False
    Else
        Valida_Fechas = True
    End If
End Function
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Selecciona_Ruta_Directorio
    'DESCRIPCIÓN: Asigna la ruta seleccionada a nivel directorio
    'PARÁMETROS : Frm, pasa la forma requerida
    '             Caption_Asignado, asigna el titulo a la ventana
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 29/Ago/2009 10:05am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Selecciona_Ruta_Directorio(Frm As Form, Caption_Asignado As String) As String
Dim Nulo As Integer
Dim Identificador As Long
Dim Ruta As String
Dim Directorios As BrowseInfo

    With Directorios
        .hWndOwner = Frm.hwnd    'Formulario
        .lpszTitle = lstrcat(Caption_Asignado, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Mostrar el cuadro de dialogo Buscar carpeta
    Identificador = SHBrowseForFolder(Directorios)
    If Identificador Then
        Ruta = String$(MAX_PATH, 0)
        'Llamamos a la APi y recuperamos el id del path y en _
         El_Path obtenemos el path seleccionado
        Call SHGetPathFromIDList(Identificador, Ruta)
        'Liberamos el bloque de memoria
        Call CoTaskMemFree(Identificador)
        'Busca la posición del primer caracter nulo
        Nulo = InStr(Ruta, vbNullChar)
        If Nulo Then
            'Formateamos la cadena anterior eliminado los espacios nulos del path
            Ruta = Left$(Ruta, Nulo - 1)
        End If
        Selecciona_Ruta_Directorio = Ruta
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: SHCopyFile
    'DESCRIPCIÓN: Copia un archivo de origen a destino
    'PARÁMETROS : from_file, ruta origen del archivo
    '             to_file, ruta destino del archivo
    'CREO       : Ismael Prieto Sánchez
    'FECHA_CREO : 29/Ago/2009 10:05am
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub SHCopyFile(ByVal from_file As String, ByVal to_file As String)
On Error GoTo HANDLER
Dim sh_op As SHFILEOPSTRUCT
    With sh_op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = from_file & vbNullChar & vbNullChar
        .pTo = to_file & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation sh_op
Exit Sub

HANDLER:
    MsgBox Er.Description
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Parametros
    'DESCRIPCIÓN: Consulta los parámetros que tiene el sistema asignados
    'PARÁMETROS :
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 15-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Parametros()
Dim Rs_Consulta_Cat_Parametros As rdoResultset 'Consulta los parámetros que tiene asignado el sistema

    'Consulta los parámetros del sistema
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Parametros.EOF Then
        With Rs_Consulta_Cat_Parametros
            Intentos_Fallidos = .rdoColumns("Intentos_Fallidos")
            Bloqueo_Por_No_Utilizar = Val(.rdoColumns("Vencimiento_Cuenta_Usuario"))
            Bloqueo_Por_Expiración_Password = Val(.rdoColumns("Limite_Cambio_Password"))
            If Not IsNull(.rdoColumns("Rol_ID_Administrador")) Then Rol_Administrador = .rdoColumns("Rol_ID_Administrador")
        End With
    End If
    Rs_Consulta_Cat_Parametros.Close
End Sub


