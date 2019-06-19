VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ayudante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private iv_ErrorMessage As String          'Propiedad de mensaje de error, Vacía si no existe error
Private iv_Form As Form                    'Forma a la que se va a redimencionar
Private iv_NoOfSizeableControls As Integer 'Cuenta la cantidad de controles que no se pueden redimencionar
Private iv_OrigHeight As Long              'Guarda el Alto original
Private iv_OrigWidth As Long               'Guarda el Ancho original
Private iv_OrigScaleHeight As Long         'Guarda el Scaleheight original
Public Servidor As String  'Guarda el nombre del servidor
Public Base As String      'Guarda el nombre de la base de datos
Public Tipo  As String     'Identifica en que esta hecha la BD Access o SQL
Public Password As String  'Guarda la contraseña para el acceso al servidor
Public Usuario_Conexion As String   'Guarda el nombre válido del usuario para el acceso a la BD
Public BaseD As String
Private Type CtlInfo                       'Tipo de Dato definido por el usuario
    CtlNum As Integer                      'Numero de Control
    T As Long                              'Manejo del propiedad Top de control
    L As Long                              'Manejo de la propiedad Left del control
    H As Long                              'Manejo de la propiedad Height del control
    W As Long                              'Manejo de la propiedad Width del control
    F As Long
End Type
Private iv_Ctl() As CtlInfo                'Variable de Usuario

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Valida_Existe_Archivo_Carpeta
    'DESCRIPCIÓN: Función para validar si existe un Archivo o Carpeta en la ruta especificada.
    'PARÁMETROS : Ruta_Archivo - cadena con la ruta del directorio a validar.
    '             Tipo - Indica si se validará una Carpeta o un Archivo.
    'CREO       : Jose Antonio Lopez Hernandez
    'FECHA_CREO : 17/Junio/2008
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Valida_Existe_Archivo_Carpeta(Ruta_Archivo As String, Tipo As String) As Boolean
Dim AttrDev%

    'Se valida si existe una carpeta en la ruta especificada.
    If Tipo = "CARPETA" Then
        If CBool(PathIsDirectory(Ruta_Archivo)) Then
            Valida_Existe_Archivo_Carpeta = True
        Else
            Valida_Existe_Archivo_Carpeta = False
        End If
    Else 'Si no es carpeta valida que sea un archivo y que exista.
        On Error Resume Next
        AttrDev = GetAttr(Ruta_Archivo)
        'Condicion para capturar un error y regersar un valor.
        'Si encuentra un error lo limpia y regersa un Valor false, eso es que no encontro el archivo.
        If Err.Number Then
            Err.Clear
            Valida_Existe_Archivo_Carpeta = False
        Else 'Si no encuentra el un error quiere decir que existe un archivo y regresa un valor True
            Valida_Existe_Archivo_Carpeta = True
        End If
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: GenRectangulo
    'DESCRIPCIÓN: De acuerdo a las coordenadas que el usuario le pasa a la funcion
    '             dibuja un rectangulo con esas dimensiones
    'PARÁMETROS:  X=El primer valor de donde va a empesar a dibujar el rectangulo
    '               con respecto a lo horizontal
    '             Y=El primero valor en donde va a empesara a dibujar el rectangulo
    '               conrespecto a lo alto
    '            dx=Hasta donde va a dibujar la linea horizontal
    '            dy=Lo Alto del rectangulo
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub GenRectangulo(X!, y!, dx!, dy!, Radius!, op%)
'    'se genera un rectángulo
'    Printer.DrawWidth = 2
'    Select Case op%
'        Case True 'esquinas redondeadas
'            'ajuste para el radio de las esquinas
'            X! = X! + Radius
'            Y! = Y! + Radius
'            dx! = dx! - 2 * Radius
'            dy! = dy! - 2 * Radius
'            'esquina izquierda superior
'            Printer.Circle (X!, Y!), Radius!, , PI / 2, PI
'            'Esquina derecha superior.
'            Printer.Circle (X! + dx!, Y!), Radius!, , 0, PI / 2
'            'Esquina izquierda inferior.
'             Printer.Circle (X!, Y! + dy!), Radius!, , PI, 3 * PI / 2
'            'Esquina derecha inferior.
'            Printer.Circle (X! + dx!, Y! + dy!), Radius!, , 3 * PI / 2, 2 * PI
'            'Ahora las lineas.
'            'Linea horizontal de arriba.
'            Printer.Line (X!, Y! - Radius!)-(X! + dx!, Y! - Radius!)
'            'Linea izquierda vertical.
'            Printer.Line (X! - Radius!, Y!)-(X! - Radius!, Y! + dy!)
'            'Linea horizontal de abajo.
'            Printer.Line (X!, Y! + dy! + Radius!)-(X! + dx!, Y! + dy! + Radius!)
'            'Linea derecha vertical.
'            Printer.Line (X! + dx! + Radius!, Y!)-(X! + dx! + Radius!, Y! + dy!)
'        Case False  'Rectangulo normal.
'            Printer.Line (X!, Y!)-(X! + dx!, Y! + dy!), , B
'    End Select
End Sub

'*******************************************************************************
    'NOMBRE DE LA PROPIEDAD: Forma
    'DESCRIPCIÓN: Ajusta al ayudante la Forma a utilizar en el redimensionado
    'PARÁMETROS:               1. av_Form La forma que se va a aplicar  el redimensionado
    'CREO: Ruben Garcia
    'FECHA_CREO:21 Nov 2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Property Set Forma(av_Form As Object)
    iv_ErrorMessage = ""            'inicializamos los errores a Ninguno
    On Error GoTo ErrorRoutine
    Dim lv_ControlNumber As Integer 'Lleva al cuenta del control
    Dim lv_Ctl As Control           'Emula al control actual
    Dim lv_Height As Long           'Toma el Alto del  control
    Dim I As Integer                'Contador
    Set iv_Form = av_Form           'variable que emula la forma afectada
    '-----------------------------------------------
    'guardamos el Alto y Ancho originales de la forma.
    '-----------------------------------------------
    iv_OrigHeight = iv_Form.Height
    iv_OrigWidth = iv_Form.Width
    iv_OrigScaleHeight = iv_Form.ScaleHeight
    '-------------------------------------------------------
    'Examinamos la coleccion de controles, determinando el numero
    'de controles que tienen la propiedad Height. Controles como
    'los menus no tienen tal propiedad.
    '-------------------------------------------------------
    iv_NoOfSizeableControls = -1
    On Error Resume Next
    For I = 0 To (iv_Form.Controls.Count - 1)
        lv_Height = iv_Form.Controls(I).Height
        If Err = 0 Then
            iv_NoOfSizeableControls = iv_NoOfSizeableControls + 1
        Else
            Err = 0
        End If
    Next
    On Error GoTo ErrorRoutine
    
    If iv_NoOfSizeableControls = -1 Then
        Exit Property
    End If
    '-----------------------------------------------
    'Redimensionamos el arreglo de Controles igual al tamaño
    'de la cantidad de estos que si tienen la propiedad Height
    '-----------------------------------------------
    ReDim iv_Ctl(iv_NoOfSizeableControls)
    '------------------------------------------------------
    'Alamcenamos el numero de colecciones y las dimensiones originales
    'de cada control que tiene la propiedad Height
    '------------------------------------------------------
    I = -1
    For lv_ControlNumber = 0 To (iv_Form.Controls.Count - 1)
        Set lv_Ctl = iv_Form.Controls(lv_ControlNumber)
        With lv_Ctl
            Debug.Print lv_Ctl.Name
            'MsgBox lv_Ctl.Name
            On Error Resume Next
            lv_Height = .Height
            If Err = 0 Then  'El control si tiene la propiedad Height
                I = I + 1
                If Not (I > iv_NoOfSizeableControls) Then 'para que sea seguro
                    iv_Ctl(I).CtlNum = lv_ControlNumber
                    iv_Ctl(I).T = .Top
                    iv_Ctl(I).L = .Left
                    iv_Ctl(I).H = .Height
                    iv_Ctl(I).W = .Width
                    iv_Ctl(I).F = .Font.Size
                End If
            End If
            On Error GoTo ErrorRoutine
        End With
    Next
    Exit Property
ErrorRoutine: 'Guardamos el error que ocurrio para tomarlo de la propiedad Mensaje_Error
    iv_ErrorMessage = _
        "Eror en " & App.EXEName & ".clsControlResizer.Form.PropertySet" & _
        Chr(13) & Chr(13) & "Error" & Str(Err) & " " & error
    Exit Property
End Property

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Redimensionar_Controles
    'DESCRIPCIÓN: Redimesiona los controles de la Forma contenida en la propiedad
    'Forma.
    'USO: '         1.  En la seccion de declaraciones de la forma:
    '                     Dim Resizer as Ayudante
    '               2.  En el evento Form_Load , escribe:
    '                     Set Resizer = new ayudante
    '                     Set Resizer.Form = Me
    '                     If Len(Resizer.Mensaje_Error) > 0 then
    '                         MsgBox Resizer.Mensaje_Error
    '                     End If
    '               3.  En el evento  Form_Resize event, escribe:
    '                     Resizer.RedimensionarControles
    '                     If Len(Resizer.Mensaje_Error) > 0 then
    '                         MsgBox Resizer.Mensaje_Error
    '                     End If
    '               4.  Antes de terminar la ejecucion de tu forma, escribe:
    '                     Set Resizer = Nothing
    'PARÁMETROS:
    'CREO: Ruben Garcia
    'FECHA_CREO: 21-Nov-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
    'NOTA: No funciona bien con el contro SSTab
    '      Si se usa este control se debera manejar en su evento Clic la aparicion
    '      Y posicion de los controles contenidos en el
'*******************************************************************************
Public Sub Redimensionar_Controles()
    On Error GoTo ErrorRoutine
    Dim lv_Ctl As Control                   'Emula el control actual
    Dim lv_Tag As String                    'Almacena la cadena destinada a cubrir el Tag
    Dim lv_TagPos As Integer                'Guarda la posicion del Tag que se va editando
    Dim lv_TagChar As String
    Dim lv_DoNotResize As Boolean           'indica si algun control no se va a redimensionar
    Dim lv_ResizeTop As Boolean
    Dim lv_ResizeLeft As Boolean
    Dim lv_ResizeHeight As Boolean
    Dim lv_ResizeWidth As Boolean
    Dim lv_ResizeFont As Boolean
    Dim lv_ResizeToKeepBottomGap As Boolean 'indica si se va respetar el margen de la forma
    Dim lv_NewT As Long                     'Almacen ala nueva posicion de altura del control
    Dim lv_NewL As Long                     'Almacena la nuev aposicion izquierda del control
    Dim lv_NewH As Long                     'Almacena la nueva altura del control
    Dim lv_NewW As Long                     'Almacena la Nueva Anchura del control
    Dim lv_NewF As Long
    Dim lv_HFactor As Single                'guarda el Factor de altura para el redimensionado
    Dim lv_WFactor As Single                'guarda el Factor de anchura para el redimensionado
    Dim I As Integer                        'Iterador de bucle
    '-----------------------------------------------------
    'Salida del procedimiento si no hay controles que s epuedan redimensionar.
    '-----------------------------------------------------
    If iv_NoOfSizeableControls = -1 Then
        Exit Sub
    End If
    '---------------------------------------------------
    'Se calcula el factor horizontal y vertical de redimensionado.
    '---------------------------------------------------
    lv_HFactor = iv_Form.Height / iv_OrigHeight
    lv_WFactor = iv_Form.Width / iv_OrigWidth
    '----------------------------------------------
    'recorrido por todos los controles,se redmensiona si es necesario
    '----------------------------------------------
    For I = LBound(iv_Ctl) To UBound(iv_Ctl)
        Set lv_Ctl = iv_Form.Controls(iv_Ctl(I).CtlNum)
        '----------------------------------------------------------
        'Para cambiar el comportamiento habitual del redimensionado, el tag de un control
        'puede contener un cadena como "Resize=TLW" el cual en este caso
        'especifica que el ajuste de tamaño aplica a la propiedaa Top, Left y Width
        'pero no para height.El codigo siguinte se encarga de ello
        '----------------------------------------------------------
        lv_DoNotResize = False
        lv_ResizeTop = True
        lv_ResizeLeft = True
        lv_ResizeHeight = True
        lv_ResizeWidth = True
        lv_ResizeFont = True
        lv_ResizeToKeepBottomGap = False
        lv_Tag = lv_Ctl.Tag
        lv_TagPos = InStr(lv_Tag, "Resize=")
        If lv_TagPos <> 0 Then
            lv_TagPos = lv_TagPos + 6
            lv_ResizeTop = False
            lv_ResizeLeft = False
            lv_ResizeHeight = False
            lv_ResizeWidth = False
            lv_ResizeFont = False
            Do
                lv_TagPos = lv_TagPos + 1
                If lv_TagPos > Len(lv_Tag) Then
                    Exit Do
                End If
                lv_TagChar = Mid(lv_Tag, lv_TagPos, 1)
                If lv_TagChar = Space(1) Then
                    Exit Do
                End If
                Select Case lv_TagChar
                    Case "N"
                        lv_DoNotResize = True
                    Case "T"
                        lv_ResizeTop = True
                    Case "L"
                        lv_ResizeLeft = True
                    Case "H"
                        lv_ResizeHeight = True
                    Case "W"
                        lv_ResizeWidth = True
                    Case "G"
                        lv_ResizeToKeepBottomGap = True
                    Case "F"
                End Select
            Loop
        End If
        If lv_DoNotResize Then
            'no hace nada
        Else
            '-----------------------------------------
            'Calcular las nuevas dimensiones paar el control
            '-----------------------------------------
            With iv_Ctl(I)
                lv_NewT = .T * IIf(lv_ResizeTop, lv_HFactor, 1)
                lv_NewL = .L * IIf(lv_ResizeLeft, lv_WFactor, 1)
                lv_NewW = .W * IIf(lv_ResizeWidth, lv_WFactor, 1)
                If lv_WFactor < lv_HFactor Then
                    lv_NewF = .F * IIf(lv_ResizeWidth, lv_WFactor, 1)
                Else
                    lv_NewF = .F * IIf(lv_ResizeWidth, lv_HFactor, 1)
                End If
                If lv_ResizeToKeepBottomGap Then
                   '-------------------------------------------------
                   'Ajusta el alto para mantener una distancia constante
                   'desde el bajo de la forma
                   '-------------------------------------------------
                   lv_NewH = _
                    (iv_Form.ScaleHeight - lv_NewT) - (iv_OrigScaleHeight - (.T + .H))
                Else
                   '------------------------------------------------
                   'calcular la nueva altura basado en el gfactor de altura
                   'Calculate new height based on new height factor.
                   '------------------------------------------------
                   lv_NewH = .H * IIf(lv_ResizeHeight, lv_HFactor, 1)
                End If
            End With
            '------------------------------
            'Position y tamaño del control
            '------------------------------
            On Error Resume Next
            lv_Ctl.Move lv_NewL, lv_NewT
            lv_Ctl.Width = lv_NewW
            lv_Ctl.Height = lv_NewH
            lv_Ctl.Font.Size = lv_NewF
            On Error GoTo ErrorRoutine
            
        End If 'No redimensiona
    Next I  'procesar el siguiente control
    Exit Sub
ErrorRoutine:
    iv_ErrorMessage = _
        "Error en " & App.EXEName & ".clsControlResizer.ResizeControls" & _
        Chr(13) & Chr(13) & "Error" & Str(Err) & " " & error
    Exit Sub
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alinea_Derecha
    'DESCRIPCIÓN: Alinea_Derecha los número a la izquierda del documento
    'PARÁMETROS:
    '             1. Numero:
    '             2. Longitud:
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Alinea_Derecha_Derecha(Numero As String, Longitud As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim I As Integer     'Cuenta el numero de caracteres de la cadena
    Nuevo = Numero
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For I = 1 To Longitud - Len(Numero)
        Nuevo = " " & Nuevo
    Next I
    Alinea_Derecha_Derecha = Nuevo
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Cargar_Picture
    'DESCRIPCIÓN: Este proceso hace visible la picture que necesitamos ocultando
    '             todo las demas a la vez ajusta el tamaño de la forma
    'PARÁMETROS:
    '             1. Picture: Nombre del Picture el cual vamos hacer visible
    '             2. Formulario: Nombre del Formulario al cual nos estamos refiriendo
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Cargar_Picture(picture As PictureBox, Formulario As Form)
Dim picOcultar As PictureBox    'Ocultar los PictureBox
Dim Control As Control          'Toma la forma del objeto al que esta apuntando en ese momento
    picture.BorderStyle = 0
    'Oculta los pictures contenidos en la forma
    For Each Control In Formulario.Controls
        If TypeOf Control Is PictureBox Then
            Control.Visible = False
        End If
    Next
    picture.Visible = True
    'Se Alinea_Derecha el formulario al centro de la forma
    If Formulario.WindowState <> 2 Then
        'Formulario.Left = (Screen.Width - Formulario.Width) \ 2
        'Formulario.Top = (Screen.Height - Formulario.Height) \ 2
        Formulario.Left = (MDIFrm_Apl_Principal.Width - Formulario.Width) \ 2
        Formulario.Top = 0
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Conexion
    'DESCRIPCIÓN: Se hace la conexión a la Base de datos verificando si es en SQL
    '             o Acces de acuerdo a los datos proporcionados por el Usuario
    'PARÁMETROS:
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Conexion()
Dim I As Integer        'Cuenta cada una de las lineas del documento Confing.ini
Dim linea As String     'Guarda lo que contenga la linea a la cual se esta apuntando

    
On Error GoTo Etiqueta
    I = 0
    'LEE EL ARCHIVO DE INICIO DEL SISTEMA
    Open App.Path & "\Config.ini" For Input As #1
        Do While Not EOF(1)
            Line Input #1, linea
            If I = 0 Then Empresa = Trim(linea)
            If I = 1 Then RFC = Trim(linea)
            If I = 2 Then Direccion = Trim(linea)
            If I = 3 Then Direccion = Direccion & " " & Trim(linea)
            If I = 4 Then CP = Trim(linea)
            If I = 5 Then Telefono = Trim(linea)
            If I = 6 Then Ciudad_Edo = Trim(linea)
            If I = 7 Then Ciudad_Edo = Ciudad_Edo & " " & Trim(linea)
            If I = 8 Then Tipo = Trim(linea)
            If I = 9 Then Servidor = Trim(linea)
            If I = 10 Then Usuario_Conexion = Trim(linea)
            If I = 11 Then Base = Trim(linea)
            If I = 12 Then Password = Trim(linea)
            I = I + 1
        Loop
    Close #1
    BaseD = Base
    Conexion_Base.CursorDriver = rdUseOdbc
    '1. Verifica si la base de datos es en Access o SQL
    '2. Establece la conexión a la base
    If Tipo = "ACCESS" Then
        Conexion_Base.Connect = "Driver={Microsoft Access Driver (*.mdb)};" & _
                                "Dbq=" & App.Path & "\" & Base & ".MDB;" & _
                                "Uid=;Pwd="
        Par_Fecha = "#"
        Base_Datos = "A"
    Else
        Conexion_Base.Connect = "uid=" & Usuario_Conexion & ";pwd=" & Password & ";server=" & Servidor & ";" _
                                & "driver={SQL Server};protocol=6.3;readonly=false;database=" & Base & ";dsn=;"
        Par_Fecha = "'"
        Server = Servidor
        Database = Base
        User_Password = Password
        User_Conexion = Usuario_Conexion
        Base_Datos = "S"
    End If
    Conexion_Base.EstablishConnection
    Conexion_Base.QueryTimeout = 80
    Load Frm_Apl_Login
    Exit Sub
Etiqueta:
    'Load Frm_Apl_Registro_Sistema
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Conexion_Respaldo
'DESCRIPCION: Realiza una conexión alterna a la base de datos para obtener un
'             respaldo del sistema
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 27-Julio-2011
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Conexion_Respaldo()
    Conexion_Base_Respaldo.CursorDriver = rdUseServer
    Conexion_Base_Respaldo.Connect = "uid=" & User_Conexion & ";pwd=" & User_Password & ";server=" & Server & ";" _
        & "driver={SQL Server};protocol=6.3;readonly=false;database=" & Database & ";dsn=;"
    Conexion_Base_Respaldo.EstablishConnection
    Conexion_Base_Respaldo.QueryTimeout = 1000
    Exit Sub
Etiqueta:
    Conexion_Base_Respaldo.Close
    MsgBox "No se pudo establecer la conexión con la base de datos", vbCritical
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Conexion_Vista
'DESCRIPCION: Realiza una conexión alterna a la base de datos para obtener el
'             acceso a la fuente información
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 19-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Public Sub Conexion_Servidor_Vista()
Dim iLinea As Integer
Dim I As Integer        'Cuenta cada una de las lineas del documento Confing.ini
Dim linea As String     'Guarda lo que contenga la linea a la cual se esta apuntando
Dim Servidor As String  'Guarda el nombre del servidor
Dim Base As String      'Guarda el nombre de la base de datos
Dim Usuario_Conexion As String
Dim Password As String  'Guarda la contraseña para el acceso al servidor

On Error GoTo Etiqueta
    I = 0
    'LEE EL ARCHIVO DE INICIO DEL SISTEMA
    iLinea = 1
    Open App.Path & "\Config_Vista.ini" For Input As #1
        Do While Not EOF(1)
            Line Input #1, linea
            If I = 0 Then Servidor = Trim(linea)
            If I = 1 Then Base = Trim(linea)
            If I = 2 Then Usuario_Conexion = Trim(linea)
            If I = 3 Then Password = Trim(linea)
            I = I + 1
        Loop
    Close #1
    'Establece la conexión a la base de Matriz
    iLinea = 2
    Conexion_Servidor.CursorDriver = rdUseOdbc
    iLinea = 3
    Conexion_Servidor.Connect = "uid=" & Usuario_Conexion & ";pwd=" & Password & ";server=" & Servidor & ";" & "driver={SQL Server};protocol=6.3;readonly=false;database=" & Base & ";dsn=;"
    iLinea = 4
    Conexion_Servidor.EstablishConnection
    iLinea = 5
    Conexion_Servidor.QueryTimeout = 1000
Exit Sub
Etiqueta:
    Conexion_Servidor.Close
    MsgBox "[Ayudante|Conexion_Servidor_Vista|Linea:" & iLinea & "] No se pudo establecer la conexión con la base de datos", vbCritical
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Cantidad_Letras
    'DESCRIPCIÓN: Convierte de cantidad de número a letras
    'PARÁMETROS:
    '             1. CantAConv: Es el número al convertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Convierte_Cantidad_Letras(CantAConv As String)
Dim I As Integer
Dim Longitud As Integer   'Obtiene la longitud de la cadena (número)
Dim PointPos As Integer
Dim Analisis As Integer
Dim Pesos As String
Dim Letras As String      'Asigna el nombre al número
Dim Centavos As String
Dim TempLetras As String
Dim Uno As Boolean

    Centavos = ""
    Pesos = ""
    Letras = ""
    Uno = False
    'Conversion  de cantidad de numero a letras
    Longitud = Len(CantAConv)
    PointPos = InStr(1, CantAConv, ".")
    If (Longitud > 0) And (PointPos <= 13) And Convierte_Valida_Cantidad(CantAConv) Then
        If PointPos > 0 Then
            Centavos = Mid(CantAConv, PointPos + 1, 2)
            Pesos = Mid(CantAConv, 1, PointPos - 1)
            Longitud = PointPos - 1
        Else
            Centavos = "00"
            Pesos = CantAConv
        End If
        I = 1
        If Longitud >= 10 Then            ' Miles de Millones }
            Select Case Longitud
                Case 10                   ' Miles de Millones }
                    If Mid(Pesos, I, 1) <> "1" Then
                        TempLetras = Mid(Pesos, I, 1)
                        Letras = Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                        I = I + 1
                    End If
                Case 11                   ' Decenas de Miles de Millones }
                    Letras = Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
    
                Case 12                   ' Centenas de Miles de Millones }
                    Letras = Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
            End Select
            Letras = Letras & " MIL"      ' Agrega la Terminaci¢n MIL }
            Longitud = Longitud - I + 1   ' Actualiza para continuar an lisis }
        End If
        If Longitud >= 7 Then             ' Millones }
            If Letras <> "" Then          ' Hay miles de millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 7)
                    I = I + 1
                    Longitud = Longitud - 1
                Loop
            End If
            Select Case Longitud           ' Millones }
                Case 7
                    If (Letras = "") And (Mid(Pesos, I, 1) = "1") Then Uno = True
                    Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    I = I + 1
                    Longitud = Longitud - 1     ' Actualiza Para continuar an lisis }
                Case 8                          ' Decenas de Millones }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                    Longitud = Longitud - 2
                Case 9                          ' Centenas de Millones }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
                    Longitud = Longitud - 3     ' Actualiza Para continuar an lisis }
            End Select
            If Uno Then
                    ' Si no hay miles de millones y es 1 entonces la
                    ' Terminación debe de ser MILLON y no MILLONES }
                Letras = Letras & " MILLON"
                Uno = False
            Else
                Letras = Letras & " MILLONES"   ' Agrega la terminaci¢n MILLONES }
            End If
        End If
        If Longitud >= 4 Then             ' Miles }
            If Letras <> "" Then          ' Hay millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 1)
                    I = I + 1
                    Longitud = Longitud - 1
                Loop
                If Longitud = 0 Then Uno = True
            End If
            Select Case Longitud           ' Miles
                Case 4
                    If (Mid(Pesos, I, 1) <> "1") Then
                        Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    End If
                    I = I + 1
                    Analisis = 1
                Case 5                      ' Decenas de Miles }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                    Analisis = 2
                Case 6                      ' Centenas de Miles }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
                    Analisis = 3
            End Select
            If Uno Then
                        ' Si no hay miles ni cientos pero hubo millones entonces
                        ' se debe de agregar la preposici¢n de Antes de PESOS }
                Letras = Letras & " DE"
                Uno = False
            Else
                If Longitud >= 4 Then
                    Letras = Letras & " MIL"    ' Agrega la terminaci¢n MIL }
                   End If
                If Longitud >= 4 Then Longitud = Longitud - Analisis
                                                ' Actualiza Para continuar an lisis }
            End If
        End If
        If Longitud >= 1 Then                   ' Cientos }
            If Letras <> "" Then                ' Hay Miles o Millones, por lo tanto elimina los ceros }
                Do While (Mid(Pesos, I, 1) = "0") And (Longitud >= 1)
                        I = I + 1
                        Longitud = Longitud - 1
                Loop
            End If
            Select Case Longitud                 ' Unidades }
                Case 1
                    If (Letras = "") And (Mid(Pesos, I, 1) = "1") Then Uno = True
                    Letras = Letras & " " & Convierte_Unidades_Letras(Mid(Pesos, I, 1))
                    I = I + 1
                Case 2                           ' Decenas }
                    Letras = Letras & " " & Convierte_Decenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1))
                    I = I + 2
                Case 3                           ' Centenas }
                    Letras = Letras & " " & Convierte_Centenas_Letras(Mid(Pesos, I, 1), Mid(Pesos, I + 1, 1), Mid(Pesos, I + 2, 1))
                    I = I + 3
            End Select
            Longitud = 0
        End If
        If Uno Then
            ' La cantidad es un Peso, ya que esta bandera ( Uno ) se prendió
            ' Cuando no Hubo Millones, ni miles y la cantidad era uno, por lo que la terminación es PESO }
            Letras = Letras & " PESO"
        Else
            Letras = Letras & " PESOS"
        End If
        Letras = Letras & " " & Centavos & "/100 M.N."
    End If
    Convierte_Cantidad_Letras = Letras
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Conectar_Ayudante.Quitar_Caracter
    'DESCRIPCIÓN: Quita los caracteres raros de los textos para poder ser guardados
    '             en la base de datos
    'PARÁMETROS:
    '             1. Cadena1: Texto al que se le van a quitar los caracteres raros
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:6/JUN/05
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Quitar_Caracter(Texto As String, Caracter As String) As String
Dim Posicion As Integer
    Posicion = 1
    While Posicion <> 0
        Posicion = InStr(1, Texto, Caracter, vbTextCompare)
        If Posicion <> 0 Then
            Texto = Mid(Texto, 1, Posicion - 1) & Mid(Texto, Posicion + 1, Len(Texto))
        End If
    Wend
    Quitar_Caracter = Texto
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Solo_Letras
    'DESCRIPCIÓN: 'Solo acepta letras, espacios y puntos
    'PARÁMETROS:
    '             1. KeyAscii: Regresa el valor de la letra
    'CREO: Yazmin Abigail Delgado Gómez
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Solo_Letras(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'Acepta las letras de la A a la Z, incluyendo Ñ o ñ, espacios  y Puntos
    If (Chr(KeyAscii) < "A" Or Chr(KeyAscii) > "Z") Or KeyAscii = 46 Then
        If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 209 Then
            Exit Function
        End If
        KeyAscii = 0
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Solo_Numeros
    'DESCRIPCIÓN: 'Solo acepta números y el punto
    'PARÁMETROS:
    '             1. KeyAscii: Regresa el valor del numero
    '             2. Cadena: Contiene la cadena a validar
    '             3. Permitir_Punto: indica si se debe permitir escribir o no un punto en la cadena
    'CREO: Yazmin Abigail Delgado Gómez
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO: Jose Antonio Lopez Hernandez
    'CAUSA_MODIFICACIÓN: Añadir validacion del punto decimal quitando la variable global "Punto"
'*******************************************************************************
Public Function Solo_Numeros(KeyAscii As Integer, Cadena As String, Permitir_Punto As Boolean)
    'Solo acepta numeros, el espacio y un punto decimal
    If (Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9) Or KeyAscii = 46 Then
        'Valida que solo haya un punto en la cadena
        If KeyAscii = 46 Then
            'Si se permite escribir un punto
            If Permitir_Punto = True Then
                'Busca el punto en la cadena (1-Existe/0-No Existe)
                If InStr(1, Cadena, ".", vbBinaryCompare) = 0 Then
                    Exit Function
                End If
            End If
        Else 'Si no es el punto
            If KeyAscii = 8 Then
                Exit Function
            End If
        End If
        KeyAscii = 0
    End If
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Es_Alfanumerico
'DESCRIPCION: Indica si el texto tiene numeros y letras
'PARÁMETROS : Texto: Texto a compara para saber si tiene letras y numeros
'CREO       : Yazmin Delgado Gómez
'FECHA_CREO : 15-Octubre-2007
'MODIFICO   : Antonio Salvador Benavides Guardado
'FECHA_MODIFICO: 14/junio/2008
'CAUSA_MODIFICO: Estandarización
'*******************************************************************************
Public Function Es_Alfanumerico(Texto As String) As Boolean
Dim Temp_Texto As String     'Obtiene el texto que se desea validar
Dim Long_Texto As Integer     'Obtiene la longitud del texto que se esta validando
Dim Pos_Caracter As Integer     'Indica la posición del caracter que se esta consultando dentro de la cadena
Dim Caracter As String     'Obtiene el caracter que se esta comparando del texto que se desea validar
Dim Hay_Numeros As Boolean   'Indica si se encontraron numeros
Dim Hay_Caracteres As Boolean 'Indica si se encontraron letras
    
    Hay_Numeros = False
    Hay_Caracteres = False
    Temp_Texto = Texto
    Long_Texto = Len(Temp_Texto)
    If Long_Texto > 0 Then
        For Pos_Caracter = 1 To Long_Texto
            Caracter = Mid(Temp_Texto, Pos_Caracter, 1)
            If Caracter Like "[0-9]" Then Hay_Numeros = True
            If Caracter Like "[A-Za-z]" Then Hay_Caracteres = True
        Next
        If Hay_Caracteres = True And Hay_Numeros = True Then
            Es_Alfanumerico = True
        Else
            Es_Alfanumerico = False
        End If
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Maximo_Catalogo
    'DESCRIPCIÓN: Obtine el número del último registro existente en la tabla
    '             especificada
    'PARÁMETROS:
    '             1. Tabla: Nombre de la tabla de la Base de datos
    '             2. Campo: Llave primaria de la tabla
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Maximo_Catalogo(Tabla As String, Campo As String) As Double
Dim Rs_Tabla As rdoResultset
Dim Mi_SQL As New rdoQuery      'Obtiene el valor máximo del registro
    'Consulta el último registro de la tabla
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        '.SQL = "SELECT MAX(" & Campo & ") FROM " & Tabla & ""
        .SQL = "SELECT TOP 1 " & Campo & " FROM " & Tabla & ""
        .SQL = .SQL & " ORDER BY 1 DESC"
        .LockType = rdConcurReadOnly
        .RowsetSize = 1
        .CursorType = rdUseOdbc
        Set Rs_Tabla = .OpenResultset(rdConcurReadOnly)
    End With
    'Obtiene el último registro de la tabla a consultar
    If Not Rs_Tabla.EOF Then
        If Not IsNull(Rs_Tabla(0)) Then
            Maximo_Catalogo = Val(Rs_Tabla(0)) + 1
        Else
            Maximo_Catalogo = 1
        End If
    Else
        Maximo_Catalogo = 1
    End If
    Rs_Tabla.Close
    Mi_SQL.Close
End Function

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Tipo_Cambio
'DESCRIPCION: Obtine el tipo de cambio de acuerdo a la fecha y moneda seleccionada
'PARAMETROS : 1. Fecha_Cambio : Fecha en que se solicita el tipo de cambio
'             2. Moneda_Cambio: Moneda del tipo de cambio
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 15-Septiembre-2010
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICACION:
'*******************************************************************************
Public Function Consulta_Tipo_Cambio(Fecha_Cambio As String, Moneda_Cambio As String) As Double
Dim Rs_Consulta_Tipo_Cambio As rdoResultset

    If Moneda_Cambio = "PESOS" Then
        Fecha_Cambio = Fecha_Cambio
    Else
        'Si es dólares consulta el tipo de cambio de la fecha anterior a la seleccionada
        If Moneda_Cambio = "DOLARES" Then
            'Valida si la fecha es lunes para tomar el dato del viernes, para los demás casos se considera el día anterior
            If UCase(Format(Fecha_Cambio, "dddd")) = "LUNES" Or UCase(Format(Fecha_Cambio, "dddd")) = "MONDAY" Then
                Fecha_Cambio = DateAdd("d", -3, Fecha_Cambio)
            Else
                Fecha_Cambio = DateAdd("d", -1, Fecha_Cambio)
            End If
        Else
            If Moneda_Cambio = "EUROS" Or Moneda_Cambio = "CADS" Then
                Fecha_Cambio = DateAdd("M", -1, Fecha_Cambio)
            End If
        End If
    End If
    'Consulta el tipo de cambio de acuerdo a la fecha y moneda
    Mi_SQL = "SELECT Tipo_Cambio_ID,Tipo_Cambio FROM Cat_Tipos_Cambio"
    If Moneda_Cambio = "PESOS" Then
        Mi_SQL = Mi_SQL & " WHERE Moneda='" & Moneda_Cambio & "'"
    Else
        If Moneda_Cambio = "DOLARES" Then
            Mi_SQL = Mi_SQL & " WHERE Moneda='" & Moneda_Cambio & "'"
            Mi_SQL = Mi_SQL & " AND Fecha='" & Fecha_Cambio & "'"
        Else
            If Moneda_Cambio = "EUROS" Or Moneda_Cambio = "CADS" Then
                Mi_SQL = Mi_SQL & " WHERE Moneda='" & Moneda_Cambio & "'"
                Mi_SQL = Mi_SQL & " AND MONTH(Fecha)=" & Month(Fecha_Cambio)
            Else
               Mi_SQL = Mi_SQL & " WHERE Fecha='" & Fecha_Cambio & "'"
            End If
        End If
    End If
    Set Rs_Consulta_Tipo_Cambio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Obtiene el último registro de la tabla a consultar
    If Not Rs_Consulta_Tipo_Cambio.EOF Then
        Consulta_Tipo_Cambio = Val(Rs_Consulta_Tipo_Cambio.rdoColumns("Tipo_Cambio"))
    Else
        Consulta_Tipo_Cambio = 1
    End If
    Rs_Consulta_Tipo_Cambio.Close
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo
    'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
    'PARÁMETROS:
    '             1. Campos: Campo a consultar para el llenado del ComboBox
    '             2. Tabla: Nombre de la tabla a consultar
    '             3. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             4. Tipo: Para saber si esta consultando
    '             5. Campo_Con: Para consultar y llenar el campo con las palabras introducidas
    '                           por el usuario
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Llena_Combo(Campos As String, Tabla As String, Combo_Control As ComboBox, Tipo As Integer, Campo_con As String)
Dim Grupo() As String           'Guarda lo contenido en el campo
Dim Cadena As String            'Asigna el valor del campo en el ComboBox
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & Campo_con & " LIKE '%" & Combo_Control.Text & "%'"
        End If
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    If Not Rs_Combo.EOF Then
        Combo_Control.Clear
        Grupo() = Split(Campos, ",")
        campos_cont = UBound(Grupo()) + 1
        While Not Rs_Combo.EOF
            Cadena = Rs_Combo(0)
            For I = 1 To campos_cont - 1
                Cadena = Cadena & "  " & Rs_Combo(I)
            Next I
            Combo_Control.AddItem Cadena
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recordset_Eliminar
    'DESCRIPCIÓN: Crea un recordset para eliminar el registro y lo devuelve
    '             para su uso posterior, basandose en la consulta SQl
    'PARÁMETROS:
    '               1. Mi_SQL: Sentencia del registro a eliminar
    'CREO: Yazmin A Delgado Gómez
    'FECHA_CREO: 06-Junio-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Eliminar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurRowVer
        .CursorType = rdUseOdbc
        .RowsetSize = 1
        Set Recordset_Eliminar = .OpenResultset(rdOpenKeyset)
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recordset_Editar
    'DESCRIPCIÓN: Crea un recordset para editar los registros y lo devuelve
    '             para su uso posterior, basandose en la consulta SQl
    'PARÁMETROS:
    '               1. Mi_SQl: Sentencia del registro a modificar
    'CREO: Yazmin A Delgado Gómez
    'FECHA_CREO: 06-Junio-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Editar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurRowVer
        .CursorType = rdUseOdbc
        .RowsetSize = 1
        Set Recordset_Editar = .OpenResultset(rdOpenKeyset)
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recordset_Consultar
    'DESCRIPCIÓN: Crea un recordset para consultar de solo lectura y lo devuelve
    '             para su uso posterior, basandose en la consulta SQl
    'PARÁMETROS:
    '               1. Mi_SQL: Sentencia a consultar
    'CREO: Yazmin A Delgado Gómez
    'FECHA_CREO: 06-Junio-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Consultar(Mi_SQL As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = Mi_SQL
        .LockType = rdConcurReadOnly
        .CursorType = rdUseOdbc
        .RowsetSize = 20
        Set Recordset_Consultar = .OpenResultset(rdOpenStatic)
    End With
End Function

Public Sub Actualiza_Saldo(Fecha_S As String, Cuenta_S As String)
Dim Rs_Maximo_Movimientos As rdoResultset
Dim Fecha As String
Dim Cuenta, Mi_SQL As String
Dim Bandera As Integer
Dim Consecutivo As Double
Dim Rs_Actualiza_Movimientos As rdoResultset
Dim Rs_Saldo_Movimientos As rdoResultset
Dim Fecha_Saldo As String
Dim Banco As String
Dim Saldo As Double

On Error GoTo MuestraError
    'Obtiene el ultimo movimiento de acuerdo a la fecha
    Fecha = Format(DateAdd("d", -1, Fecha_S), "MM/dd/yyyy")
    Cuenta = Cuenta_S
    Bandera = 0
    Saldo = 0
    'Mi_SQL = "SELECT MAX(Fecha) FROM Adm_Movimientos"
    Mi_SQL = "SELECT TOP 1 Fecha FROM Adm_Movimientos"
    Mi_SQL = Mi_SQL & " WHERE Fecha<='" & Fecha & "'"
    Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
    Mi_SQL = Mi_SQL & " ORDER BY 1 DESC"
    Set Rs_Maximo_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Maximo_Movimientos.EOF Then
        If Not IsNull(Rs_Maximo_Movimientos.rdoColumns(0)) Then
            Fecha_Saldo = Format(Rs_Maximo_Movimientos.rdoColumns(0), "MM/dd/yyyy")
            Mi_SQL = "SELECT Saldo FROM Adm_Movimientos"
            Mi_SQL = Mi_SQL & " WHERE Fecha='" & Fecha_Saldo & "'"
            Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
            Mi_SQL = Mi_SQL & " ORDER BY Fecha ASC,No_Movimiento"
            Set Rs_Saldo_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Saldo_Movimientos.EOF And Not IsNull(Rs_Saldo_Movimientos.rdoColumns("Saldo")) Then
                Saldo = Rs_Saldo_Movimientos.rdoColumns("Saldo")
            End If
        Else
            Saldo = 0
            Bandera = 1
            Fecha_Saldo = Fecha
        End If
    Else
        Saldo = 0
        Bandera = 1
        Fecha_Saldo = Fecha
    End If
    Rs_Maximo_Movimientos.Close
    'Actualiza saldos
    Mi_SQL = "SELECT * FROM Adm_Movimientos "
    Mi_SQL = Mi_SQL & " WHERE Fecha>='" & Format(Fecha_Saldo, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Cuenta='" & Cuenta & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha,No_Movimiento"
    Set Rs_Actualiza_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Saldo = 0 And Bandera = 1 And Not Rs_Actualiza_Movimientos.EOF Then
        Saldo = Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
        If Rs_Actualiza_Movimientos.rdoColumns("Tipo") = "E" Then
            Saldo = Saldo * (-1)
        End If
    End If
    While Not Rs_Actualiza_Movimientos.EOF
        Rs_Actualiza_Movimientos.Edit
            Rs_Actualiza_Movimientos.rdoColumns("Saldo") = Saldo
        Rs_Actualiza_Movimientos.Update
        Rs_Actualiza_Movimientos.MoveNext
        If Not Rs_Actualiza_Movimientos.EOF Then
            If Rs_Actualiza_Movimientos.rdoColumns("Tipo") = "E" Then
                Saldo = Saldo - Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
            Else
                Saldo = Saldo + Rs_Actualiza_Movimientos.rdoColumns("Cantidad")
            End If
        End If
    Wend
    Rs_Actualiza_Movimientos.Close
    Exit Sub
MuestraError:
    MsgBox Err.Description
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Decenas_Letras
    'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
    'PARÁMETROS:
    '             1. Ciento: Es el número al convertir a letra
    '             2. Decenas: este valor es obtenido de la funcion Decenas
    '             3. Digito: este valor eds obtenido de la función
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Centenas_Letras(Ciento, Decena, Digito As String)
Dim Letras As String 'para asignarle la cadena a letras
Letras = ""
    Select Case Ciento
        Case "1" 'Para los número del 100 al 199
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "CIEN"
                Else
                    Letras = "CIENTO " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "CIENTO " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "2" 'Para los número del 200 al 299
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "DOSCIENTOS"
                Else
                    Letras = "DOSCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "DOSCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "3" 'Para los número del 300 al 399
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "TRESCIENTOS"
                Else
                    Letras = "TRESCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "TRESCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "4" 'Para los número del 400 al 499
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "CUATROCIENTOS"
                Else
                    Letras = "CUATROCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "CUATROCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "5" 'Para los número del 500 al 599
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "QUINIENTOS"
                Else
                    Letras = "QUINIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "QUINIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "6" 'Para los número del 600 al 699
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "SEISCIENTOS"
                Else
                    Letras = "SEISCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
            Else
                Letras = "SEISCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "7" 'Para los número del 700 al 799
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "SETECIENTOS"
                Else
                    Letras = "SETECIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
              Else
                Letras = "SETECIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "8" 'Para los número del 800 al 899
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "OCHOCIENTOS"
                Else
                    Letras = "OCHOCIENTOS " & Convierte_Unidades_Letras(Digito)
                End If
              Else
                Letras = "OCHOCIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
        Case "9" 'Para los número del 900 al 999
            If Decena = "0" Then
                If Digito = "0" Then
                    Letras = "NOVECIENTOS"
                Else
                    Letras = "NOVECIENTOS " & Convierte_Unidades_Letras(Digito)
              End If
              Else
                Letras = "NOVECIENTOS " & Convierte_Decenas_Letras(Decena, Digito)
            End If
    End Select
    Convierte_Centenas_Letras = Letras
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Decenas_Letras
    'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
    'PARÁMETROS:
    '             1. Decena: Es el número al convertir a letra
    '             2. Digito: es el nombre de los números obtenidos de la funcion Unidades
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Decenas_Letras(Decena, Digito As String)
Dim Letras As String 'Para asignarle la cadena a letras
Letras = ""
    'Escribe el nombre de acuerdo al número
    Select Case Decena
        Case "1" 'Para número del 10 al 19
            Select Case Digito
                Case "0"
                    Letras = "DIEZ"
                Case "1"
                    Letras = "ONCE"
                Case "2"
                    Letras = "DOCE"
                Case "3"
                    Letras = "TRECE"
                Case "4"
                    Letras = "CATORCE"
                Case "5"
                    Letras = "QUINCE"
                Case "6", "7", "8", "9"
                    Letras = "DIECI" & Convierte_Unidades_Letras(Digito)
              End Select
        Case "2" 'Para número del 20 al 29
            If Digito = "0" Then
                Letras = "VEINTE"
            Else
                Letras = "VEINTI" & Convierte_Unidades_Letras(Digito)
            End If
        Case "3" 'Para número del 30 al 39
            If Digito = "0" Then
                Letras = "TREINTA"
            Else
                Letras = "TREINTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "4" 'Para número del 40 al 49
            If Digito = "0" Then
                Letras = "CUARENTA"
            Else
                Letras = "CUARENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "5" 'Para número del 50 al 59
            If Digito = "0" Then
                Letras = "CINCUENTA"
            Else
                Letras = "CINCUENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "6" 'Para número del 60 al 69
            If Digito = "0" Then
                Letras = "SESENTA"
            Else
                Letras = "SESENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "7" 'Para número del 70 al 79
            If Digito = "0" Then
                Letras = "SETENTA"
            Else
                Letras = "SETENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "8" 'Para número del 80 al 89
            If Digito = "0" Then
                Letras = "OCHENTA"
            Else
                Letras = "OCHENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
        Case "9" 'Para número del 90 al 99
            If Digito = "0" Then
                Letras = "NOVENTA"
            Else
                Letras = "NOVENTA Y " & Convierte_Unidades_Letras(Digito)
            End If
    End Select
    Convierte_Decenas_Letras = Letras
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Unidades_Letras
    'DESCRIPCIÓN: Convierte  de cantidad de numero a letras
    'PARÁMETROS:
    '             1. Digito Es el número al convertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Convierte_Unidades_Letras(Digito As String)
    'Escribe el nombre de los número del 1 al 9
    Select Case Digito
        Case "0"
            Convierte_Unidades_Letras = "CERO"
        Case "1"
            Convierte_Unidades_Letras = "UN"
        Case "2"
            Convierte_Unidades_Letras = "DOS"
        Case "3"
            Convierte_Unidades_Letras = "TRES"
        Case "4"
            Convierte_Unidades_Letras = "CUATRO"
        Case "5"
            Convierte_Unidades_Letras = "CINCO"
        Case "6"
            Convierte_Unidades_Letras = "SEIS"
        Case "7"
            Convierte_Unidades_Letras = "SIETE"
        Case "8"
            Convierte_Unidades_Letras = "OCHO"
        Case "9"
            Convierte_Unidades_Letras = "NUEVE"
    End Select
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Convierte_Valida_Cantidad
    'DESCRIPCIÓN: Sirve para validar la cantidad
    'PARÁMETROS:
    '             1. Cantidad: Es el número a combertir a letra
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Convierte_Valida_Cantidad(Cantidad As String)
Dim I As Integer
Dim Lista As String

    Lista = "[0-9]"
    I = 1
    Do While (I <= Len(Cantidad)) And (Mid(Cantidad, I, 1) Like Lista)
        I = I + 1
    Loop
    If I < Len(Cantidad) Then
        If Mid(Cantidad, I, 1) = "." Then
          If (Mid(Cantidad, I + 1, 1) Like Lista) And (Mid(Cantidad, I + 2, 1) Like Lista) Then
            Convierte_Valida_Cantidad = True
          Else
            Convierte_Valida_Cantidad = False
          End If
        Else
            Convierte_Valida_Cantidad = False
        End If
    Else
        Convierte_Valida_Cantidad = True
    End If
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Recordset_Agregar
    'DESCRIPCIÓN: Crea un recordset para agregar registros y lo devuelve
    '             para su uso posterior, basandose en la consulta SQL
    'PARÁMETROS:
    '               1. Catalogo: Nombre de la tabla a en la que se agregaran los registros
    'CREO: Yazmin A Delgado Gómez
    'FECHA_CREO: 06-Junio-2005
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Recordset_Agregar(Catalogo As String) As rdoResultset
Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    'Se crea el recorset de Alta
    With Mi_Query
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT *"
        .SQL = .SQL & " FROM " & Catalogo & ""
        .LockType = rdConcurRowVer
        .CursorType = rdUseOdbc
        .RowsetSize = 1
        Set Recordset_Agregar = .OpenResultset(rdOpenKeyset)
    End With
End Function

Public Function Centra(Texto As String, LongEditor As Integer)
Dim Nuevo As String
Dim TxtLong, Left, I As Integer

    Nuevo = Trim(Texto)
    TxtLong = Len(Nuevo)
    Left = (LongEditor - TxtLong) / 2
    For I = 1 To Int(Left)
        Nuevo = " " & Nuevo
    Next I
    Centra = Nuevo
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Elimina_Catalogo
    'DESCRIPCIÓN: Elimina a un cliente o un usuario especifico
    'PARÁMETROS:
    '             1. Catalogo: Nombre de la Tabla de la Base de Datos
    '             2. ID: Clave del Cliente (Cliente_ID)de la Tabla Cat_Clientes
    '             3. Valor: paso del valor contenido en la caja de texto Txt_Cliente_ID
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
'
Public Function Elimina_Catalogo(Catalogo As String, ID As String, Valor As String) As Boolean
Dim Rs_Elimina As rdoResultset     'Manejo de registro
Dim Mi_SQL As String               'SQL a enviar a la funcion
Dim Contador As Integer

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM " & Catalogo & " WHERE " & ID & " ='" & Valor & "'"
    'Selecciona el registro
    Set Rs_Elimina = Recordset_Eliminar(Mi_SQL)
    'Elimina todos los datos
    If Not Rs_Elimina.EOF Then
        'Sirva para eliminar todos los datos que pueda contener con el mismo ID
        While Not Rs_Elimina.EOF
            Rs_Elimina.Delete
            Rs_Elimina.MoveNext
        Wend
        Rs_Elimina.Close
        Elimina_Catalogo = True
    Else
        Elimina_Catalogo = False
    End If
    Conexion_Base.CommitTrans
    Exit Function
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Limpiar_Textos
    'DESCRIPCIÓN: 'Limpia las cajas de texto de un catalogo en espacífico
    'PARÁMETROS:
    '             1. Forma: Nombre de la Forma a la que queremos borrar el contenido
    '                       de las cajas de texto
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Limpiar_Textos(Forma As Form)
Dim Mi_Text As Object    'Indica que objeto se esta consultando
    '1. Consulta todos los objetos contienidos en la forma
    '2. Limpia las cajas de texto
    For Each Mi_Text In Forma.Controls
        If TypeOf Mi_Text Is TextBox Then Mi_Text.Text = ""
        'If TypeOf Mi_Text Is ComboBox Then Mi_Text.ListIndex = -1
    Next
End Sub

''*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo_Item
    'DESCRIPCIÓN: Llena y Consulta el ComboBox de la forma
    'PARÁMETROS:
    '             1. Campos: Campo a consultar para el llenado del ComboBox
    '             2. Tabla: Nombre de la tabla a consultar
    '             3. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             4. Tipo: Para saber si esta consultando
    '             5. Campo_Con: Para consultar y llenar el campo con las palabras introducidas
    '                           por el usuario
    'CREO: Jorge Razo
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Llena_Combo_Item(Campos As String, Tabla As String, Combo_Control As ComboBox, Tipo As Integer, Campo_con As String, Optional Condicion_Adicional As String = "", Optional Mostrar_Todos As Boolean, Optional Mensaje_Todos As String, Optional Order_By As Boolean = False)
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & Campo_con & " LIKE '%" & Combo_Control.Text & "%' " & Condicion_Adicional
        End If
        If Order_By Then
            .SQL = .SQL & " ORDER BY " & Campo_con
        End If
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    Combo_Control.Clear
    If Mostrar_Todos Then
        If Mensaje_Todos <> "" Then
            Combo_Control.AddItem Mensaje_Todos
        Else
            Combo_Control.AddItem "TODOS"
        End If
        Combo_Control.ItemData(Combo_Control.NewIndex) = 0
    End If
    If Not Rs_Combo.EOF Then
        While Not Rs_Combo.EOF
            Combo_Control.AddItem Rs_Combo(1)
            Combo_Control.ItemData(Combo_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Sub

''*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Combo_Item
    'DESCRIPCIÓN: Llena y Consulta el ListBox de la forma
    'PARÁMETROS:
    '             1. Campos: Campo a consultar para el llenado del ComboBox
    '             2. Tabla: Nombre de la tabla a consultar
    '             3. Combo_Control: Nombre del ComboBox de la forma el cual se
    '                               va a llenar con los valores
    '             4. Tipo: Para saber si esta consultando
    '             5. Campo_Con: Para consultar y llenar el campo con las palabras introducidas
    '                           por el usuario
    'CREO: Antonio Benavides
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Llena_List_Item(Campos As String, Tabla As String, List_Control As ListBox, Tipo As Integer, Campo_con As String, Optional Condicion_Adicional As String = "", Optional Mostrar_Todos As Boolean, Optional Mensaje_Todos As String)
Dim Mi_SQL As New rdoQuery      'Obtiene los valores de la consulta
Dim campos_cont As Integer      'Obtiene el número de campos existentes en la BD
Dim Rs_Combo As rdoResultset    'Manejo de registro
Dim I As Integer
    
    'Consulta el campo
    With Mi_SQL
        Set .ActiveConnection = Conexion_Base
        .SQL = "SELECT " & Campos
        .SQL = .SQL & " FROM " & Tabla
        If Tipo = 1 Then
            .SQL = .SQL & " WHERE " & Campo_con & " LIKE '%" & List_Control.Text & "%' " & Condicion_Adicional
        End If
        .SQL = .SQL & " ORDER BY " & Campo_con
        .LockType = rdConcurReadOnly
        Set Rs_Combo = .OpenResultset
    End With
    'Llena el ComboBox de la forma
    List_Control.Clear
    If Mostrar_Todos Then
        If Mensaje_Todos <> "" Then
            List_Control.AddItem Mensaje_Todos
        Else
            List_Control.AddItem "TODOS"
        End If
        List_Control.ItemData(List_Control.NewIndex) = 0
    End If
    If Not Rs_Combo.EOF Then
        While Not Rs_Combo.EOF
            List_Control.AddItem Rs_Combo(1)
            List_Control.ItemData(List_Control.NewIndex) = Rs_Combo(0)
            Rs_Combo.MoveNext
        Wend
    End If
    Rs_Combo.Close
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Llena_Grid
'DESCRIPCIÓN:   LLena un grid en base a la consulta requerida
'PARÁMETROS:    Grid:   Nombre del Control Grid que se desea llenar
'               Nombre_Columnas:    Columnas que se desean mostrar
'                                   estas deben coincidir con los nombres de las tablas
'               Tamaño_Columnas:    Es el tamaño de las columnas que seran mostradas en el grid
'               Al_Ancho_del_Formulario:    Indica si se dese que el grid se amplie al Ancho del formulario
'               Al_Alto_del_Formulario:    Indica si se dese que el grid se amplie al Alto del formulario
'               Porcentaje_Izquierdo:   Indica el espacio que se desea dejar de margen Izquierdo entre el formulario y el grid
'               Porcentaje_Derecho:     Indica el espacio que se desea dejar de margen Derecho entre el formulario y el grid
'               Porcentaje_Superior:    Indica el espacio que se desea dejar de margen Superior entre el formulario y el grid
'               Porcentaje_Inferior:    Indica el espacio que se desea dejar de margen Inferior entre el formulario y el grid
'               NOTA:               LOS DATOS DE NOMBRE_COLUMNAS Y TAMAÑO_COLUMNAS DEBEN ESTAR
'                                   SEPARADOS POR COMAS, ADEMAS DE CONTENER EL MISMO NUMERO DE
'                                   ELEMENTOS
'CREO:          Ricardo Soria
'FECHA_CREO:    08 de Msrzo 2007
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Llena_Grid(Grid As MSFlexGrid, Nombre_Columnas As String, Tamaño_Columnas As String, Optional Respetar_Renglones As Boolean = False)
    Dim Rs_Consulta As rdoResultset     '#  Manejador de datos para consulta de facturas
    Dim Arr_Nombre_Columnas() As String '#  Arreglo para controlar las columnas que se mostraran en el Grid
    Dim Arr_Tamaño_Columnas() As String '#  Arreglo que contiene el ancho de las columnas del Grid
    Dim Registro As String              '#  Variable para formar el registro del Grid
    Dim Cont_Arr As Integer             '#  Contador para Conocer la posicion del Arreglo
    Dim Ancho_Grid As Double           '#  Determina el Ancho del Grid
    Dim Alto_Grid As Double            '#  Determina el Alto del Grid
    Dim Contador As Integer
        
    Arr_Nombre_Columnas = Split(Nombre_Columnas, ",")   '#  OBTIENE LOS NOMBRES DE LA COLUMNAS
    Arr_Tamaño_Columnas = Split(Tamaño_Columnas, ",")   '#  OBTIENE EL TAMAÑO DE LAS COLUMNAS DEL GRID
    If Respetar_Renglones = False Then Grid.Rows = 0
    Grid.Cols = UBound(Arr_Nombre_Columnas) + 1
    Grid.Gridlines = flexGridNone
    '#  SI EL NUMERO DE COLUMNAS A MOSTRAR ES IGUAL AL NUMERO DE VALORES PARA CONFIGURAR
    '   LAS COLUMNAS A MOSTRAR
    On Error GoTo error:
    If UBound(Arr_Nombre_Columnas) = UBound(Arr_Tamaño_Columnas) Then
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta
            '#  MSI CONTIENE REGISTROS
            If Not .EOF Then
                '#  Indicador para comprobar si pondra el encabezado
                If Respetar_Renglones = False Or Grid.Rows = 0 Then
                    '#  RECORRE EL ARREGLO DE LAS COLUMNAS A MOSTRAR PARA ARMAR EL ENCABEZADO DEL GRID
                    For Cont_Arr = LBound(Arr_Nombre_Columnas) To UBound(Arr_Nombre_Columnas)
                        '#  SI SE INDICO QUE ES DATO TIPO MONEDA
                        If Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "$" Or Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "%" Or _
                        Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "#" Then
                            Registro = Registro & Mid(Arr_Nombre_Columnas(Cont_Arr), 2, Len(Arr_Nombre_Columnas(Cont_Arr)))
                        Else '#  EN CASO DE NO SER TIPO MONEDA
                            Registro = Registro & Arr_Nombre_Columnas(Cont_Arr)
                        End If
                        '#  SI NO SE HA LLEGADO AL FINAL DEL ARREGLO SE AGREGA UN ESPACIO
                        '   PARA SIMULAR EL CAMBIO DE COLUMNA
                        If Cont_Arr < UBound(Arr_Nombre_Columnas) Then Registro = Registro & Chr(9)
                    Next
                    Grid.AddItem Registro
                End If
                '#  SI NO SE HA LLEGADO AL FINAL DE LOS REGISTROS DEL RECORDSET
                While Not .EOF
                    Registro = ""
                    '#  RECORRE EL ARREGLO DE LAS COLUMNAS A MOSTRAR PARA ARMAR EL REGISTRO DEL GRID
                    For Cont_Arr = LBound(Arr_Nombre_Columnas) To UBound(Arr_Nombre_Columnas)
                        '#  SI SE INDICO QUE ES DATO TIPO MONEDA
                        If Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "$" Or Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "%" Or _
                        Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) = "#" Then
                            Registro = Registro & Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1) & " " & .rdoColumns(Mid(Arr_Nombre_Columnas(Cont_Arr), 2, Len(Arr_Nombre_Columnas(Cont_Arr))))
                        Else    '#  EN CASO DE NO SER TIPO MONEDA
                            Registro = Registro & .rdoColumns(Arr_Nombre_Columnas(Cont_Arr))
                        End If
                        '#  SI NO SE HA LLEGADO AL FINAL DEL ARREGLO SE AGREGA UN ESPACIO
                        '   PARA SIMULAR EL CAMBIO DE COLUMNA
                        If Cont_Arr < UBound(Arr_Nombre_Columnas) Then Registro = Registro & Chr(9)
                    Next
                    Grid.AddItem Registro
                    .MoveNext
                Wend
                Grid.FixedRows = 1
                Grid.FixedCols = 1
            End If
            .Close
        End With
        '#  Le da el tamaño a las columnas de acuerdo al porcentaje deseado 0%--100%
        For Cont_Arr = LBound(Arr_Tamaño_Columnas) To UBound(Arr_Tamaño_Columnas)
            Grid.ColWidth(Cont_Arr) = Arr_Tamaño_Columnas(Cont_Arr) * Grid.Width / 100
        Next
        '#  Busca las columnas para darles formato de acuerdo a sus valores
        For Cont_Arr = LBound(Arr_Nombre_Columnas) To UBound(Arr_Nombre_Columnas)
            '#  Doy formato a los encabezados
            Grid.Row = 0
            Grid.Col = Cont_Arr
            'Grid.CellFontBold = True
            Grid.CellAlignment = 3
            '#  Si se necesita formato
            Select Case Mid(Arr_Nombre_Columnas(Cont_Arr), 1, 1)
                '#  Da formato a moneda
                Case "$":
                    For Contador = 1 To Grid.Rows - 1
                        Grid.Row = Contador
                        Grid.TextMatrix(Contador, Cont_Arr) = "$ " & Format(Grid.TextMatrix(Contador, Cont_Arr), "#,###,###,#.00")
                        Grid.CellAlignment = 6
                    Next
                '#  da formato a porcentajes
                Case "%":
                    For Contador = 1 To Grid.Rows - 1
                        Grid.Row = Ciclos
                        Grid.TextMatrix(Contador, Cont_Arr) = Format(Grid.TextMatrix(Contador, Cont_Arr), "00.00") & " %"
                        Grid.CellAlignment = 3
                    Next
                '#  da formato a fechas
                Case "#":
                    For Contador = 1 To Grid.Rows - 1
                        Grid.Row = Contador
                        Grid.TextMatrix(Contador, Cont_Arr) = Format(Mid(Grid.TextMatrix(Contador, Cont_Arr), 2, Len(Grid.TextMatrix(Contador, Cont_Arr))), "dd/MMM/yyyy")
                        Grid.CellAlignment = 4
                    Next
            End Select
        Next
        Set Rs_Consulta = Nothing
    Else '#  INDICA QUE NO COINCIDEN EL TOTAL DEL COLUMNAS A MOSTRAR CON EL TOTAL DE COLUMNAS PAR CONFIGURAR
        '       EL ANCHO DE ESTAS
        MsgBox "No coinciden los tamaños", vbInformation
    End If
    Exit Sub
error:
    If Err.Number = 40041 Then
        MsgBox "Nombre de Columna no Encontrado", vbInformation
    End If
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Validacion_Accesos_Sistema
    'DESCRIPCIÓN: 1. Consulta el menu o submenu que se requiere realizar alguna
    '                operación para consultar las operaciones que puede realizar
    '                el rol al cual pertenece el usuario
    '             2. De acuerdo al nombre que tenga asignado el Tag del boton se
    '                habilita o deshabilita el mismo
    'PARÁMETROS : Nombre_Menu_submenu: Indica el nombre del menu o submenu que se
    '                                  desea realizar las operaciones
    '             Forma: Indica el nombre de la forma en el cual se va a validar
    '                    las operaciones
    'CREO       : 28-Mayo-2007
    'FECHA_CREO : Yazmin Delgado Gómez
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Validacion_Accesos_Sistema(Nombre_Menu_Submenu As String, Forma As Form)
Dim Ctrl As Control                             'Indica que control de la forma se esta consultando
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta las operaciones que puede realizar

    'Consulta las operaciones que tiene habilitada
    Mi_SQL = "SELECT Alta,Cambio,Eliminar,Consultar"
    Mi_SQL = Mi_SQL & " FROM Seguridad_Sistema "
    Mi_SQL = Mi_SQL & " WHERE Rol_ID='" & Rol_ID & "'"
    Mi_SQL = Mi_SQL & " AND Nombre_Sistema='" & Nombre_Menu_Submenu & "'"
    Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Apl_Cat_Accesos.EOF Then
        With Rs_Consulta_Apl_Cat_Accesos
            'Recorre la forma para consulta los botones que tiene
            For Each Ctrl In Forma.Controls
                If TypeOf Ctrl Is CommandButton Then
                    Select Case Mid(Ctrl.Tag, 1, 1)
                        Case "A":
                            If .rdoColumns("Alta") = "S" Then
                                If Ctrl.Enabled = True Then Ctrl.Enabled = True
                            Else
                                Ctrl.Enabled = False
                            End If
                        Case "M"
                            If .rdoColumns("Cambio") = "S" Then
                                If Ctrl.Enabled = True Then Ctrl.Enabled = True
                            Else
                                Ctrl.Enabled = False
                            End If
                        Case "B"
                            If .rdoColumns("Eliminar") = "S" Then
                                If Ctrl.Enabled = True Then Ctrl.Enabled = True
                            Else
                                Ctrl.Enabled = False
                            End If
                        Case "C"
                            If .rdoColumns("Consultar") = "S" Then
                                If Ctrl.Enabled = True Then Ctrl.Enabled = True
                            Else
                                Ctrl.Enabled = False
                            End If
                        Case ""
                            If Ctrl.Enabled = True Then Ctrl.Enabled = True
                    End Select
                End If
            Next
        End With
    End If
    Rs_Consulta_Apl_Cat_Accesos.Close
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Mover control en grid
    'DESCRIPCIÓN: Permite colocar un control en un grid en las celdas que se especifiquen
    'PARÁMETROS: Flex grid – es el nombre del grid al que se pondra el control
    '            Control_Mover es el nombre del control que se quiere poner en el grid
    'CREO: Jorge Razo
    'FECHA_CREO: 24 Agosto 2006
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Mover_Control_Grid_CheckBox(Flex_Grid As MSFlexGrid, Optional Control_Mover As CheckBox) 'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = (Flex_Grid.Left) + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            .Height = Flex_Grid.CellHeight
            .Visible = True
            .SetFocus
        End If
    End With
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Mover_Control_Grid_ListBox
    'DESCRIPCIÓN: Permite colocar un combo en un grid en las celdas que se especifiquen
    'PARÁMETROS: Flex grid – es el nombre del grid al que se pondra el control
    '            Control_Mover es el nombre del control que se quiere poner en el grid
    'CREO: Jorge Razo
    'FECHA_CREO: 10/Agosto/2007
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Mover_Control_Grid_ListBox(Flex_Grid As MSFlexGrid, Control_Mover As ListBox)        'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = Flex_Grid.Left + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            .Height = Flex_Grid.CellHeight
            'Busca el registro segun el dato que contenga el grid en la celda seleccionada
            Call Asigna_Item_List(Flex_Grid.TextMatrix(Flex_Grid.RowSel, Flex_Grid.ColSel), Control_Mover)
            .Visible = True
            .Enabled = True
            .SetFocus
        End If
    End With
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Mover control en grid
    'DESCRIPCIÓN: Permite colocar un control en un grid en las celdas que se especifiquen
    'PARÁMETROS: Flex grid – es el nombre del grid al que se pondra el control
    '            Control_Mover es el nombre del control que se quiere poner en el grid
    'CREO: Jorge Razo
    'FECHA_CREO: 24 Agosto 2006
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Mover_Control_Grid_TextBox(Flex_Grid As MSFlexGrid, Control_Mover As TextBox)      'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = Flex_Grid.Left + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            .Height = Flex_Grid.CellHeight
            .Text = Flex_Grid
            .Visible = True
            .Enabled = True
            .SetFocus
            SendKeys "{Home}+{End}"
        End If
    End With
End Sub

Public Sub Mover_Control_Grid_ComboBox(Flex_Grid As MSFlexGrid, Control_Mover As ComboBox)      'Pone el control sobre el campo selecionado
    With Control_Mover
        If Flex_Grid.CellWidth > 0 Then
            .Top = Flex_Grid.Top + Flex_Grid.CellTop
            .Left = Flex_Grid.Left + Flex_Grid.CellLeft
            .Width = Flex_Grid.CellWidth
            '.Height = Flex_Grid.CellHeight
            .Text = Flex_Grid
            .Visible = True
            .Enabled = True
            .SetFocus
            'SendKeys "{Home}+{End}"
        End If
    End With
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Asigna_Item_List
    'DESCRIPCIÓN: 'Selecciona un registro especifico de un listbox
    'PARÁMETROS:
    '             1. Valor: Clave oculta del registro en el listbox
    '             2. Cmb_Dato: Combo que contiene la lista de registros
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Asigna_Item_List(Valor As String, Lst_Dato As ListBox)
Dim I As Integer

For I = 0 To Lst_Dato.ListCount - 1
    If Lst_Dato.List(I) = Valor Then
        Lst_Dato.ListIndex = I
        Exit For
    End If
Next I
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Buscar_Cadena_Combo
    'DESCRIPCIÓN: Busca alguna cadena de caracteres dentro de los elementos de un combo
    '             agregados previamente desde diseño o con un Recordset y me
    '             regresa el indice donde lo encuentra
    'PARÁMETROS:
    '             1. Cadena: Es la cadena que se va a buscar dentro de los elementos del Combo
    '             2. ID: Clave del Cliente (Cliente_ID)de la Tabla Cat_Clientes
    '             3. Valor: paso del valor contenido en la caja de texto Txt_
    '             4. Caracteres: Manda el numero de caracteres que se quieren tomar de la
    '                izquierda del elemento del combo para compraralos contra la cadena
    '                que se busca, si caracteres=0 se comprara la cadena completa del elemento
    '                del combo
    'CREO: Susana Ledesma Ramirez
    'FECHA_CREO:2/Mayo/2006
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
    
'*******************************************************************************
Public Function Buscar_Cadena_Combo(Cadena As String, Combo As ComboBox, Optional Caracteres As Integer) As Integer
    Dim Indice As Integer
    Dim Indice_Aux As Integer
    Dim Comparar_Cadena As String       'Cadena que se usa para tomar solo los caracetres a comparar de los
                                    'elementos del combo
    'El combo debe tener algunos elementos agregados previamente,si no; se deben agregar
    'uno por uno sin usar esta funcion
    Indice_Aux = -1
    For Indice = 0 To Combo.ListCount - 1
        'Asigno todos los caracteres del elemento del combo,
        'cuando no necesito tomar solamente algunos caracteres
        Comparar_Cadena = Combo.List(Indice)
        'Si necesito solo algunos caracetres corto la cadena
        'utilizando el numero de caracteres que se mandaron en la variable OPCIONAL caracteres
        If Caracteres > 0 Then Comparar_Cadena = Left(Combo.List(Indice), Caracteres)
        If Trim(Comparar_Cadena) = Trim(Cadena) Then
            Indice_Aux = Indice
            Exit For
        End If
    Next Indice
    Buscar_Cadena_Combo = Indice_Aux
End Function
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Asigna_Item_Combo
    'DESCRIPCIÓN: 'Selecciona un registro especifico de un combo
    'PARÁMETROS:
    '             1. Valor: Clave oculta del registro en el combo
    '             2. Cmb_Dato: Combo que contiene la lista de registros
    'CREO: Ruben García
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub Asigna_Item_Combo(Valor As String, Cmb_Dato As ComboBox, Optional Tam_Format As Integer = 5)
Dim I As Integer
Dim format_string As String
    
    format_string = "00000"
    If Tam_Format <> 5 Then
        format_string = ""
        For I = 1 To Tam_Format
            format_string = format_string & "0"
        Next
    End If
    Cmb_Dato.ListIndex = -1
    For I = 0 To Cmb_Dato.ListCount - 1
        If Cmb_Dato.List(I) = Valor Or Format(Cmb_Dato.ItemData(I), format_string) = Valor Then
            Cmb_Dato.ListIndex = I
            Exit For
        End If
    Next I
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Quitar_Caracteres_Raros
    'DESCRIPCIÓN: Quita de la cadena los caracteres que pueden ocacionar un error
    '             en los querys de consulta
    'PARÁMETROS:
    '             1. KeyAscii: Regresa el valor de la letra
    'CREO: Jose Antonio Lopez Hernandez
    'FECHA_CREO: 20/Junio/2007 5:53 pm
    'MODIFICO:
    'FECHA_MODIFICO:
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Quitar_Caracteres_Raros(KeyAscii As Integer, Optional Mayusculas As Boolean = False)
    If Mayusculas = True Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'Acepta las letras de la A a la Z, incluyendo Ñ o ñ, espacios y Puntos
    If Chr(KeyAscii) = "'" Then
        If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 209 Then
            Exit Function
        End If
        KeyAscii = 0
    End If
End Function



'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Formulario_Cargado
    'DESCRIPCIÓN: Indica si el formulario esta cargado o no dentro del sistema
    'PARÁMETROS : Nombre_Formulario: Nombre del formulario el cual se pretende saber si
    '             esta cargado o no en el sistema, si esta cargada la forma
    '             manda un valor verdadero
    'CREO       : Yazmin Delgado Gómez
    'FECHA_CREO : 19-Julio-2007
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Function Formulario_Cargado(Nombre_Formulario As String) As Boolean
Dim Frm_Control As Form

On Error Resume Next
    'Recorre sólo los controles de tipo forma que esten cargados en el MDI para validar
    'que el nombre del control de la forma a buscar este cargado
    For Each Frm_Control In Forms
        'Si esta cargado entonces retorna un valor verdadero
        If StrComp(Frm_Control.Caption, Nombre_Formulario, vbTextCompare) = 0 Then
          Formulario_Cargado = True
          Exit For
        End If
    Next
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Enfocar
    'DESCRIPCIÓN: Manda el enfoque a la forma con el Caption correspondiente, ya que no se
    '             puede enviar el Setfocus directamente porque se estan utilizando objetos
    '             locales de cada formulario y el enfoque lo envia al formulario original
    '             y no al del objeto que se creo (Menus de Catalogos)
    'PARÁMETROS : Titulo: Es el Caption de la forma que se esta buscando para enviarle el enfoque
    'CREO       : Susana Ledesma Ramírez
    'FECHA_CREO : 13-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Enfocar(Titulo As String)
Dim Forma As Form 'Formulario al que se va a enviar el enfoque
    For Each Forma In Forms
        If Forma.Caption = Titulo Then
            Forma.SetFocus
            Exit For
        End If
    Next
End Sub


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Quitar_Caracter_Raro
    'DESCRIPCIÓN:          Permite caracteres en el rango establecido: rango[A-B,a-b,0-9,(,),Backspace,Tab,Spacebar,Supr]
    '                      Regresa el caracter si esta dentro del rango, ó caracter 0 caso contrario
    '                      se utiliza en el keypress del control
    'PARÁMETROS:           Caracter_KeyAscii: entrada del teclado
    'CREO:                 Yañez Rodriguez Diego Neftali
    'FECHA_CREO:           13-Diciembre-2007
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Quitar_Caracter_Raro(Caracter_KeyAscii As Integer)
    
    'Regresa el valor en mayuscula para cacteres alfabeticos
    Caracter_KeyAscii = Asc(UCase(Chr(Caracter_KeyAscii)))
    
    'Rango de caracteres validos para la funcion
    Select Case (Caracter_KeyAscii)
        'Numeros
        Case vbKey0 To vbKey9: Exit Function
        'Letras
        Case vbKeyA To vbKeyZ: Exit Function
        'Barra Espaciadora
        Case vbKeySpace: Exit Function
        'Regresa Espacio,
        Case vbKeyBack: Exit Function
        'Parentesis ( )
        Case 40 To 41: Exit Function
        'Supr
        Case vbKeyDelete: Exit Function
        'Tab
        Case vbKeyTab: Exit Function
        'Ñ
        Case 209: Exit Function
        'Tecla Enter
        Case vbKeyReturn: Exit Function
        '-
        Case 45: Exit Function
        '# " % / ` º $ *
        Case 35, 34, 37, 47, 44, 38, 96, 186, 36, 42: Exit Function
    End Select
    
    'Fuera de rango
    Caracter_KeyAscii = 0
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Buscar_List_Combo
    'DESCRIPCIÓN:           Función para buscar un string dentro de una lista del combo
    'PARÁMETROS :
    '
    'CREO       :           Yañez Rodriguez Diego
    'FECHA_CREO :           18 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Buscar_List_Combo(Cmb As ComboBox, KeyCode As Integer)
Static Cadena As String
    Dim I As Long
    
    With Cmb
        
        If KeyCode = vbKeyUp Then Exit Sub
        If KeyCode = vbKeyDown Then Exit Sub
        If KeyCode = vbKeyLeft Then Exit Sub
        If KeyCode = vbKeyRight Then Exit Sub
        If KeyCode = vbKeyHome Then Exit Sub
        ' verifica qu no se presionó la tecla backspace
        If KeyCode <> vbKeyBack Then
            Cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
        Else
            '...tecla backspace
            If Cadena <> "" Then
                Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
            Else
                Exit Sub
            End If
        End If
        For I = 0 To .ListCount - 1
            If UCase(Cadena) = UCase(Mid(.List(I), 1, Len(Cadena))) Then
                .ListIndex = I
                Exit For
            End If
        Next
        ' Seelecciona
        .SelStart = Len(Cadena)
        .SelLength = Len(.Text)
    End With
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Es_Dia_No_Laboral
    'DESCRIPCIÓN:          Regresa un valor true si el dias se encuentra en la tabla de dias no laborales
    'PARÁMETROS :          fecha: la fecha para conocer si se labora o no
    'CREO       :          Yañez Rodriguez Diego Neftali
    'FECHA_CREO :          05-Marzo-2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Es_Dia_No_Laboral(Fecha As Date) As Boolean
Dim Mi_SQL As String                      'Obtiene los valores de la consulta a realizar
Dim Rs_Consulta_Cat_Dias_No_Laborales As rdoResultset 'Consulta todos los valores que tiene dados de alta la base de datos en la tabla correspondiente
    
    Es_Dia_No_Laboral = False
    Mi_SQL = "SELECT * FROM Cat_Dias_No_Laborales WHERE Fecha = '" & Format(Fecha, "MM/dd/yyyy") & "'"
    Set Rs_Consulta_Cat_Dias_No_Laborales = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Dias_No_Laborales.EOF Then
        Es_Dia_No_Laboral = True
    End If
    Rs_Consulta_Cat_Dias_No_Laborales.Close
    Set Rs_Consulta_Cat_Dias_No_Laborales = Nothing
End Function

'*************************************************************************************
'NOMBRE DE LA FUNCIÓN: Busca_Dato_BD
'DESCRIPCIÓN:          Query_BD. Consulta a la BD
'                      Campo_Regresar:   Campo a regresar
'PARÁMETROS:
'CREO:                 Ricardo Soria
'FECHA_CREO:           2-Ene-2008
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*************************************************************************************
Public Function Busca_Dato_BD(Query_BD As String, Campo_Regresar As String) As String
Dim Rs_Consulta_Tabla As rdoResultset   '#  consulta la tabla
On Error GoTo error:

    Busca_Dato_BD = ""
    Set Rs_Consulta_Tabla = Conectar_Ayudante.Recordset_Consultar(Query_BD)
    With Rs_Consulta_Tabla
        '#  Si encontro el dato
        If Not .EOF Then
            Busca_Dato_BD = Trim(.rdoColumns(Campo_Regresar))
        End If
        .Close
    End With
    Set Rs_Consulta_Tabla = Nothing
Exit Function
error:
    MsgBox "Campo no encontrado: " + Campo_Regresar, vbInformation + vbOKOnly
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Agregar_Espacios
    'DESCRIPCIÓN:          Agrega espacios en blanco a una cadena de caracteres
    'PARÁMETROS :          Cadena: Cadena a la que se agregaran espacios
    '                      Espacios: Cantidad de espacios a agregar
    'CREO       :          Yañez Rodriguez Diego Neftali
    'FECHA_CREO :          05-Marzo-2008
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Agregar_Espacios(Cadena As String, Espacios As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim I As Integer     'Cuenta el numero de caracteres de la cadena
    Nuevo = Cadena
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For I = 1 To Espacios - Len(Cadena)
        Nuevo = Nuevo & " "
    Next I
    Agregar_Espacios = Nuevo

End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alinea_Derecha
    'DESCRIPCIÓN: Alinea los número a la derecha del documento
    'PARÁMETROS:
    '             1. Numero: Es la cadena que deseo alinear
    '             2. Longitud: Es la longitud que va a contener mi cadena
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function Alinea_Derecha(Numero As String, Longitud As Integer) As String
Dim Nuevo As String  'Asignar la cadena
Dim I As Integer     'Cuenta el numero de caracteres de la cadena
    Nuevo = Numero
    'Sirve para llenar de espacios en blanco los caracteres a la derecha
    For I = 1 To Longitud - Len(Numero)
        Nuevo = " " & Nuevo
    Next I
    Alinea_Derecha = Nuevo
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Centrar_Texto
    'DESCRIPCIÓN: Centra el texto en la línea de impresón
    'PARÁMETROS: Texto     : Texto a centrar
    '            Longitud_Editor: Numero de caracteres a centrar
    'CREO: Joel G. Romero Cervantes
    'FECHA_CREO:
    'MODIFICO: Antonio Salvador Benavides Guardado
    'FECHA_MODIFICO: 13/junio/2008
    'CAUSA_MODIFICACIÓN: Estandarización
'*******************************************************************************
Public Function Centrar_Texto(Texto As String, Longitud_Editor As Integer)
Dim Txt_Sin_Espacios As String
Dim Txt_Longitud, Espacio_Izquierda, I As Integer
    Txt_Sin_Espacios = Trim(Texto)
    Txt_Longitud = Len(Txt_Sin_Espacios)
    Espacio_Izquierda = (Longitud_Editor - Txt_Longitud) / 2
    For I = 1 To Int(Espacio_Izquierda)
        Txt_Sin_Espacios = " " & Txt_Sin_Espacios
    Next I
    Centrar_Texto = Txt_Sin_Espacios
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Valida_Campo_Duplicado
    'DESCRIPCIÓN:           Función para validar si el campo
    'PARÁMETROS :
    '
    'CREO       :           Yañez Rodriguez Diego
    'FECHA_CREO :           18 Febrero 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Function Valida_Campo_Duplicado(Tabla As String, Campo As String, Valor As String, Optional Nombre_Campo_ID_Excluyente As String, Optional ID_Regitro_Excluyente As String) As Boolean
Dim Rs_Consulta_Tabla As rdoResultset       'Informacion de la tabla a consultar

    Valida_Campo_Duplicado = False
    Mi_SQL = "SELECT " & Campo & " FROM " & Tabla
    Mi_SQL = Mi_SQL & " WHERE " & Campo & " = '" & Valor & "'"
    If Nombre_Campo_ID_Excluyente <> "" And ID_Regitro_Excluyente <> "" Then
        Mi_SQL = Mi_SQL & " AND " & Nombre_Campo_ID_Excluyente & " <> '" & ID_Regitro_Excluyente & "'"
    End If
    Set Rs_Consulta_Tabla = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Tabla
        If Not .EOF Then
            Valida_Campo_Duplicado = True
            .Close
        End If
    End With
    Set Rs_Consulta_Tabla = Nothing
End Function

Public Function Buscar_Nombre(ID As String, Tabla As String, Campo As String, Condicion As String) As String
Dim Rs_Consulta As rdoResultset
    
        Mi_SQL = "SELECT " & Campo & " FROM " & Tabla
        Mi_SQL = Mi_SQL & "  WHERE " & Condicion & "='" & ID & "'"
        Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta.EOF Then
            With Rs_Consulta
               Buscar_Nombre = .rdoColumns(Campo)
                
            End With
        Rs_Consulta.Close
    End If
End Function


Public Function Elimina_Catalogo_2(Catalogo As String, ID As String, Valor As String) As Boolean
Dim Rs_Elimina As rdoResultset     'Manejo de registro
Dim Mi_SQL As String               'SQL a enviar a la funcion
Dim Contador As Integer

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM " & Catalogo & " WHERE " & ID & " ='" & Valor & "'"
    'Selecciona el registro
    Set Rs_Elimina = Recordset_Eliminar(Mi_SQL)
    'Elimina todos los datos
    If Not Rs_Elimina.EOF Then
        'Sirva para eliminar todos los datos que pueda contener con el mismo ID
        While Not Rs_Elimina.EOF
            Rs_Elimina.Delete
            Rs_Elimina.MoveNext
        Wend
        Rs_Elimina.Close
        Elimina_Catalogo_2 = True
    Else
        Elimina_Catalogo_2 = False
    End If
'    Conexion_Base.CommitTrans
    Exit Function
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Function

Public Function Elimina_Actualiza_Catalogo(Catalogo As String, Campo As String, Valor As String, Condicion As String, ID As String) As Boolean
Dim Rs_Elimina As rdoResultset     'Manejo de registro
Dim Mi_SQL As String               'SQL a enviar a la funcion
Dim Contador As Integer
Elimina_Actualiza_Catalogo = True
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Mi_SQL = "SELECT * FROM " & Catalogo & " WHERE " & Condicion & " = " & ID
    'Selecciona el registro
    Set Rs_Elimina = Recordset_Editar(Mi_SQL)
    'Elimina todos los datos
    With Rs_Elimina
    .Edit
    .rdoColumns(Campo) = Valor
    .Update
    .Close
    End With
'    Conexion_Base.CommitTrans
    Exit Function
HANDLER:
    Conexion_Base.RollbackTrans
    Elimina_Actualiza_Catalogo = False
    For Each Er In rdoErrors
        MsgBox Er.Description, vbInformation
    Next Er
End Function


Public Function NroEnLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True, Optional Coloca_Pesos As Boolean = True) As String
    'Devuelve un número expresado en letras.
    'El parámetro blnO_Final se utiliza en la recursión para saber si se debe colocar
    'la "O" final cuando la palabra es UN(O)
    Dim dblCentavos As Double
    Dim lngContDec As Long
    Dim lngContCent As Long
    Dim lngContMil As Long
    Dim lngContMillon As Long
    Dim strNumLetras As String
    Dim strNumero As Variant
    Dim strDecenas As Variant
    Dim strCentenas As Variant
    Dim blnNegativo As Boolean
    Dim blnPlural As Boolean

    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If

    strNumero = Array(vbNullString, "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciseis", "Diecisiete", "Dieciosocho", "Diecinueve", "Veinte")

    strDecenas = Array(vbNullString, vbNullString, "Veinti", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa", "Cien")

    strCentenas = Array(vbNullString, "Ciento", "Doscientos", "Trescientos", "Cuatroscientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop

    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop

    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If

    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then 'si el número es >1000000 usa recursividad
            strNumLetras = NroEnLetras(lngContMillon, False, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " Millón" & IIf(blnPlural, "ES ", " ")
    End If

    If lngContMil > 0 Then
        If lngContMil >= 1 Then 'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & NroEnLetras(lngContMil, False, False)
            lngContMil = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " Mil "
    End If

    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "Cien"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If

    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If

        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If Coloca_Pesos Then
                strNumLetras = Trim(strNumLetras) & " Pesos"
                If dblCentavos > 0# Then
                    strNumLetras = Trim(strNumLetras) & " " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
                Else
                    strNumLetras = Trim(strNumLetras) & " 0/100"
                End If
            End If
            NroEnLetras = strNumLetras
            Exit Function
        End If
    End If

    If curNumero > 0# Then
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If
    If Coloca_Pesos Then
        strNumLetras = Trim(strNumLetras) & " Pesos"
        If dblCentavos > 0# Then
            strNumLetras = strNumLetras & " " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
        Else
            strNumLetras = Trim(strNumLetras) & " 0/100"
        End If
    End If

    NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
End Function

