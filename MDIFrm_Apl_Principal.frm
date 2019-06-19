VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIFrm_Apl_Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SISTEMA DE RECURSOS HUMANOS"
   ClientHeight    =   9615
   ClientLeft      =   2820
   ClientTop       =   1560
   ClientWidth     =   11280
   Icon            =   "MDIFrm_Apl_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm_Apl_Principal.frx":0442
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9240
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Menu_Archivo 
      Caption         =   "&Archivo"
      WindowList      =   -1  'True
      Begin VB.Menu Submenu_Impresora 
         Caption         =   "Configurar Impresora"
      End
      Begin VB.Menu Submenu_Calculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu Submenu_Apl_Respaldo_Sistema 
         Caption         =   "Respaldo del Sistema"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Raya_Archivo 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Registrarse 
         Caption         =   "Registrarse"
      End
      Begin VB.Menu Submenu_Apl_Cambio_Password 
         Caption         =   "Cambio de Password"
      End
      Begin VB.Menu Submenu_Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Menu_Recursos_Humanos 
      Caption         =   "A&dministración"
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Importacion_Asistencias 
         Caption         =   "Importacion de Asistencias"
      End
      Begin VB.Menu SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias_Almacenes 
         Caption         =   "Importacion de Asistencias Almacenes"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo 
         Caption         =   "Validación de Tiempo de Trabajo"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Asistencia_Empleados 
         Caption         =   "Asistencias de Empleados"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Solicitud_Permisos 
         Caption         =   "Solicitud de Permiso"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias 
         Caption         =   "Incidencias Extraordinarias"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias 
         Caption         =   "Mantenimiento de Asistencias"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Notificación_Aniversarios 
         Caption         =   "Notificación de Aniversarios"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Control_Calzado 
         Caption         =   "Control de Calzado"
      End
      Begin VB.Menu Raya_1 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Adm_Cambio_Turno 
         Caption         =   "Cambio de Turnos"
      End
      Begin VB.Menu Submenu_Ope_Importacion_Archivo 
         Caption         =   "Importacion de Datos"
      End
      Begin VB.Menu Submenu_Adm_Control_Bolsa_Horas 
         Caption         =   "Control Bolsa de Horas"
      End
   End
   Begin VB.Menu Menu_Adm_Comedor 
      Caption         =   "C&omedor"
      Begin VB.Menu Submenu_Adm_Entradas_Comedor 
         Caption         =   "Entradas a Comedor"
      End
      Begin VB.Menu Submenu_Rpt_Entradas_Comedor 
         Caption         =   "Reporte de Comedor"
      End
      Begin VB.Menu Submenu_Rpt_Empleados_Huella 
         Caption         =   "Reporte de Empleados con Huella"
      End
   End
   Begin VB.Menu Menu_Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta 
         Caption         =   "Reporte de Empleados"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Asistencias 
         Caption         =   "Asistencias"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos 
         Caption         =   "Histórico de Faltas"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos 
         Caption         =   "Histórico de Permisos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado 
         Caption         =   "Horas Trabajadas por Empleado"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados 
         Caption         =   "No de Empleados no Validados"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja 
         Caption         =   "Empleados con Baja"
      End
      Begin VB.Menu Submenu_Rpt_Cursos_Empleado 
         Caption         =   "Reporte de Cursos"
      End
      Begin VB.Menu Submenu_Rpt_Cursos_Por_Empleado 
         Caption         =   "Cursos Tomados por Empleado"
      End
      Begin VB.Menu Submenu_Rpt_Cursos_Hora_Hombre 
         Caption         =   "Cursos Hora Hombre"
      End
      Begin VB.Menu Submenu_Rpt_Cursos_Indice_Asistencia 
         Caption         =   "Cursos Indice de Asistencia"
      End
      Begin VB.Menu Submenu_Rpt_Cursos_Resumen_Mensual 
         Caption         =   "Cursos Resumen Mensual"
      End
      Begin VB.Menu Submenu_Rpt_General_Cursos 
         Caption         =   "Reporte General de Cursos"
      End
      Begin VB.Menu Submenu_Rpt_Historico_Vacaciones 
         Caption         =   "Histórico Vacaciones"
      End
      Begin VB.Menu Submenu_Rpt_No_Checadas 
         Caption         =   "Reporte No Checadas"
      End
      Begin VB.Menu SubMenu_Btn_Adm_Rpt_Accesos_Almacenes 
         Caption         =   "Reporte Accesos al Almacen"
      End
   End
   Begin VB.Menu Menu_Catalogos 
      Caption         =   "&Catálogos"
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Salas 
         Caption         =   "Salas"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Instituciones 
         Caption         =   "Instituciones"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Instructores 
         Caption         =   "Instructores"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Tipos_Cursos 
         Caption         =   "Tipos Cursos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Cursos 
         Caption         =   "Cursos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Empresas 
         Caption         =   "Empresas"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Departamentos 
         Caption         =   "Departamentos"
      End
      Begin VB.Menu Submenu_Cat_Areas 
         Caption         =   "Areas"
      End
      Begin VB.Menu Submenu_Cat_Gaps 
         Caption         =   "Tripulaciones"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Puestos 
         Caption         =   "Puestos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio 
         Caption         =   "Niveles de Estudio"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja 
         Caption         =   "Motivos de Baja"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Turnos 
         Caption         =   "Turnos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Calendarios_Turnos 
         Caption         =   "Calendarios Turnos"
      End
      Begin VB.Menu Submenu_Cat_Zonas 
         Caption         =   "Zonas"
      End
      Begin VB.Menu Submenu_Cat_Transportes 
         Caption         =   "Transportes"
      End
      Begin VB.Menu Submenu_Cat_Secciones 
         Caption         =   "Secciones"
      End
      Begin VB.Menu Submenu_Cat_Cursos 
         Caption         =   "Cursos"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Empleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu Raya2 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias 
         Caption         =   "Incidencias Extraordinaras"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales 
         Caption         =   "Días no Laborables"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion 
         Caption         =   "Relojes Checadores"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Equipo_Almacenes_Identificacion 
         Caption         =   "Relojes Checadores Almacenes"
      End
      Begin VB.Menu Raya3 
         Caption         =   "-"
      End
      Begin VB.Menu Submenu_Btn_Adm_RH_Panel_Cat_Parametros 
         Caption         =   "Parámetros"
      End
      Begin VB.Menu Submenu_Catalogo_Roles 
         Caption         =   "Roles"
      End
      Begin VB.Menu Submenu_Usuarios 
         Caption         =   "Usuarios"
      End
   End
   Begin VB.Menu Menu_Cursos_Capacitaciones 
      Caption         =   "Cursos y Capacitaciones"
      Begin VB.Menu Submenu_Programación_Cursos 
         Caption         =   "Programación de Cursos"
      End
      Begin VB.Menu Submenu_Registro_Asistencias 
         Caption         =   "Registro de Asistencias"
      End
      Begin VB.Menu Submenu_Registro_Calificaciones 
         Caption         =   "Registro de Calificaciones"
      End
      Begin VB.Menu Submenu_Evaluacion_Cursos 
         Caption         =   "Registrar Evaluación de Cursos"
      End
   End
   Begin VB.Menu Menu_Acerca_De 
      Caption         =   "Acerca De..."
      Begin VB.Menu Submenu_Sistema 
         Caption         =   "Sistema"
      End
   End
End
Attribute VB_Name = "MDIFrm_Apl_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Instituciones_Click(Index As Integer)

End Sub

Private Sub Menu_Cursos_Click()

End Sub

Private Sub Submenu_Adm_Cambio_Turno_Click()
Dim Frm_Cambio_Turno As New Frm_Adm_Cambio_Turno
    
    If Conectar_Ayudante.Formulario_Cargado("CAMBIO DE TURNO") Then
        Conectar_Ayudante.Enfocar ("CAMBIO DE TURNO")
    Else
        Load Frm_Cambio_Turno
        Frm_Cambio_Turno.Height = 4600
        Frm_Cambio_Turno.Width = 7180
        Call Conectar_Ayudante.Cargar_Picture(Frm_Cambio_Turno.Pic_Adm_Validacion_Horas_Trabajo, Frm_Cambio_Turno)
        Frm_Cambio_Turno.Caption = "CAMBIO DE TURNO"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Adm_Cambio_Turno", Frm_Cambio_Turno)
    End If
End Sub

Private Sub Submenu_Adm_Control_Bolsa_Horas_Click()
    Unload Frm_Adm_Bolsa_Horas
    Load Frm_Adm_Bolsa_Horas
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Adm_Bolsa_Horas", Frm_Adm_Bolsa_Horas)
End Sub

Private Sub Submenu_Adm_Entradas_Comedor_Click()
    Unload Frm_Adm_Entrada_Comedor
    Load Frm_Adm_Entrada_Comedor
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Adm_Entradas_Comedor", Frm_Adm_Entrada_Comedor)
End Sub

Private Sub Submenu_Apl_Cambio_Password_Click()
    Unload Frm_Apl_Login
    Load Frm_Apl_Cambio_Password
End Sub

Private Sub MDIForm_Load()
Dim Dia As Integer

On Error GoTo HANDLER
    Set Conectar_Ayudante = New Ayudante
    'Revisa el formato de la fecha
    Dia = Day(Now)
    If Val(Dia) <> Val(Mid(Now, 4, 2)) Then
        MsgBox "Formato de fecha corta incorrecta en su PC tiene actualmente Dia/Mes/Año" & Chr(13) & Chr(13) & " Para que el sistema funcione correctamente cambiarla a Mes/Dia/Año"
        End
    End If
    Conectar_Ayudante.Conexion 'Manda llamar a la conexión a la base desde Ayudante
    StatusBar.Panels(1) = Format(Now, "dd/MMM/yyyy")
    ''Ejecuta el programa para descarga de información
    'RETVAL = Shell(App.Path & "\SRG_Sincronizacion.exe", vbMinimizedFocus)
Exit Sub
HANDLER:
End Sub

Private Sub Submenu_Apl_Respaldo_Sistema_Click()
Dim Nombre_Respaldo As String             'Indica el nombre del respaldo que se va hacer generado
Dim Destino_Respaldo_Base_Datos As String 'Indica el destino en donde se esta generando el respaldo en la base de datos
Dim Ruta_Archivo_Destino As String        'Indica la ruta en que será guardado el respaldo y que fue dada por el usuario
Dim linea As String

On Error GoTo ErrHandler
    If MsgBox("¿Desea realizar el respaldo de su base de datos?", vbYesNo + vbQuestion) = vbYes Then
        MDIFrm_Apl_Principal.MousePointer = 11
    
            Destino_Respaldo_Base_Datos = App.Path & "\Respaldos\"
            'Si no se encuantra la carpeta la crea para guardar los respaldos
            If Not CBool(PathIsDirectory(Destino_Respaldo_Base_Datos)) Then
                'crea la carpeta
                MkDir Destino_Respaldo_Base_Datos
            End If
                    
            Nombre_Respaldo = "BK_SRG_RH_" & Format(Now, "yyyy-MM-dd-HHmm") & ".bak"
            'Si el archivo del respaldo ya fue creado entonces lo elimina para poder agregar el nuevo respaldo
            If (Len(Dir(Destino_Respaldo_Base_Datos & Nombre_Respaldo))) > 0 Then
                Kill (Destino_Respaldo_Base_Datos & Nombre_Respaldo)
            End If
            
            Conectar_Ayudante.Conexion_Respaldo 'Se conecta a la base de datos para poder realizar el respaldo de la misma
            'Realiza el respaldo de la base de datos
            Conexion_Base_Respaldo.Execute "USE [" & Database & "]"
            Mi_SQL = "BACKUP DATABASE " & Database & _
                                  " TO DISK='" & Destino_Respaldo_Base_Datos & Nombre_Respaldo & "'" & _
                                  " WITH FORMAT, NAME = 'Full Backup of " & Database & "'"
            Conexion_Base_Respaldo.Execute Mi_SQL
            Conexion_Base_Respaldo.Close
            
            MsgBox "El respaldo de la base de datos fue creado con éxito", vbInformation
        MDIFrm_Apl_Principal.MousePointer = 0
    End If
    Exit Sub
ErrHandler:
    Conexion_Base_Respaldo.Close
    MDIFrm_Apl_Principal.MousePointer = 0
    MsgBox Err.Description
    MsgBox "Ocurrio un error al copiar el archivo, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados_Click()
Dim Frm_Asistencias_Empleados As New Frm_Adm_Asistencias_Empleados
    
    If Conectar_Ayudante.Formulario_Cargado("ASISTENCIAS EMPLEADOS") Then
        Conectar_Ayudante.Enfocar ("ASISTENCIAS EMPLEADOS")
    Else
        Load Frm_Asistencias_Empleados
        Frm_Asistencias_Empleados.Operacion = "Asistencias_Empleados"
        Frm_Asistencias_Empleados.Height = 4335
        Frm_Asistencias_Empleados.Width = 6225
        Call Conectar_Ayudante.Cargar_Picture(Frm_Asistencias_Empleados.Pic_Adm_Asistencias_Empleados_Consulta, Frm_Asistencias_Empleados)
        Frm_Asistencias_Empleados.Caption = "ASISTENCIAS EMPLEADOS"
        Frm_Asistencias_Empleados.Inicializar
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Asistencia_Empleados", Frm_Asistencias_Empleados)
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Calendarios_Turnos_Click()
Dim Frm_Calendarios_Turnos As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE CALENDARIOS TURNOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE CALENDARIOS TURNOS")
    Else
        Load Frm_Calendarios_Turnos
        Frm_Calendarios_Turnos.Catalogo = "Cat_Calendarios_Turnos"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Calendarios_Turnos.Pic_Cat_Calendarios_Turnos, Frm_Calendarios_Turnos)
        Frm_Calendarios_Turnos.Caption = "CATALOGO DE CALENDARIOS TURNOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Btn_Adm_RH_Panel_Cat_Calendarios_Turnos", Frm_Calendarios_Turnos)
        Frm_Calendarios_Turnos.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Cursos_Click()
Dim Frm_Cursos As New Frm_Cat_Cursos

'    Load Frm_Instituciones
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE CURSOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE CURSOS")
    Else
        Load Frm_Cursos
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Cursos.Caption = "CATALOGO DE CURSOS"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Cursos.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos_Click()
Dim Frm_Departamentos As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE DEPARTAMENTOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE DEPARTAMENTOS")
    Else
        Load Frm_Departamentos
        Frm_Departamentos.Catalogo = "Cat_Departamentos"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Departamentos.Pic_Cat_Departamentos, Frm_Departamentos)
        Frm_Departamentos.Caption = "CATALOGO DE DEPARTAMENTOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Departamentos", Frm_Departamentos)
        Frm_Departamentos.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales_Click()
Dim Frm_Dias_No_Laborales As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE DIAS NO LABORALES") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE DIAS NO LABORALES")
    Else
        Load Frm_Dias_No_Laborales
        Frm_Dias_No_Laborales.Catalogo = "Cat_Dias_No_Laborales"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Dias_No_Laborales.Pic_Cat_Dias_No_Laborales, Frm_Dias_No_Laborales)
        Frm_Dias_No_Laborales.Caption = "CATALOGO DE DIAS NO LABORALES"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Dias_No_Laborales", Frm_Dias_No_Laborales)
        Frm_Dias_No_Laborales.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Empleados_Click()
Dim Frm_Empleados As New Frm_Cat_Empleados
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE EMPLEADOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE EMPLEADOS")
    Else
        Load Frm_Empleados
        Frm_Empleados.Catalogo = "Cat_Empleados"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Empleados.Pic_Cat_Empleados, Frm_Empleados)
        Frm_Empleados.Caption = "CATALOGO DE EMPLEADOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empleados", Frm_Empleados)
        Frm_Empleados.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Empresas_Click()
Dim Frm_Empresas As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE EMPRESAS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE EMPRESAS")
    Else
        Load Frm_Empresas
        Frm_Empresas.Catalogo = "Cat_Empresas"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Empresas.Caption = "CATALOGO DE EMPRESAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Empresas.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion_Click()
Dim Frm_Equipos_Identificacion As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE EQUIPOS DE IDENTIFICACION") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE EQUIPOS DE IDENTIFICACION")
    Else
        Load Frm_Equipos_Identificacion
        Frm_Equipos_Identificacion.Catalogo = "Cat_Equipos_Identificacion"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Equipos_Identificacion.Pic_Cat_Equipos, Frm_Equipos_Identificacion)
        Frm_Equipos_Identificacion.Caption = "CATALOGO DE EQUIPOS DE IDENTIFICACION"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Identificacion", Frm_Equipos_Identificacion)
        Frm_Equipos_Identificacion.Inicializa
    End If
End Sub
Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Almacenes_Identificacion_Click()
Dim Frm_Equipos_Almacenes_Identificacion As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE EQUIPOS DE IDENTIFICACION ALMACENES") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE EQUIPOS DE IDENTIFICACION ALMACENES")
    Else
        Load Frm_Equipos_Almacenes_Identificacion
        Frm_Equipos_Almacenes_Identificacion.Catalogo = "Cat_Equipos_Almacenes_Identificacion"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Equipos_Almacenes_Identificacion.Pic_Cat_Equipos_Almacenes, Frm_Equipos_Almacenes_Identificacion)
        Frm_Equipos_Almacenes_Identificacion.Caption = "CATALOGO DE EQUIPOS DE IDENTIFICACION ALMACENES"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Equipo_Almacenes_Identificacion", Frm_Equipos_Almacenes_Identificacion)
        Frm_Equipos_Almacenes_Identificacion.Inicializa
    End If
End Sub
Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias_Click()
Dim Frm_Faltas_Extraordinarias As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE INCIDENCIAS EXTRAORDINARIAS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE INCIDENCIAS EXTRAORDINARIAS")
    Else
        Load Frm_Faltas_Extraordinarias
        Frm_Faltas_Extraordinarias.Catalogo = "Cat_Tipos_Faltas"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Faltas_Extraordinarias.Pic_Cat_Tipos_Faltas, Frm_Faltas_Extraordinarias)
        Frm_Faltas_Extraordinarias.Caption = "CATALOGO DE INCIDENCIAS EXTRAORDINARIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Incidencias_Extraordinarias", Frm_Faltas_Extraordinarias)
        Frm_Faltas_Extraordinarias.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Instituciones_Click()
Dim Frm_Instituciones As New Frm_Cat_Instituciones
    
'    Load Frm_Instituciones
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE INSTITUCIONES") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE INSTITUCIONES")
    Else
        Load Frm_Instituciones
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Instituciones.Caption = "CATALOGO DE INSTITUCIONES"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Instituciones.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Instructores_Click()
Dim Frm_Instructores As New Frm_Cat_Instructores

'    Load Frm_Instituciones
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE INTRUSCTORES") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE INTRUCTORES")
    Else
        Load Frm_Instructores
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Instructores.Caption = "CATALOGO DE INTRUCTORES"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Instructores.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja_Click()
Dim Frm_Motivos_Baja As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE MOTIVOS DE BAJA") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE MOTIVOS DE BAJA")
    Else
        Load Frm_Motivos_Baja
        Frm_Motivos_Baja.Catalogo = "Cat_Motivos_Baja"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Motivos_Baja.Pic_Cat_Motivos_Baja, Frm_Motivos_Baja)
        Frm_Motivos_Baja.Caption = "CATALOGO DE MOTIVOS DE BAJA"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Motivos_Baja", Frm_Motivos_Baja)
        Frm_Motivos_Baja.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio_Click()
Dim Frm_Nivel_Estudio As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE NIVELES DE ESTUDIO") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE NIVELES DE ESTUDIO")
    Else
        Load Frm_Nivel_Estudio
        Frm_Nivel_Estudio.Catalogo = "Cat_Nivel_Estudio"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Nivel_Estudio.Pic_Cat_Nivel_Estudio, Frm_Nivel_Estudio)
        Frm_Nivel_Estudio.Caption = "CATALOGO DE NIVELES DE ESTUDIO"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Nivel_Estudio", Frm_Nivel_Estudio)
        Frm_Nivel_Estudio.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Parametros_Click()
Dim Frm_Parametros_RH As New Frm_Cat_Parametros_RH
    
    If Conectar_Ayudante.Formulario_Cargado("PARAMETROS") Then
        Conectar_Ayudante.Enfocar ("PARAMETROS")
    Else
        Load Frm_Parametros_RH
        'Frm_Parametros_RH.Catalogo = "Cat_Parametros"
        'Call Conectar_Ayudante.Cargar_Picture(Frm_Parametros_RH.Pic_Cat_Puestos, Frm_Parametros_RH)
        Frm_Parametros_RH.Caption = "PARAMETROS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Parametros", Frm_Parametros_RH)
        'Frm_Parametros_RH.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Puestos_Click()
Dim Frm_Puestos As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE PUESTOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE PUESTOS")
    Else
        Load Frm_Puestos
        Frm_Puestos.Catalogo = "Cat_Puestos"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Puestos.Pic_Cat_Puestos, Frm_Puestos)
        Frm_Puestos.Caption = "CATALOGO DE PUESTOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Puestos", Frm_Puestos)
        Frm_Puestos.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Salas_Click()
Dim Frm_Salas As New Frm_Cat_Salas

'    Load Frm_Instituciones
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE SALAS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE SALAS")
    Else
        Load Frm_Salas
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Salas.Caption = "CATALOGO DE SALAS"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Salas.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Cat_Tipos_Cursos_Click()
Dim Frm_Tipos_Cursos As New Frm_Cat_Tipos_Curos

'    Load Frm_Instituciones
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE TIPOS CURSOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE TIPOS CURSOS")
    Else
        Load Frm_Tipos_Cursos
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Tipos_Cursos.Caption = "CATALOGO DE TIPOS CURSOS"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Tipos_Cursos.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Cat_Turnos_Click()
Dim Frm_Turnos As New Frm_Cat_Generales_RH
    
    If Conectar_Ayudante.Formulario_Cargado("CATALOGO DE TURNOS") Then
        Conectar_Ayudante.Enfocar ("CATALOGO DE TURNOS")
    Else
        Load Frm_Turnos
        Frm_Turnos.Catalogo = "Cat_Turnos"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Turnos.Pic_Cat_Turnos, Frm_Turnos)
        Frm_Turnos.Caption = "CATALOGO DE TURNOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Turnos", Frm_Turnos)
        Frm_Turnos.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Correo_Validacion_Click()
Dim Frm_Correos As New Frm_Adm_Envio_Correo_Validacion
    
    If Conectar_Ayudante.Formulario_Cargado("CORREO DE VALIDACION") Then
        Conectar_Ayudante.Enfocar ("CORREO DE VALIDACION")
    Else
        Load Frm_Correos
        Frm_Correos.Caption = "CORREO DE VALIDACION"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Correo_Validacion", Frm_Correos)
        Frm_Correos.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Control_Calzado_Click()
        Load Frm_Adm_Control_Calzado_Empleado
        Frm_Adm_Control_Calzado_Empleado.Inicializa
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias_Click()
Dim Frm_Importacion_Asistencias As New Frm_Adm_Importacion
    
    If Conectar_Ayudante.Formulario_Cargado("IMPORTACION ASISTENCIAS") Then
        Conectar_Ayudante.Enfocar ("IMPORTACION ASISTENCIAS")
    Else
        Load Frm_Importacion_Asistencias
        Frm_Importacion_Asistencias.Opcion = "Importacion_Asistencias"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Importacion_Asistencias.Pic_Importacion_Keri_System, Frm_Importacion_Asistencias)
        Frm_Importacion_Asistencias.Caption = "INCIDENCIAS EXTRAORDINARIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias", Frm_Importacion_Asistencias)
        Frm_Importacion_Asistencias.Inicializa
    End If
End Sub
Private Sub SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias_Almacenes_Click()
Dim Frm_Importacion_Asistencias_Almacenes As New Frm_Adm_Importacion_Almacenes
    
    If Conectar_Ayudante.Formulario_Cargado("IMPORTACION ASISTENCIAS ALMACENES") Then
        Conectar_Ayudante.Enfocar ("IMPORTACION ASISTENCIAS ALMACENES")
    Else
        Load Frm_Importacion_Asistencias_Almacenes
        Frm_Importacion_Asistencias_Almacenes.Opcion = "Importacion_Asistencias_Almacenes"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Importacion_Asistencias_Almacenes.Pic_Importacion_Keri_System, Frm_Importacion_Asistencias_Almacenes)
        Frm_Importacion_Asistencias_Almacenes.Caption = "INCIDENCIAS EXTRAORDINARIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Importacion_Asistencias_Almacenes", Frm_Importacion_Asistencias_Almacenes)
        Frm_Importacion_Asistencias_Almacenes.Inicializa
    End If
End Sub
Private Sub SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias_Click()
Dim Frm_Incidencias As New Frm_Adm_Incidencias_Extraordinarias
    
    If Conectar_Ayudante.Formulario_Cargado("INCIDENCIAS EXTRAORDINARIAS") Then
        Conectar_Ayudante.Enfocar ("INCIDENCIAS EXTRAORDINARIAS")
    Else
        Load Frm_Incidencias
        'Frm_Permisos.Height = 3510
        'Frm_Permisos.Width = 7080
        Frm_Incidencias.Top = 0
        Call Conectar_Ayudante.Cargar_Picture(Frm_Incidencias.Pic_Solicitud_Permisos, Frm_Incidencias)
        Frm_Incidencias.Operacion = "Permisos"
        Frm_Incidencias.Pic_Logo.Visible = True
        Frm_Incidencias.Pic_Logo.ZOrder vbBringToFront
        Frm_Incidencias.Caption = "INCIDENCIAS EXTRAORDINARIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Incidencias_Extraordinarias", Frm_Incidencias)
        Frm_Incidencias.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias_Click()
Dim Frm_Mantenimiento As New Frm_Adm_Asistencias_Mantenimiento
    
    If Conectar_Ayudante.Formulario_Cargado("MANTENIMIENTO ASISTENCIAS") Then
        Conectar_Ayudante.Enfocar ("MANTENIMIENTO ASISTENCIAS")
    Else
        Load Frm_Mantenimiento
        'Frm_Mantenimiento.Operacion = "Mantenimiento_Asistencias"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Mantenimiento.Pic_Adm_Asistencias_Consulta, Frm_Mantenimiento)
        Frm_Mantenimiento.Caption = "MANTENIMIENTO ASISTENCIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Mantenimiento_Asistencias", Frm_Mantenimiento)
        Frm_Mantenimiento.Inicializa
    End If
End Sub

Private Sub Submenu_Btn_Adm_RH_Panel_Notificación_Aniversarios_Click()

        Load Frm_Adm_Notificaciones_Aniversarios
       Frm_Adm_Notificaciones_Aniversarios.Inicializa
'    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias_Click()
Dim Frm_Rpt_Asistencias As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("ASISTENCIAS DE EMPLEADOS") Then
        Conectar_Ayudante.Enfocar ("ASISTENCIAS DE EMPLEADOS")
    Else
        Load Frm_Rpt_Asistencias
        Frm_Rpt_Asistencias.Caption = "ASISTENCIAS DE EMPLEADOS"
        Frm_Rpt_Asistencias.Reporte = "Asistencias_Empleados"
        Call Frm_Rpt_Asistencias.Cargar_Frame(Frm_Rpt_Asistencias.Fra_Rpt_Asistencia_Empleados, Frm_Rpt_Asistencias)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Asistencias", Frm_Rpt_Asistencias)
        Frm_Rpt_Asistencias.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta_Click()
Dim Frm_Rpt_Empleados_Alta As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("EMPLEADOS EN ALTA") Then
        Conectar_Ayudante.Enfocar ("EMPLEADOS EN ALTA")
    Else
        Load Frm_Rpt_Empleados_Alta
        Frm_Rpt_Empleados_Alta.Caption = "EMPLEADOS EN ALTA"
        Frm_Rpt_Empleados_Alta.Reporte = "Empleados_Alta"
        Call Frm_Rpt_Empleados_Alta.Cargar_Frame(Frm_Rpt_Empleados_Alta.Fra_Rpt_Empleados_Alta, Frm_Rpt_Empleados_Alta)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Alta", Frm_Rpt_Empleados_Alta)
        Frm_Rpt_Empleados_Alta.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja_Click()
Dim Frm_Rpt_Empleados_Baja As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("EMPLEADOS EN BAJA") Then
        Conectar_Ayudante.Enfocar ("EMPLEADOS EN BAJA")
    Else
        Load Frm_Rpt_Empleados_Baja
        Frm_Rpt_Empleados_Baja.Caption = "EMPLEADOS EN BAJA"
        Frm_Rpt_Empleados_Baja.Reporte = "Empleados_Baja"
        Call Frm_Rpt_Empleados_Baja.Cargar_Frame(Frm_Rpt_Empleados_Baja.Fra_Rpt_Empleados_Baja, Frm_Rpt_Empleados_Baja)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_Baja", Frm_Rpt_Empleados_Baja)
        Frm_Rpt_Empleados_Baja.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados_Click()
Dim Frm_Rpt_Empleados_No_Validados As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("EMPLEADOS NO VALIDADOS") Then
        Conectar_Ayudante.Enfocar ("EMPLEADOS NO VALIDADOS")
    Else
        Load Frm_Rpt_Empleados_No_Validados
        Frm_Rpt_Empleados_No_Validados.Caption = "EMPLEADOS NO VALIDADOS"
        Frm_Rpt_Empleados_No_Validados.Reporte = "Empleados_No_Validados"
        Call Frm_Rpt_Empleados_No_Validados.Cargar_Frame(Frm_Rpt_Empleados_No_Validados.Fra_Rpt_Empleados_No_Validados, Frm_Rpt_Empleados_No_Validados)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Empleados_No_Validados", Frm_Rpt_Empleados_No_Validados)
        Frm_Rpt_Empleados_No_Validados.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos_Click()
Dim Frm_Rpt_Faltas_Retardos As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("HISTORICO DE FALTAS") Then
        Conectar_Ayudante.Enfocar ("HISTORICO DE FALTAS")
    Else
        Load Frm_Rpt_Faltas_Retardos
        Frm_Rpt_Faltas_Retardos.Caption = "HISTORICO DE FALTAS"
        Frm_Rpt_Faltas_Retardos.Reporte = "Historico_Faltas_Retardos"
        Call Frm_Rpt_Faltas_Retardos.Cargar_Frame(Frm_Rpt_Faltas_Retardos.Fra_Rpt_Historico_Faltas_Retardos, Frm_Rpt_Faltas_Retardos)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Faltas_Retardos", Frm_Rpt_Faltas_Retardos)
        Frm_Rpt_Faltas_Retardos.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos_Click()
Dim Frm_Rpt_Permisos As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("HISTORICO DE PERMISOS") Then
        Conectar_Ayudante.Enfocar ("HISTORICO DE PERMISOS")
    Else
        Load Frm_Rpt_Permisos
        Frm_Rpt_Permisos.Caption = "HISTORICO DE PERMISOS"
        Frm_Rpt_Permisos.Reporte = "Historico_Permisos"
        Call Frm_Rpt_Permisos.Cargar_Frame(Frm_Rpt_Permisos.Fra_Rpt_Historico_Permisos, Frm_Rpt_Permisos)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Historico_Permisos", Frm_Rpt_Permisos)
        Frm_Rpt_Permisos.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado_Click()
Dim Frm_Rpt_Horas_Trabajas_Empleados As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("HORAS TRABAJADAS EMPLEADOS") Then
        Conectar_Ayudante.Enfocar ("HORAS TRABAJADAS EMPLEADOS")
    Else
        Load Frm_Rpt_Horas_Trabajas_Empleados
        Frm_Rpt_Horas_Trabajas_Empleados.Caption = "HORAS TRABAJADAS EMPLEADOS"
        Frm_Rpt_Horas_Trabajas_Empleados.Reporte = "Horas_Trabajadas_Empleado"
        Call Frm_Rpt_Horas_Trabajas_Empleados.Cargar_Frame(Frm_Rpt_Horas_Trabajas_Empleados.Fra_Rpt_Horas_Trabajadas_Empleado, Frm_Rpt_Horas_Trabajas_Empleados)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Rpt_Horas_Trabajadas_Empleado", Frm_Rpt_Horas_Trabajas_Empleados)
        Frm_Rpt_Horas_Trabajas_Empleados.Inicializar
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos_Click()
Dim Frm_Permisos As New Frm_Adm_Solicitud_Permisos
    
    If Conectar_Ayudante.Formulario_Cargado("SOLICITUD DE PERMISOS") Then
        Conectar_Ayudante.Enfocar ("SOLICITUD DE PERMISOS")
    Else
        Load Frm_Permisos
        'Frm_Permisos.Height = 3510
        'Frm_Permisos.Width = 7080
        Frm_Permisos.Top = 0
        Call Conectar_Ayudante.Cargar_Picture(Frm_Permisos.Pic_Solicitud_Permisos, Frm_Permisos)
        Frm_Permisos.Operacion = "Permisos"
        Frm_Permisos.Caption = "SOLICITUD DE PERMISOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Solicitud_Permisos", Frm_Permisos)
        Frm_Permisos.Inicializa
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo_Click()
Dim Frm_Validacion As New Frm_Adm_Validación_Tiempo_Trabajo
    
    If Conectar_Ayudante.Formulario_Cargado("VALIDACION HORAS TRABAJO") Then
        Conectar_Ayudante.Enfocar ("VALIDACION HORAS TRABAJO")
    Else
        Load Frm_Validacion
        Frm_Validacion.Opcion = "Validacion_Tiempo"
        Frm_Validacion.Height = 4300
        Frm_Validacion.Width = 7180
        Call Conectar_Ayudante.Cargar_Picture(Frm_Validacion.Pic_Adm_Validacion_Horas_Trabajo, Frm_Validacion)
        Frm_Validacion.Caption = "VALIDACION HORAS TRABAJO"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Validacion_Tiempo_Trabajo", Frm_Validacion)
    End If
End Sub

Private Sub SubMenu_Btn_Adm_RH_Panel_Visor_Registros_Click()
Dim Frm_Visor As New Frm_Adm_Visor_Asistencias
    
    If Conectar_Ayudante.Formulario_Cargado("VISOR DE ASISTENCIAS") Then
        Conectar_Ayudante.Enfocar ("VISOR DE ASISTENCIAS")
    Else
        Load Frm_Visor
        Frm_Visor.Caption = "VISOR DE ASISTENCIAS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Visor_Registros", Frm_Visor)
        Frm_Visor.Inicializa
    End If
End Sub

Private Sub Submenu_Cat_Cursos_Click()
    Catalogo = "CURSOS"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Cursos"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Cursos, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Cursos("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Cursos", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Areas_Click()
    Unload Frm_Cat_Areas
    Load Frm_Cat_Areas
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Areas", Frm_Cat_Areas)
End Sub

Private Sub Submenu_Cat_Gaps_Click()
    Catalogo = "GAPS"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Tripulaciones"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Gaps, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Gaps("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Gaps", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Gerencias_UAP_Click()
    Catalogo = "GERENCIAS"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Gerencias UAP"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Gerencias, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Gerencia("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Gerencias_UAP", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Secciones_Click()
    Catalogo = "SECCIONES"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Secciones"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Secciones, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Secciones("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Secciones", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Transportes_Click()
    Catalogo = "TRANSPORTES"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Transportes"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Transportes, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Transportes("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Transportes", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Cat_Zonas_Click()
    Catalogo = "ZONAS"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Zonas"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Zonas, Frm_Cat_Generales
    Call Frm_Cat_Generales.Consulta_Zonas("")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Zonas", Frm_Cat_Generales)
End Sub

Private Sub SubMenu_Catalogo_Roles_Click()
    'Le asigna la palabra usuarios para manejar el picture del catálogo usuarios
    Catalogo = "ROLES"
    Unload Frm_Cat_Generales
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Roles"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    'Carga el picture de Usuarios
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Apl_Cat_Roles, Frm_Cat_Generales
    'Llama a la función de consulta usuarios para llenar el grid del catálogo
    Call Frm_Cat_Generales.Consulta_Roles("%")
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Catalogo_Roles", Frm_Cat_Generales)
End Sub

Private Sub Submenu_Formato_Click()
    Frm_Cfg_Impresion.Show
End Sub


Private Sub Submenu_Evaluacion_Cursos_Click()
Dim Frm_Eval_Cursos As New Frm_Ope_Registrar_Evaluacion_Cursos

    If Conectar_Ayudante.Formulario_Cargado("REGISTRAR EVALUACION DE CURSOS") Then
        Conectar_Ayudante.Enfocar ("REGISTRAR EVALUACION DE CURSOS")
    Else
        Load Frm_Eval_Cursos
        Frm_Eval_Cursos.Caption = "REGISTRAR EVALUACION DE CURSOS"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Evaluacion_Cursos", Frm_Eval_Cursos)
        Frm_Eval_Cursos.Inicializa
    End If
End Sub

Private Sub Submenu_Ope_Importacion_Archivo_Click()
    Unload Frm_Ope_Importacion_Datos
    Load Frm_Ope_Importacion_Datos
    Frm_Ope_Importacion_Datos.Pic_Actualizar_Productos.Visible = True
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Ope_Importacion_Archivo", Frm_Ope_Importacion_Datos)
End Sub

Private Sub Submenu_Programación_Cursos_Click()
'Dim Frm_Programacion_Cursos As New Frm_Ope_Programacion_Cursos

'    Load Frm_Instituciones
'    If Conectar_Ayudante.Formulario_Cargado("PROGRAMACIÓN DE CURSOS") Then
'        Conectar_Ayudante.Enfocar ("PROGRAMACIÓN DE CURSOS")
'    Else
        Load Frm_Ope_Programacion_Cursos
'        Frm_Instituciones.Catalogo = "Cat_Instituciones"
'        Call Conectar_Ayudante.Cargar_Picture(Frm_Empresas.Pic_Cat_Empresas, Frm_Empresas)
        Frm_Ope_Programacion_Cursos.Caption = "PROGRAMACIÓN DE CURSOS"
'        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_RH_Panel_Cat_Empresas", Frm_Empresas)
        Frm_Ope_Programacion_Cursos.Inicializa
'    End If
End Sub

Private Sub Submenu_Registrarse_Click()
    Tipo_Validacion = False
    Frm_Apl_Login.Show
End Sub

Private Sub Submenu_Registro_Asistencias_Click()
 Load Frm_Ope_Asistencia_Cursos
 'Frm_Ope_Asistencia_Cursos.Caption = "PROGRAMACIÓN DE CURSOS"
 Frm_Ope_Asistencia_Cursos.Inicializa
End Sub

Private Sub Submenu_Registro_Calificaciones_Click()
Dim Frm_Reg_Calif As New Frm_Ope_Registro_Calificaciones
    
    If Conectar_Ayudante.Formulario_Cargado("REGISTRO DE CALIFICACIONES") Then
        Conectar_Ayudante.Enfocar ("REGISTRO DE CALIFICACIONES")
    Else
        Load Frm_Reg_Calif
        Frm_Reg_Calif.Caption = "REGISTRO DE CALIFICACIONES"
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Registro_Calificaciones", Frm_Reg_Calif)
        Frm_Reg_Calif.Inicializa
    End If
End Sub

Private Sub Submenu_Rpt_Cursos_Empleado_Click()
Dim Frm_Rpt_Cursos As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("REPORTE DE CURSOS") Then
        Conectar_Ayudante.Enfocar ("REPORTE DE CURSOS")
    Else
        Load Frm_Rpt_Cursos
        Frm_Rpt_Cursos.Caption = "REPORTE DE CURSOS"
        Frm_Rpt_Cursos.Reporte = "Reporte_Curso"
        Call Frm_Rpt_Cursos.Cargar_Frame(Frm_Rpt_Cursos.Fra_Reporte_Cursos, Frm_Rpt_Cursos)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Empleado", Frm_Rpt_Cursos)
        Frm_Rpt_Cursos.Inicializar
        Frm_Rpt_Cursos.Dtp_Fecha_Inicio_Curso.Value = Now
        Frm_Rpt_Cursos.Dtp_Fecha_Fin_Curso.Value = Now
    End If
End Sub

Private Sub Submenu_Rpt_Faltas_Empleados_Click()
Dim Frm_Rpt_Faltas_Empleados As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("FALTAS DEL DIA") Then
        Conectar_Ayudante.Enfocar ("FALTAS DEL DIA")
    Else
        Load Frm_Rpt_Faltas_Empleados
        Frm_Rpt_Faltas_Empleados.Caption = "FALTAS DEL DIA"
        Frm_Rpt_Faltas_Empleados.Reporte = "Empleados_Faltas"
        Call Frm_Rpt_Faltas_Empleados.Cargar_Frame(Frm_Rpt_Faltas_Empleados.Fra_Faltas_Empleados, Frm_Rpt_Faltas_Empleados)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Faltas_Empleados", Frm_Rpt_Faltas_Empleados)
        Frm_Rpt_Faltas_Empleados.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Cursos_Hora_Hombre_Click()
Dim Frm_Rpt_Cursos_Hora_Hombre As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("CURSOS HORA HOMBRE") Then
        Conectar_Ayudante.Enfocar ("CURSOS HORA HOMBRE")
    Else
        Load Frm_Rpt_Cursos_Hora_Hombre
        Frm_Rpt_Cursos_Hora_Hombre.Caption = "CURSOS HORA HOMBRE"
        Frm_Rpt_Cursos_Hora_Hombre.Reporte = "Cursos_Hora_Hombre"
        Frm_Rpt_Cursos_Hora_Hombre.Btn_Exportar_PDF.Enabled = True
        Frm_Rpt_Cursos_Hora_Hombre.Btn_Exportar_PDF.Visible = True
        Call Frm_Rpt_Cursos_Hora_Hombre.Cargar_Frame(Frm_Rpt_Cursos_Hora_Hombre.Fra_Cursos_Horas_Hombre, Frm_Rpt_Cursos_Hora_Hombre)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Hora_Hombre", Frm_Rpt_Cursos_Hora_Hombre)
        Frm_Rpt_Cursos_Hora_Hombre.Inicializar
    End If
    
End Sub

Private Sub Submenu_Rpt_Cursos_Indice_Asistencia_Click()
Dim Frm_Rpt_Cursos_Indice_Asistencia As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("CURSOS INDICE DE ASISTENCIA") Then
        Conectar_Ayudante.Enfocar ("CURSOS INDICE DE ASISTENCIA")
    Else
        Load Frm_Rpt_Cursos_Indice_Asistencia
        Frm_Rpt_Cursos_Indice_Asistencia.Caption = "CURSOS INDICE DE ASISTENCIA"
        Frm_Rpt_Cursos_Indice_Asistencia.Reporte = "Cursos_Indice_Asistencia"
        Frm_Rpt_Cursos_Indice_Asistencia.Btn_Exportar_PDF.Enabled = True
        Frm_Rpt_Cursos_Indice_Asistencia.Btn_Exportar_PDF.Visible = True
        Call Frm_Rpt_Cursos_Indice_Asistencia.Cargar_Frame(Frm_Rpt_Cursos_Indice_Asistencia.Fra_Rpt_Cursos_Indice_Asistencias, Frm_Rpt_Cursos_Indice_Asistencia)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Indice_Asistencia", Frm_Rpt_Cursos_Indice_Asistencia)
        Frm_Rpt_Cursos_Indice_Asistencia.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Cursos_Por_Empleado_Click()
Dim Frm_Rpt_Cursos_Por_Empleado As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("CURSOS TOMADOS POR EMPLEADO") Then
        Conectar_Ayudante.Enfocar ("CURSOS TOMADOS POR EMPLEADO")
    Else
        Load Frm_Rpt_Cursos_Por_Empleado
        Frm_Rpt_Cursos_Por_Empleado.Caption = "CURSOS TOMADOS POR EMPLEADO"
        Frm_Rpt_Cursos_Por_Empleado.Reporte = "Cursos_Por_Empleado"
        Frm_Rpt_Cursos_Por_Empleado.Btn_Exportar_PDF.Enabled = True
        Frm_Rpt_Cursos_Por_Empleado.Btn_Exportar_PDF.Visible = True
        Call Frm_Rpt_Cursos_Por_Empleado.Cargar_Frame(Frm_Rpt_Cursos_Por_Empleado.Fra_Cursos_Tomados_Por_Empleado, Frm_Rpt_Cursos_Por_Empleado)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Por_Empleado", Frm_Rpt_Cursos_Por_Empleado)
        Frm_Rpt_Cursos_Por_Empleado.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Cursos_Resumen_Mensual_Click()
Dim Frm_Rpt_Cursos_Resumen_Mensual As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("CURSOS RESUMEN MENSUAL") Then
        Conectar_Ayudante.Enfocar ("CURSOS RESUMEN MENSUAL")
    Else
        Load Frm_Rpt_Cursos_Resumen_Mensual
        Frm_Rpt_Cursos_Resumen_Mensual.Caption = "CURSOS RESUMEN MENSUAL"
        Frm_Rpt_Cursos_Resumen_Mensual.Reporte = "Cursos_Resumen_Mensual"
        Frm_Rpt_Cursos_Resumen_Mensual.Btn_Exportar_PDF.Enabled = True
        Frm_Rpt_Cursos_Resumen_Mensual.Btn_Exportar_PDF.Visible = True
        Call Frm_Rpt_Cursos_Resumen_Mensual.Cargar_Frame(Frm_Rpt_Cursos_Resumen_Mensual.Fra_Cursos_Resumen_Mensual, Frm_Rpt_Cursos_Resumen_Mensual)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Cursos_Resumen_Mensual", Frm_Rpt_Cursos_Resumen_Mensual)
        Frm_Rpt_Cursos_Resumen_Mensual.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Empleados_Huella_Click()
Dim Frm_Rpt_Empleados_Huella As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("EMPLEADOS CON HULLEA COMEDOR") Then
        Conectar_Ayudante.Enfocar ("EMPLEADOS CON HULLEA COMEDOR")
    Else
        Load Frm_Rpt_Empleados_Huella
        Frm_Rpt_Empleados_Huella.Caption = "EMPLEADOS CON HULLEA COMEDOR"
        Frm_Rpt_Empleados_Huella.Reporte = "Empleados_Huella_Comedor"
        Call Frm_Rpt_Empleados_Huella.Cargar_Frame(Frm_Rpt_Empleados_Huella.Fra_Rpt_Empleados_Alta, Frm_Rpt_Empleados_Huella)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Empleados_Huella", Frm_Rpt_Empleados_Huella)
        Frm_Rpt_Empleados_Huella.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Entradas_Comedor_Click()
Dim Frm_Rpt_Comedor As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("REPORTE DE COMEDOR") Then
        Conectar_Ayudante.Enfocar ("REPORTE DE COMEDOR")
    Else
        Load Frm_Rpt_Comedor
        Frm_Rpt_Comedor.Caption = "REPORTE DE COMEDOR"
        Frm_Rpt_Comedor.Reporte = "Reporte_Comedor"
        
        Call Frm_Rpt_Comedor.Cargar_Frame(Frm_Rpt_Comedor.Fra_Reporte_Cursos, Frm_Rpt_Comedor)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Entradas_Comedor", Frm_Rpt_Comedor)
        Frm_Rpt_Comedor.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Faltas_Validadas_Click()
Dim Frm_Rpt_Faltas_Empleados_Validadas As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("FALTAS VALIDADAS") Then
        Conectar_Ayudante.Enfocar ("FALTAS VALIDADAS")
    Else
        Load Frm_Rpt_Faltas_Empleados_Validadas
        Frm_Rpt_Faltas_Empleados_Validadas.Caption = "FALTAS VALIDADAS"
        Frm_Rpt_Faltas_Empleados_Validadas.Reporte = "Empleados_Faltas_Validadas"
        Call Frm_Rpt_Faltas_Empleados_Validadas.Cargar_Frame(Frm_Rpt_Faltas_Empleados_Validadas.Fra_Faltas_Empleados, Frm_Rpt_Faltas_Empleados_Validadas)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_Faltas_Validadas", Frm_Rpt_Faltas_Empleados_Validadas)
        Frm_Rpt_Faltas_Empleados_Validadas.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_General_Cursos_Click()
Dim Frm_Rpt_General_Cursos As New Frm_Rpt_Reportes_RH

    If Conectar_Ayudante.Formulario_Cargado("REPORTE GENERAL DE CURSOS") Then
        Conectar_Ayudante.Enfocar ("REPORTE GENERAL DE CURSOS")
    Else
        Load Frm_Rpt_General_Cursos
        Frm_Rpt_General_Cursos.Caption = "REPORTE GENERAL DE CURSOS"
        Frm_Rpt_General_Cursos.Reporte = "Reporte_General_Cursos"
        Frm_Rpt_General_Cursos.Btn_Exportar_PDF.Enabled = True
        Frm_Rpt_General_Cursos.Btn_Exportar_PDF.Visible = True
        Call Frm_Rpt_General_Cursos.Cargar_Frame(Frm_Rpt_General_Cursos.Fra_Rpt_General_Cursos, Frm_Rpt_General_Cursos)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Rpt_General_Cursos", Frm_Rpt_General_Cursos)
        Frm_Rpt_General_Cursos.Inicializar
    End If
End Sub

Private Sub Submenu_Rpt_Historico_Vacaciones_Click()
 Load Frm_Adm_Historico_Vacaciones
 Frm_Adm_Historico_Vacaciones.Inicializa
End Sub

Private Sub Submenu_Rpt_No_Checadas_Click()
DoEvents
Load Frm_Rpt_No_Checadas_2
Frm_Rpt_No_Checadas_2.Inicializa
End Sub

Private Sub Submenu_Salir_Click()
    End
End Sub

Private Sub Submenu_Calculadora_Click()
Dim szfilename As String
Dim nLength As Long
Const MAX_PATH = 255
    szfilename = Space(MAX_PATH)
    nLength = GetWindowsDirectory(szfilename, Len(szfilename))
    'Indica si no existe el Kernel 32
    If nLength = 0 Then
        MsgBox "Unable to Obtain the Windows Directory"
    End If
    szfilename = Left$(szfilename, nLength) & "\SYSTEM32\CALC.exe"
    RETVAL = Shell(szfilename, 1)
End Sub

Private Sub Submenu_Impresora_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    CommonDialog1.ShowPrinter
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

Private Sub Submenu_Sistema_Click()
    Frm_Apl_Acerca_de.Show
End Sub

Private Sub Submenu_Usuarios_Click()
    'Le asigna la palabra usuarios para manejar el picture del catálogo usuarios
    Catalogo = "USUARIOS"
    Load Frm_Cat_Generales
    Frm_Cat_Generales.Lbl_Titulo.Caption = "Usuarios"
    Frm_Cat_Generales.Caption = Frm_Cat_Generales.Lbl_Titulo
    'Carga el picture de Usuarios
    Conectar_Ayudante.Cargar_Picture Frm_Cat_Generales.Pic_Usuarios, Frm_Cat_Generales
    'Llama a la función de consulta usuarios para llenar el grid del catálogo
    Call Frm_Cat_Generales.Consulta_Usuarios
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Usuarios", Frm_Cat_Generales)
End Sub

Private Sub SubMenu_Btn_Adm_Rpt_Accesos_Almacenes_Click()
Dim Frm_Rpt_Accesos_Almacenes As New Frm_Rpt_Reportes_RH
    
    If Conectar_Ayudante.Formulario_Cargado("ACCESOS ALMACENES") Then
        Conectar_Ayudante.Enfocar ("ACCESOS AL ALMACEN")
    Else
        Load Frm_Rpt_Accesos_Almacenes
        Frm_Rpt_Accesos_Almacenes.Caption = "ACCESOS AL ALMACEN"
        Frm_Rpt_Accesos_Almacenes.Reporte = "Accesos_Almacenes"
        Call Frm_Rpt_Accesos_Almacenes.Cargar_Frame(Frm_Rpt_Accesos_Almacenes.Fra_Accesos_Almacen, Frm_Rpt_Accesos_Almacenes)
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Btn_Adm_Rpt_Accesos_Almacenes", Frm_Rpt_Accesos_Almacenes)
        Frm_Rpt_Accesos_Almacenes.Inicializar
    End If
End Sub
