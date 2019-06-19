VERSION 5.00
Begin VB.Form Frm_Apl_Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   1545
   ClientLeft      =   6255
   ClientTop       =   5370
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_Login 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton Btn_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton Btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   975
      Width           =   1140
   End
   Begin VB.TextBox txt_Password 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Lbl_Login 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Login"
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
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Lbl_Password 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
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
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   825
   End
End
Attribute VB_Name = "Frm_Apl_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Intentos As Integer
    
'***********************************************************************************
'NOMBRE DE LA FUNCIÓN: Habilitar_Configuración
'DESCRIPCIÓN: Habilita los menus de acuerdo al usuario y sus permiso para
'             acceder a ellos
'PARÁMETROS:  Usuario: Usuario_ID y es usada para consultar a cuales menu
'             puede acceder el usuario
'CREO      :  Jorge Razo
'FECHA_CREO:  12-Marzo-2005
'MODIFICO          : Yazmin Abigail Delgado Gómez, Jorge Razo, Yazmin Delgado
'FECHA_MODIFICO    : 16-Junio-2005, 17-Noviembre-2005, 28-Mayo-2007
'CAUSA_MODIFICACIÓN: Porque no habilitaba los menus adecuadamente, marcaba error
'                    al momento de deshabilitar algun menu
'                  : Para que ocultara los menus no habiitados y no que los
'                    pusiera como deshabilitados
'                  : Porque se cambio la forma de habilitar o deshabilitar
'                    los menus y submenus del usuario
'**********************************************************************************
Private Function Habilita_Configuracion()
Dim Rs_Consulta_Apl_Cat_Accesos As rdoResultset 'Consulta los menus y submenus a los cuales puede entrar el usuario
Dim Ctl As Control                              'Toma la forma del objeto al que esta apuntando en ese momento
Dim Encabezado As Integer                       'Almacenara el valor 1 al encontrar un encabezado y valida los siguientes menus para que los ocule o no

On Error GoTo HANDLER:
    '1. Busca en la forma si el objeto se llama menu o submenu
    '2. Por medio del Usuario_ID habilita o deshabilita los menus
    For Each Ctl In MDIFrm_Apl_Principal.Controls
        On Error Resume Next
        If UCase(Mid(Ctl.Name, 1, 4)) = "MENU" Or UCase(Mid(Ctl.Name, 1, 7)) = "SUBMENU" Then
            'Consulta que el menu que se esta seleccionado de la pantalla se encuentre
            'habilitado
            Mi_SQL = "SELECT * FROM Seguridad_Sistema"
            Mi_SQL = Mi_SQL & " WHERE Nombre_Sistema = '" & UCase(Ctl.Name) & "'"
            Mi_SQL = Mi_SQL & " AND Rol_ID = '" & Rol_ID & "'"
            Set Rs_Consulta_Apl_Cat_Accesos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Consulta_Apl_Cat_Accesos.EOF Then
                With Rs_Consulta_Apl_Cat_Accesos
                    If .rdoColumns("Habilitar") = "S" Then
                        Ctl.Visible = True
                    Else
                        Ctl.Visible = False
                    End If
                End With
            Else
                Ctl.Visible = False
            End If
            Rs_Consulta_Apl_Cat_Accesos.Close
        End If
    Next Ctl
    If Usuario_ID = "00001" Then
        MDIFrm_Apl_Principal.Submenu_Apl_Respaldo_Sistema.Visible = True
        MDIFrm_Apl_Principal.Submenu_Apl_Respaldo_Sistema.Enabled = True
    End If
Exit Function
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Function

'*************************************************************************************
'NOMBRE DE LA FUNCIÓN: Oculta_Menus
'DESCRIPCIÓN: Dehabilita los menus más importantes del sistema para que cuando el
'             usuario no se logie o no tenga permisos de entrar al sistema no pueda
'             manipular la información que se existe en el sistema
'PARÁMETROS:
'CREO:        Jorge Razo
'FECHA_CREO:  12-Marzo-2005
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*************************************************************************************
Public Sub Oculta_Menus()
Dim Objeto As Control       '#  Busca todos los controles
    
    For Each Objeto In MDIFrm_Apl_Principal.Controls
        If UCase(Mid(Objeto.Name, 1, 4)) = "MENU" Then
            If (Objeto.Caption <> "&Archivo" And Objeto.Caption <> "&Ventanas") Then Objeto.Visible = False
        End If
    Next
    'MDIFrm_Apl_Principal.Submenu_Formato.Visible = False
    MDIFrm_Apl_Principal.Submenu_Apl_Respaldo_Sistema.Visible = False
    MDIFrm_Apl_Principal.Submenu_Apl_Respaldo_Sistema.Enabled = False
End Sub

Private Sub Btn_Cancel_Click()
    Unload Frm_Apl_Login
End Sub

Private Sub Btn_OK_Click()
Dim Rs_Aceptar_Cat_Usuarios As rdoResultset     'Obtiene el login del usuario
Dim Mi_SQL As String                            'Consulta para el login del usuario
Dim Security As Integer                         '
Dim SIGUIENTE As Integer
Dim Rs_Empleado As rdoResultset

On Error GoTo HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 11
    Mi_SQL = "SELECT Cat_Usuarios.*,Cat_Roles.Nombre AS Rol"
    Mi_SQL = Mi_SQL & " FROM Cat_Usuarios,Cat_Roles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Usuarios.Rol_ID=Cat_Roles.Rol_ID"
    Mi_SQL = Mi_SQL & " AND Login='" & UCase(Trim(Txt_Login)) & "'"
    Mi_SQL = Mi_SQL & " AND Estatus='ACTIVO'"
    Set Rs_Aceptar_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Si durante la consulta no encontro al usuario manda un mensaje
    If Rs_Aceptar_Cat_Usuarios.EOF Then
        MsgBox "Usuario Inválido, ¡Verifique su usuario!", , "Login"
        Txt_Login.SetFocus
        SendKeys "{Home}+{End}"
        MDIFrm_Apl_Principal.MousePointer = 0
        Exit Sub
    End If
    '1. Valida primero que las cajas de texto no esten vacías y compara lo que tienen
    '2. Valida el login con la base de datos
    '3. Compara el password y manda mensaje de password incorrecto si no es igual
    '4. Si es exitoso guarda las variables globales en Usuario
    '5. Habilita loe menus de acuerdo a su seguridad
    With Rs_Aceptar_Cat_Usuarios
        If Trim(Txt_Login.Text) <> "" Or Trim(Txt_Password.Text) <> "" Then
            If UCase(Txt_Login.Text) = UCase(.rdoColumns("Login")) Then
                If UCase(Txt_Password.Text) = UCase(.rdoColumns("Contraseña")) Then
                    'Valida la caducidad del password, superusuario se brinca esta validación
                    Consulta_Parametros_Generales
                    If (DateDiff("d", .rdoColumns("Fecha_Ultimo_Cambio_Password"), Now) < Dias_Caducidad_Contraseñas Or Dias_Caducidad_Contraseñas = 0 Or .rdoColumns("Usuario_ID") = "00001") Then
                        Usuario_ID = .rdoColumns("Usuario_ID")
                        Nombre_Usuario = .rdoColumns("Nombre")
                        Rol_ID = .rdoColumns("Rol_ID")
                        Rol = .rdoColumns("Rol")
                        If Not IsNull(.rdoColumns("Area_ID")) Then
                            Area_ID = .rdoColumns("Area_ID")
                        End If
                        MDIFrm_Apl_Principal.StatusBar.Panels(3).Text = Nombre_Usuario
                        Usuario = Txt_Login.Text
                        'Consulta el ID del empleado si no es ADMINISTRADOR del sistema para filtar reportes
                        If Not IsNull(.rdoColumns("No_Nomina")) And Rol_ID <> "00001" Then
                            Mi_SQL = "SELECT Empleado_ID,Supervisor_ID,Tipo FROM Cat_Empleados WHERE No_Tarjeta='" & .rdoColumns("No_Nomina") & "'"
                            Set Rs_Empleado = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                            If Not Rs_Empleado.EOF Then
                                If Rs_Empleado.rdoColumns("Tipo") = "S" Then
                                    Empleado_Supervisor_ID = Rs_Empleado.rdoColumns("Empleado_ID")
                                Else
                                    Empleado_Supervisor_ID = Rs_Empleado.rdoColumns("Supervisor_ID")
                                End If
                            Else
                                Empleado_Supervisor_ID = ""
                            End If
                            Rs_Empleado.Close
                        Else
                            Empleado_Supervisor_ID = ""
                        End If
                        Unload Frm_Apl_Login
                        Habilita_Configuracion      'Permisos de rol
                        Ruta_Temporal = Obtiene_Ruta_Temporal
                        Consulta_Parametros
                        If Rol_ID = "00001" Then    'Acualiza los turnos en automático
                            Actualiza_Turnos_Programacion
                        End If
                        'Crea el ODBC
                        Call Crear_ODBC_BD
                        MDIFrm_Apl_Principal.Caption = "SISTEMA DE RECURSOS HUMANOS [" & UCase(Rol) & ": " & Nombre_Usuario & "]"
                    Else
                        MsgBox "¡Su Password ha Caducado!", vbCritical
                        Unload Frm_Apl_Login
                        Load Frm_Apl_Cambio_Password
                    End If
                Else
                    MDIFrm_Apl_Principal.MousePointer = 0
                    Intentos = Intentos + 1
                    If Intentos_Sesion_Fallidos = Intentos Then
                        Deshabilita_Usuario 'Deshabilita la cuenta del usuario para que no pueda acceder al sistema
                    Else
                        If Intentos_Sesion_Fallidos = (Intentos + 1) Then
                            MsgBox "Invalido Password, Verifique su password!" & Chr(13) & Chr(13) & _
                                   "Le resta un intento para no inabilitar la cuenta", vbCritical
                        Else
                            MsgBox "Invalido Password, Verifique su password!", vbCritical, "Login"
                        End If
                    End If
                    Txt_Password.SetFocus
                    SendKeys "{Home}+{End}"
                    Exit Sub
                End If
            Else
                MsgBox "Invalido usuario, ¡Veririque su usuario!", , "Login"
                Txt_Login.SetFocus
                SendKeys "{Home}+{End}"
                MDIFrm_Apl_Principal.MousePointer = 0
                Exit Sub
            End If
        Else
            MsgBox "Invalido Usuario, ¡Verifique su usuario!", , "Login"
            Txt_Login.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End With
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
HANDLER:
    MDIFrm_Apl_Principal.MousePointer = 0
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub

Private Sub Deshabilita_Usuario()
Dim Rs_Modifica_Apl_Cat_Usuarios As rdoResultset 'Modifica el estatus del usuario de activo a inactivo

On Error GoTo HANDLER:
    'Consulta los datos generales del usuario al cual se le va a deshabilitar la cuenta
    Mi_SQL = "SELECT * FROM Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Login='" & UCase(Txt_Login.Text) & "'"
    Set Rs_Modifica_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Cambia el estatus de activo a inactivo
    If Not Rs_Modifica_Apl_Cat_Usuarios.EOF Then
        With Rs_Modifica_Apl_Cat_Usuarios
            .Edit
                .rdoColumns("Estatus") = "INACTIVO"
                .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                .rdoColumns("Fecha_Modifico") = Now
            .Update
        End With
        MsgBox "Usuario Deshabilitado", vbCritical
    End If
    Rs_Modifica_Apl_Cat_Usuarios.Close
    Exit Sub
HANDLER:
    Debug.Print Err, error
    For Each Er In rdoErrors
        MsgBox Err.Description
    Next
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Oculta_Menus
    MDIFrm_Apl_Principal.Caption = "SISTEMA DE ADMINISTRACION INTEGRAL"
End Sub

Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Password_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

