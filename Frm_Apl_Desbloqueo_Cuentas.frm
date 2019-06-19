VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_Apl_Desbloqueo_Cuentas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Desbloqueo_Cuentas"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8550
   Begin VB.CommandButton Btn_Salir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   555
      Left            =   6885
      Picture         =   "Frm_Apl_Desbloqueo_Cuentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4500
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   555
      Left            =   225
      Picture         =   "Frm_Apl_Desbloqueo_Cuentas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "M"
      Top             =   4500
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame Fra_Desbloqueo_Cuentas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   135
      TabIndex        =   3
      Top             =   465
      Width           =   8250
      Begin MSFlexGridLib.MSFlexGrid Grid_Usuarios 
         Height          =   3645
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6429
         _Version        =   393216
         Rows            =   0
         Cols            =   3
         FixedRows       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
   End
   Begin VB.TextBox Txt_Usuario_ID 
      Height          =   330
      Left            =   6750
      TabIndex        =   5
      Top             =   3015
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Lbl_USUARIOS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DESBLOQUEO DE CUENTAS"
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
      Left            =   1665
      TabIndex        =   4
      Top             =   45
      Width           =   5250
   End
End
Attribute VB_Name = "Frm_Apl_Desbloqueo_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:Desbloqueo_Cuentas
    'DESCRIPCIÓN: Cambia los parametros de bloqueo a un estado que permita utilizar la cuenta
    'PARÁMETROS :
    'CREO       : Miguel Segura
    'FECHA_CREO        : 29-Octubre-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Desbloqueo_Cuentas()
Dim Rs_Modificacion_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro de la tabla Cat_Usuarios, modifica los valores del registro que tiene el usuario seleccionado

Set Conectar_Ayudante = New Ayudante
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
        'Consulta el Usuario actual seleccionado
        Mi_SQL = "SELECT * FROM Apl_Cat_Usuarios"
        Mi_SQL = Mi_SQL & " WHERE Usuario_ID ='" & Trim(Txt_Usuario_ID.Text) & "'"
        Set Rs_Modificacion_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
        'Modifica los datos de la tabla Cat_Usuarios
        With Rs_Modificacion_Apl_Cat_Usuarios
            .Edit
                .rdoColumns("Sesion_Abierta") = "NO"
            .Update
        End With
        Rs_Modificacion_Apl_Cat_Usuarios.Close
        MsgBox "La cuenta se ha desbloqueado con exito", vbInformation
        Unload Me
        Tipo_Validacion = "Login"
        Load Frm_Apl_Login
        Frm_Apl_Login.txt_Login.SetFocus
    Conexion_Base.CommitTrans
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
        'Call Captura_Automatica_Reporte_Bitacora("Modulo: Desbloqueo de Cuentas" & Chr(13) & "Forma:Frm_Apl_Desbloqueo_Cuentas" & Chr(13) & "Evento: Desbloqueo_Cuentas" & Chr(13) & "Error: " & Er.Description, "Desbloqueo de Cuentas")
    Next Er
End Sub


Private Sub Btn_Modificar_Click()
If Grid_Usuarios.Rows > 1 Then
    'Funcion para desbloquear las uentas
    Call Desbloqueo_Cuentas
Else
    MsgBox "No hay usuarios que desbloquear", vbInformation + vbOKOnly, Me.Caption
End If
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
    Tipo_Validacion = False
    Load Frm_Apl_Login
End Sub

Private Sub Form_Load()
    Me.Height = 5600
    Me.Width = 8670
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Usuarios
    'DESCRIPCIÓN: Consulta todos los Usuarios que hay en la tabla Cat_Usuarios
    '             llenando el Grid
    'PARÁMETROS : Nombre: Indica el nombre del rol que se pretende buscar
    'CREO       : Jorge Razo
    'FECHA_CREO :
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Consulta_Usuarios()
On Error GoTo HANDLER
Dim Rs_Consulta_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro, consulta los datos generales de los usuarios
Set Conectar_Ayudante = New Ayudante
    
    Grid_Usuarios.Rows = 0
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Usuario_ID, Nombre, Login"
    Mi_SQL = Mi_SQL & " FROM Cat_Usuarios"
    Mi_SQL = Mi_SQL & " WHERE Sesion_Abierta = 'SI'"
    Mi_SQL = Mi_SQL & " AND Nombre_Equipo = '" & GetCurrentMachineName & "'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
    Set Rs_Consulta_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'Llena el grid con los datos obtenidos de la consulta
    If Not Rs_Consulta_Apl_Cat_Usuarios.EOF Then
        'Coloca un encabezado en el grid
        Grid_Usuarios.AddItem "Usuario ID" & Chr(9) & "Nombre" & Chr(9) & "Login"
        While Not Rs_Consulta_Apl_Cat_Usuarios.EOF
            Grid_Usuarios.AddItem Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Usuario_ID") _
            & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Nombre") _
            & Chr(9) & Rs_Consulta_Apl_Cat_Usuarios.rdoColumns("Login")
            Grid_Usuarios.FixedRows = 1
            Rs_Consulta_Apl_Cat_Usuarios.MoveNext
        Wend
        With Grid_Usuarios
            .Col = 0
            .Row = 1
            .ColSel = .Cols - 1
            .RowSel = 1
            .TopRow = .Row
            .SetFocus
        End With
        Txt_Usuario_ID.Text = Grid_Usuarios.TextMatrix(1, 0)
        'Configura el tamaño de las columnas del grid_usuarios
        Grid_Usuarios.ColWidth(0) = 1000 'Usuario_ID
        Grid_Usuarios.ColWidth(1) = 5000 'Nombre
        Grid_Usuarios.ColWidth(2) = 1550 'Login
    End If
    'Cierra el manejador del registro
    Rs_Consulta_Apl_Cat_Usuarios.Close
    Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
        'Call Captura_Automatica_Reporte_Bitacora("Modulo: Desbloqueo de Cuentas" & Chr(13) & "Forma:Frm_Apl_Desbloqueo_Cuentas" & Chr(13) & "Evento: Consulta_Usuarios" & Chr(13) & "Error: " & Er.Description, "Desbloqueo de Cuentas")
    Next Er
End Sub

Private Sub Grid_Usuarios_Click()
    Txt_Usuario_ID.Text = ""
    If Grid_Usuarios.Rows > 1 Then
        Txt_Usuario_ID.Text = Grid_Usuarios.TextMatrix(Grid_Usuarios.RowSel, 0)
    End If
End Sub
