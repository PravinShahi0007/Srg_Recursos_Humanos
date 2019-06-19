VERSION 5.00
Begin VB.Form Frm_Ope_Login_Password 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOGIN Y PASSWORD"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Contraseña 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4380
      TabIndex        =   4
      Top             =   -15
      Width           =   4380
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
         Left            =   135
         TabIndex        =   2
         Top             =   2050
         Width           =   1140
      End
      Begin VB.Frame Fra_Password 
         BackColor       =   &H00FFFFFF&
         Height          =   1950
         Left            =   135
         TabIndex        =   5
         Top             =   45
         Width           =   4110
         Begin VB.TextBox Txt_Login 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1380
            TabIndex        =   0
            Top             =   900
            Width           =   2500
         End
         Begin VB.TextBox Txt_Password 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            IMEMode         =   3  'DISABLE
            Left            =   1380
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   1395
            Width           =   2500
         End
         Begin VB.Label Lbl_Titulo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mensaje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   90
            TabIndex        =   8
            Top             =   180
            Width           =   3930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   7
            Top             =   945
            Width           =   585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   6
            Top             =   1455
            Width           =   1035
         End
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
         Left            =   3105
         TabIndex        =   3
         Top             =   2050
         Width           =   1140
      End
   End
   Begin VB.Frame Fra_Ubicacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selecciona tu Ubicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   4170
      Begin VB.CommandButton Btn_Aceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   1215
         TabIndex        =   11
         Top             =   1365
         Width           =   1275
      End
      Begin VB.ComboBox Cmb_Ubicacion 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   675
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Frm_Ope_Login_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Salir As Boolean 'Indica si se puede salir de la forma o no
Public Operacion As String

Private Sub Btn_Aceptar_Click()
    If Cmb_Ubicacion.ListIndex > -1 Then
        Ubicacion_Usuario = Format(Cmb_Ubicacion.ItemData(Cmb_Ubicacion.ListIndex), "00000")
        Salir = True
        Unload Me
    End If
End Sub

Private Sub Btn_Cancel_Click()
'    If Operacion = "RECOLECCION_DINERO" Then
'        If Consulta_Limite_Efectivo = True Or Bloquear_Exceder_Limite_Efectivo = True Then
'            MsgBox "Necesita hacer la recolección" & Chr(13) & Chr(13) & _
'                   "ya que el límite de efectivo fue excedido"
'            Txt_Login.SetFocus
'            Exit Sub
'        End If
'    End If
    Salir = True
    Me.Hide
    MDIFrm_Apl_Principal.Enabled = True
End Sub

Private Sub Btn_OK_Click()
    If Trim(Txt_Login.Text) <> "" And Trim(Txt_Password.Text) <> "" Then
        Consulta_Usuario_Valido 'Consulta que sea un usuario valido el que este entrado a la pantalla
    Else
        If Trim(Txt_Login.Text) = "" Then
            MsgBox "Proporcione el login", vbInformation
            Txt_Login.SetFocus
        Else
            MsgBox "Proporcione el password", vbInformation
            Txt_Password.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Salir = False
    Me.Height = 3000
    Me.Width = 4500
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Txt_Login.Text = ""
    Txt_Password.Text = ""
    Call Conectar_Ayudante.Llena_Combo_Item("Almacen_ID,Nombre", "Cat_Almacenes", Cmb_Ubicacion, 0, "Almacen_ID")
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Usuario_Valido
    'DESCRIPCIÓN: Valida que el login y password introducido por el usuario sea
    '             valido en el sistema si no es valido manda un mensaje al
    '             usuario, pero si es valildo entonces hace visible el
    '             pic_recoleccion que es en donde se va a llebar la captura
    '             de la recolección del dinero
    'PARÁMETROS :
    'CREO       : Yazmin A Delgado Gómez
    'FECHA_CREO : 24-Abril-2007
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Usuario_Valido()
Dim Rs_Consulta_Cat_Usuarios As rdoResultset 'Consulta si es un usuario valido el que se esta logeando

    'Consulta que el login y password que fueron introducidos por el usuario sea valido en el sistema
    Mi_SQL = "SELECT Usuario_ID,Cat_Usuarios.Nombre,Login,Contraseña,Cat_Roles.Nombre AS Rol_Lista"
    Mi_SQL = Mi_SQL & " FROM Cat_Usuarios,Cat_Roles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Usuarios.Rol_ID=Cat_Roles.Rol_ID"
    Mi_SQL = Mi_SQL & " AND Login='" & Trim(Txt_Login.Text) & "'"
    Mi_SQL = Mi_SQL & " AND Contraseña='" & Trim(Txt_Password.Text) & "'"
    Set Rs_Consulta_Cat_Usuarios = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    'si se encontraron validos los datos dados por el usuario entonces de acuerdo a la
    'operación se ejecuta las acciones a realizar
    If Not Rs_Consulta_Cat_Usuarios.EOF Then
        'Si se trata de una operación de recolección entonces manda llamar la forma correspondiente a esta operación
        Select Case Operacion
            Case "CAMBIO_TIPO_PRECIO"
                If Rs_Consulta_Cat_Usuarios.rdoColumns("Rol_Lista") = "ADMINISTRADOR" Or Rs_Consulta_Cat_Usuarios.rdoColumns("Rol_Lista") = "VENTAS" Then
                    Frm_Ope_Pedidos.Cmb_Tipo_Precio.Enabled = True
                    Frm_Ope_Pedidos.Btn_Cambiar_Tipo_Precio.Enabled = False
                    Frm_Ope_Pedidos.Cmb_Tipo_Precio.SetFocus
                    Rs_Consulta_Cat_Usuarios.Close
                    Unload Me
                Else
                    MsgBox "Usted no tiene sufucientes privilegios para hacer este cambio", vbExclamation
                End If
            Case "CAMBIO_PRECIO_CLIENTE"
'                If Rs_Consulta_Cat_Usuarios.rdoColumns("Rol_Lista") = "ADMINISTRADOR" Or Rs_Consulta_Cat_Usuarios.rdoColumns("Rol_Lista") = "VENTAS" Then
                    Frm_Cat_Clientes.Cmb_Tipo_Precio.Enabled = True
                    Frm_Cat_Clientes.Btn_Cambiar_Tipo_Precio.Enabled = False
                    Frm_Cat_Clientes.Cmb_Tipo_Precio.SetFocus
                    Rs_Consulta_Cat_Usuarios.Close
                    Unload Me
'                Else
'                    MsgBox "Usted no tiene sufucientes privilegios para hacer este cambio", vbExclamation
'                End If
            
'            Case "DESCUENTO_VENTA"
'                Frm_Ope_Punto_Venta.Fra_Total_Venta.Enabled = True
'                Frm_Ope_Punto_Venta.Txt_Descuento.Enabled = True
'                Frm_Ope_Punto_Venta.Txt_Descuento.Locked = False
'                Frm_Ope_Punto_Venta.Txt_Descuento.Appearance = 1
'                Frm_Ope_Punto_Venta.Txt_Descuento.SetFocus
'                Rs_Consulta_Cat_Usuarios.Close
'                Unload Me
'            Case "CANCELACION_PEDIDO"
'                Rs_Consulta_Cat_Usuarios.Close
'                Frm_Ope_Punto_Venta.Cancelar_Abono_Pedido
'                Unload Me
'
'            Case "RECOLECCION_DINERO"
'                'Si se trata de recoleccion de dinero el usuario tiene que tener
'                Unload Frm_Ope_Recoleccion_Dinero
'                Frm_Ope_Recoleccion_Dinero.Show
'                Frm_Ope_Recoleccion_Dinero.Txt_Cantidad_Recolectada.SetFocus
'            'Indica si la operación a realizar es correspondiente a la apertura de un nuevo
'            'turno dentro del sistema
'            Case "ABRIR_TURNO"
'                Load Frm_Ope_Abrir_Turno
'                'Si operación esta vacia entonces no se puede abrir otro turno
'                If Operacion = "" Then
'                    Unload Frm_Ope_Abrir_Turno
'                'Si no entonces agrega el No del turno a dar de alta
'                Else
'                    Frm_Ope_Abrir_Turno.Txt_No_Turno.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Ope_Turnos", "No_Turno"), "0000000000")
'                    Frm_Ope_Abrir_Turno.Cmb_Turnos.SetFocus
'                End If
'            'Indica que la operación a realizar es un arqueo de caja
'            Case "ARQUEO_CAJA"
'                'Load Frm_Ope_Corte_Caja
'            'Indica si la operación a realizar es con respecto al corte de la caja
'            Case "CORTE_CAJA"
'                'Load Frm_Ope_Corte_Caja
        End Select
        Txt_Login.Text = ""
        Txt_Password.Text = ""
        Operacion = ""
        Me.Hide
        MDIFrm_Apl_Principal.Enabled = True
    'Si el Login o password no son validos en el sistema entonces manda un mensaje
    'al usuario
    Else
        MsgBox "El login y passwod no son validos para el sistema", vbInformation
        Txt_Login.SetFocus
        SendKeys "{Home}+{End}"
    End If
    'Rs_Consulta_Cat_Usuarios.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Salir = False Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii)
End Sub

Private Sub Txt_Password_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii)
End Sub
