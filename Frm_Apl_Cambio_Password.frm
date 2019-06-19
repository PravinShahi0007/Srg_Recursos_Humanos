VERSION 5.00
Begin VB.Form Frm_Apl_Cambio_Password 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CAMBIO DE CONTRASEÑA"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   4500
   Begin VB.CommandButton Btn_Actualizar 
      Caption         =   "Actualizar"
      Height          =   420
      Left            =   1365
      TabIndex        =   5
      Top             =   1965
      Width           =   1275
   End
   Begin VB.TextBox Txt_Contraseña_Anterior 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   795
      Width           =   1700
   End
   Begin VB.TextBox Txt_Login 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   20
      TabIndex        =   0
      Top             =   420
      Width           =   1700
   End
   Begin VB.TextBox Txt_Confirmar_Contraseña_Nueva 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1530
      Width           =   1700
   End
   Begin VB.TextBox Txt_Contraseña_Nueva 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1185
      Width           =   1700
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  CAMBIO DE PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   9
      Top             =   45
      Width           =   3465
   End
   Begin VB.Label Lbl_Contraseña_Anterior 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Anterior"
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
      Left            =   90
      TabIndex        =   8
      Top             =   885
      Width           =   1695
   End
   Begin VB.Label Lbl_Contraseña_Nueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Nueva"
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
      Left            =   90
      TabIndex        =   7
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label Lbl_Login 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   480
   End
   Begin VB.Label Lbl_Confirmar_Contraseña_Nueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Contraseña Nueva"
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
      Left            =   90
      TabIndex        =   3
      Top             =   1605
      Width           =   2445
   End
End
Attribute VB_Name = "Frm_Apl_Cambio_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Actualizar_Click()
Dim Rs_Modificacion_Apl_Cat_Usuarios As rdoResultset 'Manejo de registro de la tabla Cat_Usuarios, modifica el password del usuario logeado
Dim Rs_Consulta_Passwords_Anteriores As rdoResultset

Set Conectar_Ayudante = New Ayudante
On Error GoTo HANDLER
    'Valida que los tex no esten vacios
    If Trim(Txt_Login.Text) <> "" And Trim(Txt_Contraseña_Anterior.Text) <> "" And Trim(Txt_Contraseña_Nueva.Text) <> "" And Trim(Txt_Confirmar_Contraseña_Nueva.Text) <> "" Then
        'If Conectar_Ayudante.Es_Alfanumerico(Txt_Contraseña_Nueva.Text) = True Then
            'Valida qu la contraseña sea de por lo menos el parámetro de caracteres
            If Len(Txt_Contraseña_Nueva.Text) >= Longitud_Minima_Password Then
                'Valida que sea diferente la contraseña nueva que va a capturar
                If Txt_Contraseña_Anterior.Text <> Txt_Contraseña_Nueva.Text Or Historico_Password = 0 Then
                    'Valida que la confirmación de la contraseña sea igual a la nueva contraseña
                    If Txt_Contraseña_Nueva.Text = Txt_Confirmar_Contraseña_Nueva.Text Then
                        'Consulta el Usuario actual seleccionado
                        Mi_SQL = "SELECT * FROM Cat_Usuarios"
                        Mi_SQL = Mi_SQL & " WHERE Login='" & Trim(Txt_Login.Text) & "' AND Contraseña='" & Trim(Txt_Contraseña_Anterior.Text) & "' "
                        Set Rs_Modificacion_Apl_Cat_Usuarios = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        'Modifica los datos de la tabla Cat_Usuarios
                        With Rs_Modificacion_Apl_Cat_Usuarios
                            If Not .EOF Then
                                'Verifica que el password no sea el mismo que ya se tenia dado de alta
                                Mi_SQL = "SELECT TOP " & Historico_Password & " *"
                                Mi_SQL = Mi_SQL & " FROM Cat_Usuarios_Password"
                                Mi_SQL = Mi_SQL & " WHERE Usuario_ID='" & .rdoColumns("Usuario_ID") & "'"
                                Mi_SQL = Mi_SQL & " ORDER BY No_Partida DESC"
                                Set Rs_Consulta_Passwords_Anteriores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                                While Not Rs_Consulta_Passwords_Anteriores.EOF
                                    If Rs_Consulta_Passwords_Anteriores.rdoColumns("Password") = Trim(Txt_Contraseña_Nueva.Text) Then
                                        MsgBox "El nuevo password no puede ser el mismo que ha usado en las ultimas " & Historico_Password & " ocasiones", vbCritical
                                        Rs_Consulta_Passwords_Anteriores.Close
                                        Exit Sub
                                    End If
                                    Rs_Consulta_Passwords_Anteriores.MoveNext
                                Wend
                                Rs_Consulta_Passwords_Anteriores.Close
                                .Edit
                                    .rdoColumns("Contraseña") = Trim(Txt_Contraseña_Nueva.Text)
                                    .rdoColumns("Estatus") = "ACTIVO"
                                    .rdoColumns("Fecha_Ultimo_Cambio_Password") = Format(Now, "MM/dd/yyyy")
                                    .rdoColumns("Fecha_Caduca") = DateAdd("d", Dias_Caducidad_Contraseñas, .rdoColumns("Fecha_Ultimo_Cambio_Password"))
                                    .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                                    .rdoColumns("Fecha_Modifico") = Now
                                .Update
                                'Guarda el password en la tabla
                                Mi_SQL = "INSERT INTO Cat_Usuarios_Password (Usuario_ID, Password, Fecha_Password)"
                                Mi_SQL = Mi_SQL & " VALUES('" & .rdoColumns("Usuario_ID") & "'"
                                Mi_SQL = Mi_SQL & " , '" & .rdoColumns("Contraseña") & "'"
                                Mi_SQL = Mi_SQL & " , '" & .rdoColumns("Fecha_Ultimo_Cambio_Password") & "')"
                                Conexion_Base.Execute Mi_SQL
                                MsgBox "Password Modificado Exitosamente", vbInformation
                                Unload Frm_Apl_Cambio_Password
                                Load Frm_Apl_Login
                            Else
                                MsgBox "No coincide el login o el password, favor de revisarlo", vbExclamation
                            End If
                        End With
                        Rs_Modificacion_Apl_Cat_Usuarios.Close
                    Else
                        MsgBox "La confirmación del nuevo password no coincide con el nuevo password!" & Chr(13) & Chr(13) & _
                                                   "Confirme nuevamente su password", vbCritical
                    End If
                Else
                    MsgBox "La contraseña anterior y la nueva deben ser diferentes", vbCritical
                End If
            Else
                MsgBox "La longitud del password debe ser por lo menos de " & Longitud_Minima_Password & " caracteres", vbCritical
            End If
        'Else
        '    MsgBox "El password no esta compuesto por letras y numeros!" & Chr(13) & Chr(13) & _
        '                                       "Para obtener mayor seguridad, el password deve estar conformado por letras y numeros", vbCritical
        'End If
    Else
        MsgBox "Faltan Datos para poder modificar su password!", vbCritical
    End If
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Form_Load()
'    Txt_Login.SetFocus
    Me.Width = 4590
    Me.Height = 3000
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub


Private Sub Txt_Confirmar_Contraseña_Nueva_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Contraseña_Anterior_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Contraseña_Nueva_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Login_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

