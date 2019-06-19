VERSION 5.00
Begin VB.Form Frm_Apl_Registro_Sistema 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro del Software"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6540
   Begin VB.TextBox Txt_Nombre_Ext 
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   285
      Width           =   855
   End
   Begin VB.CommandButton Btn_Registrar 
      Caption         =   "Registrar"
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Fra_Datos_Configuracion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de Configuracion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   75
      TabIndex        =   23
      Top             =   2355
      Width           =   6375
      Begin VB.CommandButton Btn_Ajustes 
         Height          =   195
         Left            =   6165
         TabIndex        =   31
         Top             =   1215
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Txt_Usuario 
         Height          =   285
         Left            =   1095
         TabIndex        =   11
         Top             =   1080
         Width           =   2385
      End
      Begin VB.TextBox Txt_Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4530
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   720
         Width           =   1710
      End
      Begin VB.ComboBox Cmb_Tipo_Base 
         Height          =   315
         ItemData        =   "Frm_Apl_Registro_Sistema.frx":0000
         Left            =   1095
         List            =   "Frm_Apl_Registro_Sistema.frx":000A
         TabIndex        =   9
         Top             =   345
         Width           =   2385
      End
      Begin VB.TextBox Txt_Servidor 
         Height          =   285
         Left            =   1095
         TabIndex        =   10
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox Txt_Base_Datos 
         Height          =   285
         Left            =   4530
         TabIndex        =   12
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Lbl_Usuario 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   29
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Lbl_Password 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   28
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Lbl_Tipo_Base 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Base"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   26
         Top             =   405
         Width           =   720
      End
      Begin VB.Label Lbl_Servidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Servidor"
         Height          =   195
         Index           =   13
         Left            =   195
         TabIndex        =   25
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Lbl_Base_Datos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Datos"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   24
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame Fra_Datos_Empresa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   90
      TabIndex        =   15
      Top             =   0
      Width           =   6330
      Begin VB.TextBox Txt_Telefono 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Txt_Estado 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   1875
         Width           =   1935
      End
      Begin VB.TextBox Txt_Ciudad 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   1875
         Width           =   2415
      End
      Begin VB.TextBox Txt_CP 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   1230
         Width           =   1935
      End
      Begin VB.TextBox Txt_Colonia 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1230
         Width           =   2415
      End
      Begin VB.TextBox Txt_Direccion 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   915
         Width           =   5055
      End
      Begin VB.TextBox Txt_RFC 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox Txt_Nombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   285
         Width           =   4095
      End
      Begin VB.Label Lbl_Telefono 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telefono"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   1605
         Width           =   630
      End
      Begin VB.Label Lbl_RFC 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R.F.C."
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   22
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Lbl_CP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
         Height          =   195
         Index           =   5
         Left            =   3600
         TabIndex        =   21
         Top             =   1275
         Width           =   300
      End
      Begin VB.Label Lbl_Estado 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Lbl_Ciudad 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Lbl_Colonia 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   18
         Top             =   1275
         Width           =   525
      End
      Begin VB.Label Lbl_Direccion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Lbl_Nombre 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   16
         Top             =   330
         Width           =   555
      End
   End
End
Attribute VB_Name = "Frm_Apl_Registro_Sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Ajustes_Click()
    'Ajustar_Detalles_Anticipos_Proveedores
End Sub

Private Sub Btn_Registrar_Click()
'Contador para especificar la linea en la cual se va a ir guardando los datos
Dim I As Integer
    Set Conectar_Ayudante = New Ayudante
    'Abre el documento llamado Config.ini y empiesa a escribir lo contenido en las cajas de texto
    Open App.Path & "\Config.ini" For Output As #1
        Print #1, Txt_Nombre.Text
        Print #1, Txt_RFC.Text
        Print #1, Txt_Direccion.Text
        Print #1, Txt_Colonia.Text
        Print #1, Txt_CP.Text
        Print #1, Txt_Telefono.Text
        Print #1, Txt_Ciudad.Text
        Print #1, Txt_Estado.Text
        Print #1, Cmb_Tipo_Base.Text
        Print #1, Txt_Servidor.Text
        Print #1, Txt_Usuario.Text
        Print #1, Txt_Base_Datos.Text
        Print #1, Txt_Password.Text
    Close #1
    MsgBox "Datos Registrados", vbInformation
    Unload Frm_Apl_Registro_Sistema
    Conexion_Base.Close
    Conectar_Ayudante.Conexion 'Manda llamar la función Conexion contenida en el Module1
End Sub

Private Sub Form_Load()
Dim I As Integer        'Contador que indica que linea se esta leyendo del documento
Dim Linea As String     'Guarda el valor de la linea
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    On Error GoTo Etiqueta
        Open App.Path & "\Config.ini" For Input As #1
        I = 0
        '1. Llena las cajas de texto de acuerdo a lo contenido en el documento Config.ini
        Do While Not EOF(1)
            Line Input #1, Linea
            Select Case I
                Case 0
                    Txt_Nombre.Text = Linea
                Case 1
                    Txt_RFC.Text = Linea
                Case 2
                    Txt_Direccion = Linea
                Case 3
                    Txt_Colonia.Text = Linea
                Case 4
                    Txt_CP.Text = Linea
                Case 5
                    Txt_Telefono.Text = Linea
                Case 6
                    Txt_Ciudad.Text = Linea
                Case 7
                    Txt_Estado.Text = Linea
                Case 8
                    Cmb_Tipo_Base.Text = Linea
                Case 9
                    Txt_Servidor.Text = Linea
                Case 10
                    Txt_Usuario.Text = Linea
                Case 11
                    Txt_Base_Datos.Text = Linea
                Case 12
                    Txt_Password.Text = Linea
            End Select
            I = I + 1
        Loop
    Close #1
    Exit Sub
Etiqueta:
    MsgBox "El sistema no se encuentra registrado" & Chr(13) & "Favor de llenar sus datos", vbExclamation
End Sub


Private Sub Ajustar_Detalles_Anticipos_Proveedores()
Dim Mi_SQL As String
Dim Rs_Consulta_Anticipos_Proveedores As rdoResultset
Dim Rs_Modifica_Anticipos_Proveedores As rdoResultset
Dim Rs_Consulta_Anticipos_Proveedores_Detalles As rdoResultset
Dim Rs_Consulta_Movimientos As rdoResultset
Dim Rs_Modifica_Movimientos As rdoResultset
Dim Rs_Agrega_Adm_Movimiento As rdoResultset
Dim No_Movimiento As String
Dim Saldo_Anticipo As Double

On Error GoTo Handler
    Conexion_Base.BeginTrans
    'Consulta los anticpos aplicados
    Mi_SQL = "SELECT * FROM Adm_Anticipos_Proveedores"
    Mi_SQL = Mi_SQL & " WHERE Aplicado='S'"
    Set Rs_Consulta_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Anticipos_Proveedores.EOF
        Saldo_Anticipo = Rs_Consulta_Anticipos_Proveedores.rdoColumns("Total")
        'Busca si existe en la tabla de detalles
        Mi_SQL = "SELECT * FROM Adm_Detalles_Anticipos_Proveedores"
        Mi_SQL = Mi_SQL & " WHERE No_Anticipo=" & Rs_Consulta_Anticipos_Proveedores.rdoColumns("No_Anticipo")
        Set Rs_Consulta_Anticipos_Proveedores_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Anticipos_Proveedores_Detalles.EOF
            'Consulta el anticipo para modificarle el número de movimiento
            Mi_SQL = "SELECT * FROM Adm_Anticipos_Proveedores"
            Mi_SQL = Mi_SQL & " WHERE No_Anticipo=" & Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("No_Anticipo")
            Set Rs_Modifica_Anticipos_Proveedores = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            If Not Rs_Modifica_Anticipos_Proveedores.EOF Then
                'Valida si existe el movimeinto correcto
                Mi_SQL = "SELECT * FROM Adm_Movimientos"
                Mi_SQL = Mi_SQL & " WHERE Estatus='A'"
                Mi_SQL = Mi_SQL & " AND Tipo='E'"
                If Rs_Consulta_Anticipos_Proveedores.rdoColumns("Forma_Pago") = "CHEQUE" Then
                    Mi_SQL = Mi_SQL & " AND No_Cheque='" & Format(Rs_Consulta_Anticipos_Proveedores.rdoColumns("Referencia"), "0000000000") & "'"
                Else
                    Mi_SQL = Mi_SQL & " AND Referencia='" & Format(Rs_Consulta_Anticipos_Proveedores.rdoColumns("Referencia"), "0000000000") & "'"
                End If
                Mi_SQL = Mi_SQL & " AND Proveedor_Cliente='" & Rs_Consulta_Anticipos_Proveedores.rdoColumns("Proveedor_ID") & "'"
                Mi_SQL = Mi_SQL & " AND No_Factura='" & Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("No_Factura") & "'"
                Mi_SQL = Mi_SQL & " AND Empresa=" & Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Empresa")
                Mi_SQL = Mi_SQL & " AND Cantidad=" & Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Cantidad")
                Set Rs_Consulta_Movimientos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Rs_Consulta_Movimientos.EOF Then
                    If Saldo_Anticipo > Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Cantidad") Then
                        Saldo_Anticipo = Saldo_Anticipo - Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Cantidad")
                        'Actualiza el movimiento del anticipo
                        Mi_SQL = "SELECT * FROM Adm_Movimientos"
                        Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Rs_Modifica_Anticipos_Proveedores.rdoColumns("No_Movimiento") & "'"
                        Set Rs_Modifica_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        If Not Rs_Modifica_Movimientos.EOF Then
                            'Da de alta el nuevo movimiento
                            Set Rs_Agrega_Adm_Movimiento = Conectar_Ayudante.Recordset_Agregar("Adm_Movimientos")
                            No_Movimiento = Format(Conectar_Ayudante.Maximo_Catalogo("Adm_Movimientos", "No_Movimiento"), "0000000000")
                            With Rs_Agrega_Adm_Movimiento
                                .AddNew
                                    .rdoColumns("No_Movimiento") = No_Movimiento
                                    .rdoColumns("Fecha") = Format(Rs_Modifica_Movimientos.rdoColumns("Fecha"), "MM/dd/yyyy")
                                    .rdoColumns("Banco_ID") = Rs_Modifica_Movimientos.rdoColumns("Banco_ID")
                                    .rdoColumns("Tipo") = Rs_Modifica_Movimientos.rdoColumns("Tipo")
                                    .rdoColumns("Proveedor_Cliente") = Rs_Modifica_Movimientos.rdoColumns("Proveedor_Cliente")
                                    .rdoColumns("Estatus") = Rs_Modifica_Movimientos.rdoColumns("Estatus")
                                    .rdoColumns("Forma_Pago") = Rs_Modifica_Movimientos.rdoColumns("Forma_Pago")
                                    .rdoColumns("No_Cheque") = Rs_Modifica_Movimientos.rdoColumns("No_Cheque")
                                    .rdoColumns("Referencia") = Rs_Modifica_Movimientos.rdoColumns("Referencia")
                                    .rdoColumns("Concepto") = Rs_Modifica_Movimientos.rdoColumns("Concepto")
                                    .rdoColumns("Cantidad") = Saldo_Anticipo
                                    .rdoColumns("Banco") = Rs_Modifica_Movimientos.rdoColumns("Banco")
                                    .rdoColumns("Cuenta") = Rs_Modifica_Movimientos.rdoColumns("Cuenta")
                                    .rdoColumns("Beneficiario") = Rs_Modifica_Movimientos.rdoColumns("Beneficiario")
                                    .rdoColumns("Saldo") = 0
                                .Update
                            End With
                            Rs_Agrega_Adm_Movimiento.Close
                            With Rs_Modifica_Movimientos
                                .Edit
                                    .rdoColumns("Cantidad") = Val(Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Cantidad"))
                                    .rdoColumns("No_Factura") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("No_Factura")
                                    .rdoColumns("Empresa") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Empresa")
                                .Update
                            End With
                        End If
                        Rs_Modifica_Movimientos.Close
                    Else
                        'Actualiza el movimiento del anticipo
                        Mi_SQL = "SELECT * FROM Adm_Movimientos"
                        Mi_SQL = Mi_SQL & " WHERE No_Movimiento='" & Rs_Modifica_Anticipos_Proveedores.rdoColumns("No_Movimiento") & "'"
                        Set Rs_Modifica_Movimientos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                        If Not Rs_Modifica_Movimientos.EOF Then
                            No_Movimiento = Rs_Modifica_Anticipos_Proveedores.rdoColumns("No_Movimiento")
                            With Rs_Modifica_Movimientos
                                .Edit
                                    .rdoColumns("No_Factura") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("No_Factura")
                                    .rdoColumns("Empresa") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Empresa")
                                .Update
                            End With
                        End If
                        Rs_Modifica_Movimientos.Close
                    End If
                    'Actualiza el anticipo con el nuevo movimiento y factura
                    With Rs_Modifica_Anticipos_Proveedores
                        .Edit
                            .rdoColumns("No_Movimiento") = No_Movimiento
                            .rdoColumns("No_Factura") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("No_Factura")
                            .rdoColumns("Empresa") = Rs_Consulta_Anticipos_Proveedores_Detalles.rdoColumns("Empresa")
                        .Update
                    End With
                End If
                Rs_Consulta_Movimientos.Close
            End If
            Rs_Consulta_Anticipos_Proveedores_Detalles.MoveNext
        Wend
        Rs_Consulta_Anticipos_Proveedores_Detalles.Close
        Rs_Consulta_Anticipos_Proveedores.MoveNext
    Wend
    Rs_Consulta_Anticipos_Proveedores.Close
    Conexion_Base.CommitTrans
    MsgBox "Operación realizada con éxito", vbInformation
    Exit Sub
Handler:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

