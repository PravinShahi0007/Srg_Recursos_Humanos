VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Adm_Envio_Correo_Validacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   ClipControls    =   0   'False
   Icon            =   "Frm_Adm_Envio_Correo_Validacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic_Cat_Tipos_Faltas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4350
      Left            =   0
      ScaleHeight     =   4350
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Enviar"
         Height          =   555
         Left            =   45
         Picture         =   "Frm_Adm_Envio_Correo_Validacion.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "A"
         Top             =   3735
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   555
         Left            =   4185
         Picture         =   "Frm_Adm_Envio_Correo_Validacion.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3735
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Envio_Correos 
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   0
         TabIndex        =   2
         Top             =   405
         Width           =   6000
         Begin VB.CommandButton Btn_Buscar 
            Caption         =   "Buscar"
            Height          =   735
            Left            =   4815
            Picture         =   "Frm_Adm_Envio_Correo_Validacion.frx":0B20
            Style           =   1  'Graphical
            TabIndex        =   12
            Tag             =   "C"
            Top             =   180
            Width           =   1080
         End
         Begin VB.ComboBox Cmb_Correos_Enviados 
            Height          =   315
            ItemData        =   "Frm_Adm_Envio_Correo_Validacion.frx":10AA
            Left            =   1395
            List            =   "Frm_Adm_Envio_Correo_Validacion.frx":10B7
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   225
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Bitacora 
            Height          =   2115
            Left            =   90
            TabIndex        =   4
            Top             =   1080
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   3731
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Inicio 
            Height          =   315
            Left            =   1395
            TabIndex        =   8
            Top             =   630
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   39931
         End
         Begin MSComCtl2.DTPicker Dtp_Importacion_Fecha_Termino 
            Height          =   315
            Left            =   3285
            TabIndex        =   9
            Top             =   630
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   39931
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Al"
            Height          =   195
            Left            =   2955
            TabIndex        =   11
            Top             =   690
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha"
            Height          =   195
            Left            =   135
            TabIndex        =   10
            Top             =   690
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mostrar Correos"
            Height          =   195
            Left            =   135
            TabIndex        =   3
            Top             =   285
            Width           =   1110
         End
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENVIO MANUAL DE CORREO"
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
         Left            =   225
         TabIndex        =   1
         Top             =   0
         Width           =   5280
      End
   End
End
Attribute VB_Name = "Frm_Adm_Envio_Correo_Validacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean


Private Sub Btn_Buscar_Click()
    Grid_Bitacora.Rows = 0
    Consulta_Bitacora_Importacion (Cmb_Correos_Enviados.Text)
End Sub

Private Sub Btn_Nuevo_Click()
Dim cadena_mensaje As String
On Error GoTo HANDLER
Conexion_Base.BeginTrans
    'Envia el correo de confirmacion para supervisores
    cadena_mensaje = "La información de asistencias se ha generado"
    cadena_mensaje = cadena_mensaje & vbCrLf & "Favor de generar la validación de horas trabajadas"
    cadena_mensaje = cadena_mensaje & vbCrLf & "de la fecha: " & Format(Grid_Bitacora.TextMatrix(Grid_Bitacora.RowSel, 1), "ddd dd/MM/yyyy")
    Call Enviar_Correo(Email_Sistema, "Interfase Man Hours Management System", Email_validacion, "Supervision", "Importacion automática", cadena_mensaje)
    Call Inserta_Registro_Bitacora_Importacion(Grid_Bitacora.TextMatrix(Grid_Bitacora.RowSel, 1))
    MsgBox "El Correo de fecha: " & Format(Format(Grid_Bitacora.TextMatrix(Grid_Bitacora.RowSel, 1), "dd/MM/yyyy"), "MM/dd/yyyy") & ", se ha enviado satisfactoriamente", vbOKOnly + vbInformation, Me.Caption
    Grid_Bitacora.Rows = 0
    Consulta_Bitacora_Importacion (Cmb_Correos_Enviados.Text)
Conexion_Base.CommitTrans
    Exit Sub
HANDLER:
Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Bitacora_Importacion
    'DESCRIPCIÓN:           Realiza la consulta de la información de la bitacora
    'PARÁMETROS :
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Bitacora_Importacion(Tipo_Consulta As String)
Dim Rs_Consulta_Adm_Bitacora_Importacion As rdoResultset    'Informacion de la bitacora

Mi_SQL = "SELECT Consecutivo, Fecha, Hora_Ejecucion, Tipo_Importacion,Enviado"
Mi_SQL = Mi_SQL & " FROM Adm_Bitacora_Importacion"
Mi_SQL = Mi_SQL & " WHERE Fecha >= " & Par_Fecha & Format(Dtp_Importacion_Fecha_Inicio.Value, "MM/dd/yyyy") & Par_Fecha
Mi_SQL = Mi_SQL & " AND Fecha <= " & Par_Fecha & Format(Dtp_Importacion_Fecha_Termino.Value, "MM/dd/yyyy") & Par_Fecha

Select Case Tipo_Consulta
    Case "TODOS"
    Case "ENVIADOS"
        Mi_SQL = Mi_SQL & " AND Enviado = 'SI'"
    Case "SIN ENVIAR"
        Mi_SQL = Mi_SQL & " AND Enviado = 'NO'"
End Select
Mi_SQL = Mi_SQL & " ORDER BY Fecha"
Set Rs_Consulta_Adm_Bitacora_Importacion = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Adm_Bitacora_Importacion
    If Not .EOF Then
        Grid_Bitacora.Cols = 5
        If Grid_Bitacora.Rows = 0 Then
            Grid_Bitacora.AddItem "Consecutivo" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Importacion" & Chr(9) & "Enviado"
        End If
        While Not .EOF
            Grid_Bitacora.AddItem .rdoColumns("Consecutivo") & Chr(9) & Format(.rdoColumns("Fecha"), "MM/dd/yyyy") & Chr(9) & Format(.rdoColumns("Hora_Ejecucion"), "HH:mm:ss") & Chr(9) & .rdoColumns("Tipo_Importacion") & Chr(9) & .rdoColumns("Enviado")
            .MoveNext
        Wend
        .Close
        With Grid_Bitacora
            .FixedRows = 1
            .ColWidth(0) = 0 'Consecutivo
            .ColWidth(1) = 1000 'Fecha
            .ColWidth(2) = 1000 'Hora
            .ColWidth(3) = 2500 'Tipo Importacion
            .ColWidth(4) = 900 'Enviado
        End With
    End If
End With
Set Rs_Consulta_Adm_Bitacora_Importacion = Nothing
End Sub

Public Sub Inicializa()
    Dtp_Importacion_Fecha_Inicio.Value = Now
    Dtp_Importacion_Fecha_Termino.Value = Now
    Consulta_Bitacora_Importacion ("TODOS")
    Cmb_Correos_Enviados.ListIndex = 0
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Enviar_Correo
    'DESCRIPCIÓN:           Envia el correo con los parametros establecido
    'PARÁMETROS :           From_Email: correo de quien envia
    '                       Nombre_From: Nombre quien envia
    '                       To_Email:correo a quien se envia
    '                       Nombre_To: nombre a quien se envia
    '                       Asunto: asunto del correo
    '                       Mensaje_Email: mensaje del correo
    'CREO       :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO :           19 Mayo 2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Enviar_Correo(From_Email As String, Nombre_From As String, To_Email As String, Nombre_To As String, Asunto As String, Mensaje_Email As String)
    Set poSendMail = New clsSendMail
    Me.MousePointer = vbHourglass

    With poSendMail
        ' Propiedades opcionales para envio de correo, deberan ser primero configuradas si se utilizan
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
        ' Propiedades básicas para envio de correos
        .SMTPHost = Servidor_SMTP           ' Required the fist time, optional thereafter
        .From = From_Email                  ' Required the fist time, optional thereafte
        .FromDisplayName = Nombre_From      ' Optional, saved after first use
        .Recipient = To_Email               ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = Nombre_To   ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName          ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc              ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text      ' Optional, used when different than 'From' address
        .Subject = Asunto                   ' Optional
        .Message = Mensaje_Email            ' Optional
        '.Attachment = Trim(txtAttach.Text)  ' Optional, separate multiple entries with delimiter character

        ' Propiedades opcionales adicionales, utilizar si son requeridas por la aplicacion
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        '.UserName = txtUserName                     ' Optional, default = Null String
        '.Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' Propiedades avanzadas, cambiar solo si tienes una buena razon para hacerlos
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
         .SMTPPort = Puerto_SMTP                    ' Optional, default = 25

        ' Envio de correo
        ' .Connect                                  ' Optional, use when sending bulk mail
        .send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'txtServer.Text = .SMTPhost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Set poSendMail = Nothing
    Me.MousePointer = vbDefault
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Inserta_Registro_Bitacora_Importacion
    'DESCRIPCIÓN:          Inserta o modifica el registro de la importacion
    'PARÁMETROS:           Tipo: Verifica si es alta o modificacion
    '                      Fecha: Fecha del registro
    'CREO      :           Yañez Rodriguez Diego Neftali
    'FECHA_CREO:           27/Julio/2009
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Public Sub Inserta_Registro_Bitacora_Importacion(Fecha As Date)
Dim Rs_Adm_Bitacora As rdoResultset 'Permite agregar el registro de auditoria de sincronizacion

    Mi_SQL = "SELECT * FROM Adm_Bitacora_Importacion"
    Mi_SQL = Mi_SQL & " WHERE Fecha = " & Par_Fecha & Format(Fecha, "MM/dd/yyyy") & Par_Fecha
    Set Rs_Adm_Bitacora = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    With Rs_Adm_Bitacora
        .Edit
            .rdoColumns("Enviado") = "SI"
            .rdoColumns("Usuario_Modifico") = "SERVICIO IMPORTACION"
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Adm_Bitacora = Nothing
    
End Sub

