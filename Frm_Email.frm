VERSION 5.00
Begin VB.Form Frm_Email 
   Caption         =   "Enviar Correo"
   ClientHeight    =   660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   660
   ScaleWidth      =   3960
   Begin VB.PictureBox SMTP 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Enviando Correos"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Frm_Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mensaje As String
Public Correo As String
Public Asunto As String
Private Correro_Correo_SMTP As String
Private Contraseña_Correo_SMTP As String
Private Puerto_Correo_SMTP As String
Private Servidor_Correo_SMTP As String
Public Sub Obtener_Parametros_Correos()
Dim Rs_Consulta_Cat_Parametros As rdoResultset

On Error GoTo HANDLER
    Mi_SQL = "SELECT * FROM Cat_Parametros"
    Set Rs_Consulta_Cat_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Parametros
        If Not .EOF Then
            Correro_Correo_SMTP = Trim(.rdoColumns("Email_Correo"))
            Contraseña_Correo_SMTP = (.rdoColumns("Contrasenia_Correos"))
            Servidor_Correo_SMTP = Trim(.rdoColumns("Servidor_Correos"))
            Puerto_Correo_SMTP = Trim(.rdoColumns("Puerto_Correos"))
        End If
    End With
    Set Rs_Consulta_Cat_Parametros = Nothing
Exit Sub
HANDLER:
    MsgBox Err.Description
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Public Function Mandar_Correo() As Boolean
Dim Obj_Email As CDO.Message                    'Variable de objeto Cdo.Message
Set Obj_Email = New CDO.Message                 'Crea un Nuevo objeto CDO.Message
Dim Usar_Autentificacion As Boolean
Usar_Autentificacion = True
Dim Enviar_Mail_CDO As Boolean
    If Trim(Correo) <> "" Then
        ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
          del servidor o su dirección IP )
'        Obj_Email.Configuration.Fields(cdoSMTPServer) = "smtp.gmail.com"
        Obj_Email.Configuration.Fields(cdoSMTPServer) = Servidor_Correo_SMTP
        Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
          
        ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
          465 o  el puerto 587 ( este último me dio error )
        Obj_Email.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto_Correo_SMTP)
          
        ' Indica el tipo de autentificación con el servidor de correo _
         El valor 0 no requiere autentificarse, el valor 1 es con autentificación
        Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                    "configuration/smtpauthenticate") = Abs(True)
          
            ' Tiempo máximo de espera en segundos para la conexión
        Obj_Email.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
          
        ' Configura las opciones para el login en el SMTP
        If Usar_Autentificacion Then
          ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
           mas el @gmail.com )
          Obj_Email.Configuration.Fields.Item _
              ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Correro_Correo_SMTP
        
          ' Password de la cuenta
          Obj_Email.Configuration.Fields.Item _
              ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Contraseña_Correo_SMTP
        
          ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
          Obj_Email.Configuration.Fields.Item _
              ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        End If
          
        ' *********************************************************************************
        ' Estructura del mail
        '**********************************************************************************
        ' Dirección del Destinatario
        Obj_Email.To = Correo
        ' Dirección del remitente
        Obj_Email.From = "restaurante.socios@gmail.com"
          
        ' Asunto del mensaje
        Obj_Email.Subject = Asunto
          
        ' Cuerpo del mensaje
        Obj_Email.TextBody = Mensaje
          
'        'Ruta del archivo adjunto
'        Obj_Email.AddAttachment "C:\Reportes\Corte_Caja.txt"
'
        ' Actualiza los datos antes de enviar
        Obj_Email.Configuration.Fields.Update
          
        On Error Resume Next
        ' Envía el email
        Obj_Email.send
                
        If Err.Number = 0 Then
            Enviar_Mail_CDO = True
            Mandar_Correo = True
        Else
            Enviar_Mail_CDO = False
            Mandar_Correo = False
            MsgBox Err.Description, vbCritical, " Error al enviar el amil "
        End If
          
        If Enviar_Mail_CDO Then
'            MsgBox " .. Meneje enviado ", vbInformation
        End If
        ' Descarga la referencia
        If Not Obj_Email Is Nothing Then
            Set Obj_Email = Nothing
        End If
    End If
    'Call MailTo(Correo_Electronico, Correo_Electronico, "Corte de Caja Turno: " & No_Turno, "C:\Corte_Caja.txt")
End Function

