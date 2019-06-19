VERSION 5.00
Begin VB.Form Frm_Adm_Entrada_Comedor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRADAS AL COMEDOR"
   ClientHeight    =   9015
   ClientLeft      =   4050
   ClientTop       =   4020
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   8925
   Begin VB.CheckBox Chk_Imprimir_Ticket 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir Ticket"
      Height          =   195
      Left            =   3570
      TabIndex        =   15
      Top             =   8775
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.ListBox Lst_Estatus 
      Height          =   2205
      ItemData        =   "Frm_Adm_Entrada_Comedor.frx":0000
      Left            =   3600
      List            =   "Frm_Adm_Entrada_Comedor.frx":0002
      TabIndex        =   12
      Top             =   6180
      Width           =   4575
   End
   Begin VB.PictureBox Pic_Huella 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   495
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   6135
      Width           =   2775
   End
   Begin VB.PictureBox Pic_Oculto 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8220
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   8235
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Tmr_Limpiar_Datos 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   60
      Top             =   6270
   End
   Begin VB.PictureBox Pic_Entradas_Comedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   -15
      ScaleHeight     =   6165
      ScaleWidth      =   8925
      TabIndex        =   0
      Top             =   -15
      Width           =   8925
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   405
         Left            =   7170
         TabIndex        =   9
         Top             =   5430
         Width           =   1485
      End
      Begin VB.Frame Fra_Cat_Empleados_Datos_Personales 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Left            =   60
         TabIndex        =   2
         Top             =   465
         Width           =   8805
         Begin VB.TextBox Txt_Cat_Empleados_Empleado_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   90
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Image Img_Cat_Empleados_Foto 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1845
            Left            =   6645
            Picture         =   "Frm_Adm_Entrada_Comedor.frx":0004
            Stretch         =   -1  'True
            ToolTipText     =   "Doble click para cambiar la imagen"
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label Lbl_No_Empleado 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Empleado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   105
            TabIndex        =   5
            Top             =   165
            Width           =   3060
         End
         Begin VB.Label Lbl_Nombre_Empleado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre Empleado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   105
            TabIndex        =   4
            Top             =   750
            Width           =   6420
         End
      End
      Begin VB.TextBox Txt_No_Empleado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2865
         TabIndex        =   1
         Top             =   6315
         Width           =   3495
      End
      Begin VB.Label Lbl_Entradas_Comedor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ENTRADAS AL COMEDOR"
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
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8775
      End
      Begin VB.Label Lbl_Fecha_Hora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha y Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   7
         Top             =   2670
         Width           =   8700
      End
      Begin VB.Label Lbl_Mensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingrese su Huella"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   90
         TabIndex        =   6
         Top             =   3345
         Width           =   8775
      End
   End
   Begin VB.Label Lbl_Rango_Aceptacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rango de Aceptación:"
      Height          =   195
      Left            =   3555
      TabIndex        =   14
      Top             =   8505
      Width           =   1605
   End
   Begin VB.Label Lbl_FAR 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5475
      TabIndex        =   13
      Top             =   8445
      Width           =   2655
   End
End
Attribute VB_Name = "Frm_Adm_Entrada_Comedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim Crear_Features As DPFPFeatureExtraction
Dim Verificacion As DPFPVerification
Dim Convertir_Sample As DPFPSampleConversion
Dim Template_BD As DPFPTemplate
Dim Segundos_Espera As Integer

'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Empleado
'DESCRIPCION: Consulta el empleado en la base de datos y valida el estatus para el comedor
'PARAMETROS : Texto_Busqueda- Es el código por el que va a reconocer al empleado
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 31-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Consulta_Empleado(Texto_Busqueda As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Rs_Consulta_Adm_Entradas_Comedor As rdoResultset

On Error GoTo HANDLER
    'Consulta el empleado y valida si existe o está activo
    Mi_SQL = "SELECT Empleado_ID,(Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre,No_Tarjeta,Estatus,Imagen_Perfil"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE No_Tarjeta=" & Val(Texto_Busqueda)
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cat_Empleados.EOF Then
        With Rs_Consulta_Cat_Empleados
            Lbl_No_Empleado.Caption = .rdoColumns("No_Tarjeta")
            Lbl_Nombre_Empleado.Caption = .rdoColumns("Nombre")
            Lbl_Fecha_Hora.Caption = Format(Now, "dddd, dd MMMM yyyy HH:mm:ss")
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(PG_Ruta_Fotos & "\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                Img_Cat_Empleados_Foto.picture = LoadPicture(PG_Ruta_Fotos & "\" & .rdoColumns("Imagen_Perfil"))
            Else
                Img_Cat_Empleados_Foto.picture = LoadPicture("")
            End If
            'Valida el estatus
            If .rdoColumns("Estatus") = "A" Then
                'Valida el parámetro de comidas si es 0 no hay restriccion
                If PG_Cantidad_Comidas > 0 Then
                    'Valida los accesos que ha tenido en el comedor en el día
                    Mi_SQL = "SELECT ISNULL(COUNT(*),0) AS Registros FROM Adm_Entradas_Comedor"
                    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & .rdoColumns("Empleado_ID") & "'"
'                    Mi_SQL = Mi_SQL & " AND Fecha BETWEEN '" & Format(Now, "MM/dd/yyyy") & " 00:00:00' AND '" & Format(Now, "MM/dd/yyyy") & " 23:59:59'"
                    Mi_SQL = Mi_SQL & " AND Fecha >= CONVERT(VARCHAR, GETDATE(), 112)"
                    Mi_SQL = Mi_SQL & " AND Fecha < CONVERT(VARCHAR, DATEADD(DAY, 1, GETDATE()), 112)"
                    Set Rs_Consulta_Adm_Entradas_Comedor = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If PG_Cantidad_Comidas > Rs_Consulta_Adm_Entradas_Comedor.rdoColumns("Registros") Then
                        Lbl_Mensaje.Caption = "Bienvenido, ¡¡¡buen provecho!!!!"
                        Lbl_Mensaje.ForeColor = vbBlack
                        'Almacena en la base de datos el registro de acceso a la comida
                        Call Inserta_Registro_Comida(.rdoColumns("Empleado_ID"))
                        'Imprime el ticket
                        If Chk_Imprimir_Ticket.Value = 1 Then
                            Imprimir_Tickets
                        End If
                        Segundos_Espera = 0
                        Tmr_Limpiar_Datos.Enabled = True
                        Exit Sub
                    Else
                        Lbl_Mensaje.Caption = "Ya se había registrado una comida previamente"
                        Lbl_Mensaje.ForeColor = vbRed
                        Txt_No_Empleado.Text = ""
                        Segundos_Espera = 0
                        Tmr_Limpiar_Datos.Enabled = True
                        Exit Sub
                    End If
                End If
                Lbl_Mensaje.Caption = "Bienvenido, ¡¡¡buen provecho!!!!"
                Lbl_Mensaje.ForeColor = vbBlack
                'Almacena en la base de datos el registro de acceso a la comida
                Call Inserta_Registro_Comida(.rdoColumns("Empleado_ID"))
                'Imprime el ticket
                If Chk_Imprimir_Ticket.Value = 1 Then
                    Imprimir_Tickets
                End If
            Else    'El empleado no está activo
                Lbl_Mensaje.Caption = "El empleado no se encuentra activo"
                Lbl_Mensaje.ForeColor = vbRed
            End If
        End With
    Else    'No encontró empleado
        Lbl_No_Empleado.Caption = ""
        Lbl_Nombre_Empleado.Caption = ""
        Img_Cat_Empleados_Foto.picture = LoadPicture("")
        Lbl_Fecha_Hora.Caption = Format(Now, "dddd, dd MMMM yyyy HH:mm:ss")
        Lbl_Mensaje.Caption = "No se encuentra registrado el empleado"
        Lbl_Mensaje.ForeColor = vbRed
    End If
    Rs_Consulta_Cat_Empleados.Close
    Txt_No_Empleado.Text = ""
    Segundos_Espera = 0
    Tmr_Limpiar_Datos.Enabled = True
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Inserta_Registro_Comida
'DESCRIPCION: Inserta un registro de comida del empleado registrado
'PARAMETROS : Empleado_ID- Es el ID del empleado que registra una entrada a la comida
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 31-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Inserta_Registro_Comida(Empleado_ID As String)
Dim Rs_Alta_Adm_Entradas_Comedor As rdoResultset

On Error GoTo HANDLER
    'Valida los accesos que ha tenido en el comedor en el día
    Set Rs_Alta_Adm_Entradas_Comedor = Conectar_Ayudante.Recordset_Agregar("Adm_Entradas_Comedor")
    With Rs_Alta_Adm_Entradas_Comedor
        .AddNew
            .rdoColumns("Empleado_ID") = Empleado_ID
            .rdoColumns("Fecha") = Format(Now, "MM/dd/yyyy")
            .rdoColumns("Hora") = Format(Now, "HH:mm:ss")
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha") = Now
        .Update
    End With
    Rs_Alta_Adm_Entradas_Comedor.Close
    Txt_No_Empleado.Text = ""
Exit Sub
HANDLER:
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Salir_Click()
    'Datiene la captura del lector de huella
    Capture.StopCapture
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Lbl_No_Empleado.Caption = ""
    Lbl_Nombre_Empleado.Caption = ""
    Lbl_Fecha_Hora.Caption = Format(Now, "dddd, dd MMMM yyyy")
    Chk_Imprimir_Ticket.Value = PG_Imprime_Comidas
    'Inicializa las librerías del lector
    'Create capture operation.
    Set Capture = New DPFPCapture
    'Start capture operation.
    Capture.StartCapture
    'Create DPFPFeatureExtraction object.
    Set Crear_Features = New DPFPFeatureExtraction
    'Create DPFPVerification object.
    Set Verificacion = New DPFPVerification
    'Create DPFPSampleConversion object.
    Set Convertir_Sample = New DPFPSampleConversion
End Sub

Private Sub Tmr_Limpiar_Datos_Timer()
    If Segundos_Espera = 5 Then
        Lbl_No_Empleado.Caption = ""
        Lbl_Nombre_Empleado.Caption = ""
        Img_Cat_Empleados_Foto.picture = LoadPicture("")
        Lbl_Fecha_Hora.Caption = Format(Now, "dddd, dd MMMM yyyy")
        Lbl_Mensaje.Caption = "Ingrese su Huella"
        Lbl_Mensaje.ForeColor = vbBlack
        Txt_No_Empleado.Text = ""
        Segundos_Espera = 0
        Tmr_Limpiar_Datos.Enabled = False
    Else
        Segundos_Espera = Segundos_Espera + 1
    End If
End Sub

Private Sub Txt_No_Empleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Consulta_Empleado(Txt_No_Empleado.Text)
    Else
        Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
    End If
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella está conectado")
End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella está desconectado")
End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella fue tocado")
End Sub
Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El dedo fue removido del lector de huella")
End Sub
Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, ByVal Feedback As DPFPCaptureFeedbackEnum)
    If Feedback = CaptureFeedbackGood Then
        Reportar_Estatus ("La calidad de muestra del lector es buena")
    Else
        Reportar_Estatus ("La calidad de muestra del lector es pobre")
    End If
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
Dim Feedback As DPFPCaptureFeedbackEnum
Dim Resultado As DPFPVerificationResult
Dim Template_Consulta As Object
Dim Template_Imagen() As Byte
Dim Rs_Consulta_Huellas As rdoResultset
Dim Ruta_Almacenamiento As String

    Reportar_Estatus ("La huella ha sido capturada")
    'Draw fingerprint image.
    Dibujar_Imagen Convertir_Sample.ConvertToPicture(Sample)
    'Process sample and create feature set for purpose of verification.
    Feedback = Crear_Features.CreateFeatureSet(Sample, DataPurposeVerification)
    'Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
        'Consulta los registros de huella digital
        Mi_SQL = "SELECT Empleado_ID,No_Tarjeta,Huella_Ruta,Huella_Digital"
        Mi_SQL = Mi_SQL & " FROM Cat_Empleados_Huellas"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID in (SELECT Empleado_ID FROM Cat_Empleados ce WHERE ce.Estatus != 'I')"
        Set Rs_Consulta_Huellas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        While Not Rs_Consulta_Huellas.EOF
            
'            If Not IsNull(Rs_Consulta_Huellas.rdoColumns("Huella_Digital")) Then
                'Lectura desde la base de datos
                Template_Imagen = Rs_Consulta_Huellas.rdoColumns("Huella_Digital")
'            Else
                'Lectura desde un archivo físico
                'Ruta_Almacenamiento = PG_Ruta_Huellas & "\" & Rs_Consulta_Huellas.rdoColumns("Huella_Ruta")
                'Read binary data from file.
                'Open Ruta_Almacenamiento For Binary As #1
                '    ReDim Template_Imagen(LOF(1))
                '    Get #1, , Template_Imagen()
                'Close #1
'            End If

            'Template can be empty, it must be created first.
            If Template_BD Is Nothing Then Set Template_BD = New DPFPTemplate
            'Import binary data to template.
            Template_BD.Deserialize Template_Imagen
            Set Template_Consulta = Template_BD
            'Compare feature set with template.
            Set Resultado = Verificacion.Verify(Crear_Features.FeatureSet, Template_Consulta)
            'Show results of comparison.
            Lbl_FAR.Caption = Resultado.FARAchieved
            If Resultado.Verified = True Then
                Reportar_Estatus ("La huella ha sido verificada")
                'Ejecuta la acción de buscar el empleado registrado
                Txt_No_Empleado.Text = Rs_Consulta_Huellas.rdoColumns("No_Tarjeta")
                Call Txt_No_Empleado_KeyPress(13)
                Rs_Consulta_Huellas.Close
                Exit Sub
            End If
            Rs_Consulta_Huellas.MoveNext
        Wend
        Rs_Consulta_Huellas.Close
        Reportar_Estatus ("No fue verificada la huella digital")
        Lbl_Mensaje.Caption = "No se encontró la huella, intente de nuevo"
        Lbl_Mensaje.ForeColor = vbRed
    Else
        Reportar_Estatus ("La calidad de muestra del lector es pobre")
    End If
End Sub

Public Function Get_Template() As Object
    'Template can be empty. If so, then returns Nothing.
    If Template_BD Is Nothing Then
    Else
        Set Get_Template = Template_BD
    End If
End Function

Private Sub Reportar_Estatus(ByVal Texto As String)
    Lst_Estatus.AddItem (Texto)
    Lst_Estatus.ListIndex = Lst_Estatus.NewIndex
End Sub

Private Sub Dibujar_Imagen(ByVal Imagen As IPictureDisp)
    ' Must use hidden PictureBox to easily resize picture.
    Set Pic_Oculto.picture = Imagen
    Pic_Huella.PaintPicture Pic_Oculto.picture, 0, 0, Pic_Huella.ScaleWidth, Pic_Huella.ScaleHeight, 0, 0, Pic_Oculto.ScaleWidth, Pic_Oculto.ScaleHeight, vbSrcCopy
    Pic_Huella.picture = Pic_Huella.Image
End Sub

Private Sub Leer_Template_BD()
Dim Template_Imagen() As Byte
Dim Rs_Consulta_Huellas As rdoResultset
Dim Ruta_Almacenamiento As String
    
    'Consulta los registros de huella digital
    Mi_SQL = "SELECT * FROM Cat_Empleados_Huellas"
    Set Rs_Consulta_Huellas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Huellas.EOF
        
        Ruta_Almacenamiento = PG_Ruta_Huellas & "\" & Rs_Consulta_Huellas.rdoColumns("Huella_Ruta")
    
        'Read binary data from file.
        Open Ruta_Almacenamiento For Binary As #1
            ReDim Template_Imagen(LOF(1))
            Get #1, , Template_Imagen()
        Close #1
        'Template can be empty, it must be created first.
        If Template_BD Is Nothing Then Set Template_BD = New DPFPTemplate
        'Import binary data to template.
        Template_BD.Deserialize Template_Imagen
        
        Rs_Consulta_Huellas.MoveNext
    Wend
    Rs_Consulta_Huellas.Close
End Sub

Private Sub Imprimir_Tickets()
Dim Fila As Double                   'Indica el número de la fila del grid_productos que se esta consultando
Dim Impresora As String              'Tomna el nombre la impresora
Dim Mi_Impresora As Printer          'Toma el nombre de la impresora
Dim Ubicacion_Impresora As String    'Toma el valor de la ubicacion dela impresora

On Error GoTo HANDLER
    Me.MousePointer = 11
'    'Printer.DeviceName = Impresora  'Cambiamos la impresora a la que se va a mandar
'    For Each Mi_Impresora In Printers
'        If UCase(Mi_Impresora.DeviceName) Like "*" & UCase(PG_Impresora_Comidas) & "*" Then
'            Set Printer = Mi_Impresora
'            Exit For
'        End If
'    Next
    Debug.Print Printer.DeviceName
    'Comienza la impresión del ticket
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8
    Printer.Print Conectar_Ayudante.Centra("ENTRADA A COMEDOR", Len("---------------------------------------"))
    Printer.Print
    Printer.Print Conectar_Ayudante.Centra(Empresa, Len("---------------------------------------"))
    Printer.Print Conectar_Ayudante.Centra(RFC, Len("---------------------------------------"))
    Printer.Print Conectar_Ayudante.Centra(Direccion, Len("---------------------------------------"))
    Printer.Print Conectar_Ayudante.Centra(Ciudad_Edo, Len("---------------------------------------"))
    Printer.Print
    Printer.Print "---------------------------------------"
    Printer.Print "No. Empleado: ..........."; Conectar_Ayudante.Alinea_Derecha(Lbl_No_Empleado.Caption, 12)
    Printer.Print Mid(Lbl_Nombre_Empleado.Caption, 1, Len("---------------------------------------"))
    Printer.Print
    Printer.Print "Fecha : ................."
    Printer.Print Mid(Lbl_Fecha_Hora.Caption, 1, Len("---------------------------------------"))
    Printer.Print
    Printer.Print Conectar_Ayudante.Centra("¡¡¡Buen Provecho!!!", Len("---------------------------------------"))
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "."
    Printer.EndDoc
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    MsgBox Err.Description
End Sub

