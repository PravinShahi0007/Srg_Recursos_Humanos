VERSION 5.00
Begin VB.Form Frm_Adm_Enrollment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Huella Digital"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8925
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
      TabIndex        =   7
      Top             =   15
      Width           =   8805
      Begin VB.TextBox Txt_Cat_Empleados_Empleado_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   150
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
         TabIndex        =   10
         Top             =   750
         Width           =   6420
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
         TabIndex        =   9
         Top             =   165
         Width           =   3060
      End
      Begin VB.Image Img_Cat_Empleados_Foto 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1845
         Left            =   6645
         Picture         =   "Frm_Adm_Enrollment.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.PictureBox Pic_Oculto 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   5475
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List_Estatus 
      Height          =   2985
      ItemData        =   "Frm_Adm_Enrollment.frx":C042
      Left            =   3735
      List            =   "Frm_Adm_Enrollment.frx":C044
      TabIndex        =   3
      Top             =   2280
      Width           =   5010
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   7320
      TabIndex        =   2
      Top             =   5385
      Width           =   1320
   End
   Begin VB.PictureBox Pic_Huella 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   120
      ScaleHeight     =   2925
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   2265
      Width           =   3270
   End
   Begin VB.Label Lbl_Cantidad_Muestras 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cantidad de Muestras:"
      Height          =   195
      Left            =   780
      TabIndex        =   6
      Top             =   5385
      Width           =   1590
   End
   Begin VB.Label Lbl_Samples 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2745
      TabIndex        =   4
      Top             =   5295
      Width           =   615
   End
   Begin VB.Label Lbl_Prompt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   570
      TabIndex        =   1
      Top             =   5775
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "Frm_Adm_Enrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim Crear_Feature As DPFPFeatureExtraction
Dim Crear_Template As DPFPEnrollment
Dim Convertir_Sample As DPFPSampleConversion
Dim Template_Empleado As DPFPTemplate

Private Sub Dibuja_Imagen(ByVal Imagen As IPictureDisp)
    Set Pic_Oculto.picture = Imagen
    Pic_Huella.PaintPicture Pic_Oculto.picture, 0, 0, Pic_Huella.ScaleWidth, Pic_Huella.ScaleHeight, 0, 0, Pic_Oculto.ScaleWidth, Pic_Oculto.ScaleHeight, vbSrcCopy
    Pic_Huella.picture = Pic_Huella.Image
End Sub
Private Sub Reportar_Estatus(ByVal Texto As String)
    List_Estatus.AddItem (Texto)
    List_Estatus.ListIndex = List_Estatus.NewIndex
End Sub

Private Sub Btn_Salir_Click()
    Capture.StopCapture
    Unload Me
End Sub

Private Sub Inicializar_Captura()
    'Create capture operation.
    Set Capture = New DPFPCapture
    'Start capture operation.
    Capture.StartCapture
    'Create DPFPFeatureExtraction object.
    Set Crear_Feature = New DPFPFeatureExtraction
    'Create DPFPEnrollment object.
    Set Crear_Template = New DPFPEnrollment
    'Show number of Lbl_Samples needed.
    Lbl_Samples.Caption = Crear_Template.FeaturesNeeded
    'Create DPFPSampleConversion object.
    Set Convertir_Sample = New DPFPSampleConversion
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella está conectado")
End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella está desconectado")
End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El lector de huella fué tocado")
End Sub

Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
    Reportar_Estatus ("El dedo fue removido del lector de huella")
End Sub

Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, ByVal Feedback As DPFPCaptureFeedbackEnum)
    If Feedback = CaptureFeedbackGood Then
        Reportar_Estatus ("La calidad de la huella es buena")
    Else
        Reportar_Estatus ("La calidad de la huella es pobre")
    End If
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
Dim Feedback As DPFPCaptureFeedbackEnum
Dim Resultado As DPFPVerificationResult
Dim Template_Verificacion As Object
 
    Reportar_Estatus ("La huella ha sido capturada")
    'Draw fingerprint image.
    Call Dibuja_Imagen(Convertir_Sample.ConvertToPicture(Sample))
    'Process sample and create feature set for purpose of enrollment.
    Feedback = Crear_Feature.CreateFeatureSet(Sample, DataPurposeEnrollment)
    'Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
        Reportar_Estatus ("Los componentes de la huella han sido creados")
        Lbl_Prompt.Caption = "Toque el lector de huella otra vez con el mismo dedo"
        'Add feature set to template.
        Crear_Template.AddFeatures Crear_Feature.FeatureSet
        'Show number of Lbl_Samples needed to complete template.
        Lbl_Samples.Caption = Crear_Template.FeaturesNeeded
        'Check if template has been created.
        If Crear_Template.TemplateStatus = TemplateStatusTemplateReady Then
            Call Set_Template(Crear_Template.Template)
            'Template has been created, so stop capturing Lbl_Samples.
            Capture.StopCapture
            Guardar_Template
            Lbl_Prompt.Caption = "Click en Salir, y verifique la lectura de su huella"
            MsgBox "El registro de la huella ha sido creado", vbInformation
            Unload Me
        End If
    End If
End Sub

Public Sub Set_Template(ByVal Template As Object)
    Set Template_Empleado = Template
End Sub

Public Function Get_Template() As Object
    'Template can be empty. If so, then returns Nothing.
    If Template_Empleado Is Nothing Then
    Else
        Set Get_Template = Template_Empleado
    End If
End Function

Private Sub Guardar_Template()
Dim Template_Bytes() As Byte
Dim Ruta_Almacenamiento As String
Dim Rs_Actualiza_Huella As rdoResultset
Dim Rs_Alta_Huella As rdoResultset

On Error GoTo Fin
    'First verify that template is not empty.
    If Template_Empleado Is Nothing Then
        MsgBox "Verifique que la huella se haya creado para poder continuar", vbExclamation
        Exit Sub
    End If
    'Export template to binary data.
    Template_Bytes = Template_Empleado.Serialize
    Ruta_Almacenamiento = PG_Ruta_Huellas & "\" & Trim(Lbl_No_Empleado.Caption) & ".fpt"
    'Save binary data to file.
    Open Ruta_Almacenamiento For Binary As #1
    Put #1, , Template_Bytes
    Close #1
    
    'Busca si existe el registro del empleado
    Mi_SQL = "SELECT * FROM Cat_Empleados_Huellas"
    Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Trim(Txt_Cat_Empleados_Empleado_ID.Text) & "'"
    Set Rs_Actualiza_Huella = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    If Not Rs_Actualiza_Huella.EOF Then
        Rs_Actualiza_Huella.Edit
            Rs_Actualiza_Huella.rdoColumns("No_Tarjeta") = Val(Lbl_No_Empleado.Caption)
            Rs_Actualiza_Huella.rdoColumns("Huella_Digital") = Template_Bytes
            Rs_Actualiza_Huella.rdoColumns("Huella_Ruta") = Trim(Lbl_No_Empleado.Caption) & ".fpt"
        Rs_Actualiza_Huella.Update
    Else    'No lo encontró lo da de alta
        Set Rs_Alta_Huella = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados_Huellas")
            Rs_Alta_Huella.AddNew
                Rs_Alta_Huella.rdoColumns("Empleado_ID") = Trim(Txt_Cat_Empleados_Empleado_ID.Text)
                Rs_Alta_Huella.rdoColumns("No_Tarjeta") = Val(Lbl_No_Empleado.Caption)
                Rs_Alta_Huella.rdoColumns("Huella_Digital") = Template_Bytes
                Rs_Alta_Huella.rdoColumns("Huella_Ruta") = Trim(Lbl_No_Empleado.Caption) & ".fpt"
            Rs_Alta_Huella.Update
        Rs_Alta_Huella.Close
    End If
    Rs_Actualiza_Huella.Close
Exit Sub
Fin:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If
End Sub

Private Sub Leer_Template()
Dim Template_Bytes() As Byte
Dim Ruta_Almacenamiento As String

    Ruta_Almacenamiento = PG_Ruta_Huellas & "\" & "5.fpt"
    'Read binary data from file
    Open Ruta_Almacenamiento For Binary As #1
    ReDim Template_Bytes(LOF(1))
    Get #1, , Template_Bytes()
    Close #1
    'Template can be empty, it must be created first.
    If Template_Empleado Is Nothing Then
        Set Template_Empleado = New DPFPTemplate
    End If
    'Import binary data to template.
    Template_Empleado.Deserialize Template_Bytes
End Sub

Private Sub Form_Load()
    Inicializar_Captura
End Sub
