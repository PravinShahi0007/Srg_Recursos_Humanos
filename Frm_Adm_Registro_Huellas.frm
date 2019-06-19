VERSION 5.00
Begin VB.Form Frm_Adm_Registro_Huellas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE HUELLAS"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Limpiar_Eventos 
      Caption         =   "Limpiar"
      Height          =   405
      Left            =   30
      TabIndex        =   0
      Top             =   6885
      Width           =   1515
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "Cerrar"
      Height          =   405
      Left            =   5865
      TabIndex        =   3
      Top             =   6885
      Width           =   1515
   End
   Begin VB.Frame Fra_Eventos 
      Caption         =   "Eventos"
      Height          =   2145
      Left            =   0
      TabIndex        =   1
      Top             =   4695
      Width           =   7395
      Begin VB.ListBox ListEvents 
         Height          =   1815
         Left            =   60
         TabIndex        =   2
         Top             =   195
         Width           =   7230
      End
   End
End
Attribute VB_Name = "Frm_Adm_Registro_Huellas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxFingers As Integer
Dim EnrolledFingersMask As Integer
Dim MaxEnrollFingerCount As Integer
Dim IsEventHandlerSucceeds As Boolean
Dim IsFeatureSetMatched As Boolean
Dim FalseAcceptRate As Integer
Dim Templates(9) As DPFPTemplate
Dim Templates_Enrroll As DPFPEnrollment

Private Sub Btn_Limpiar_Eventos_Click()
    ListEvents.Clear
End Sub

Private Sub EnrollmentControl_OnCancelEnroll(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnCancelEnroll: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnComplete(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnComplete: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnDelete(ByVal lFingerMask As Long, ByVal pStatus As Object)
    If (IsEventHandlerSucceeds) Then
        ListEvents.AddItem (Format("OnDelete: finger {0}"))
    End If
End Sub

Private Sub EnrollmentControl_OnEnroll(ByVal lFingerMask As Long, ByVal pTemplate As Object, ByVal pStatus As Object)
Dim ID As Long
Dim aRawData() As Byte

    Set Templates(lFingerMask) = pTemplate
    aRawData = Templates(lFingerMask).Serialize
    
    'Call DB.Inserta_Access(aRawData, 5)
    
    'ID = DB.AddTemplate(aRawData, 5)
    
    Mi_SQL = "UPDATE Cat_Empleados SET Huella_Empleado=" & Templates(lFingerMask).Serialize
    Mi_SQL = Mi_SQL & " WHERE No_Tarjeta=5"
    Conexion_Base.Execute Mi_SQL
        
    
End Sub

Private Sub EnrollmentControl_OnFingerRemove(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnFingerRemove: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnFingerTouch(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnFingerTouch: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnReaderConnect(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnReaderConnect: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnReaderDisconnect(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (Format("OnReaderDisconnect: {0}, finger {1}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnSampleQuality(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long, ByVal lSampleQuality As Long)
    ListEvents.AddItem (Format("OnSampleQuality: {0}, finger {1}, {2}", pSerialNumber))
End Sub

Private Sub EnrollmentControl_OnStartEnroll(ByVal pSerialNumber As String, ByVal lEnrolledFinger As Long)
    ListEvents.AddItem (pSerialNumber)
End Sub

Private Sub Form_Load()
    MaxFingers = 10
    EnrolledFingersMask = 0
    MaxEnrollFingerCount = MaxFingers
    IsEventHandlerSucceeds = True
    IsFeatureSetMatched = False
    FalseAcceptRate = 0
End Sub

