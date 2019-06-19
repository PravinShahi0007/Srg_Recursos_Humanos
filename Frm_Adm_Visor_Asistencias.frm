VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Adm_Visor_Asistencias 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8235
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   75
      Left            =   105
      OleObjectBlob   =   "Frm_Adm_Visor_Asistencias.frx":0000
      TabIndex        =   6
      Top             =   5685
      Width           =   45
   End
   Begin VB.CommandButton Btn_Salir_2 
      Caption         =   "Salir"
      Height          =   690
      Left            =   6960
      Picture         =   "Frm_Adm_Visor_Asistencias.frx":0024
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5340
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.ComboBox Cmb_Dipositivos 
      Height          =   315
      ItemData        =   "Frm_Adm_Visor_Asistencias.frx":05AE
      Left            =   1110
      List            =   "Frm_Adm_Visor_Asistencias.frx":05B0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   6675
   End
   Begin VB.Frame Fra_Asistencias 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registros"
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid Grid_Registro_Checadas 
         Height          =   4515
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7964
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         ScrollBars      =   2
         Appearance      =   0
      End
   End
   Begin VB.Label Lbl_Estatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   5400
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dispositivo"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "Frm_Adm_Visor_Asistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IP As String
Dim Puerto As Long
Dim Maquina As Long

Private Sub Btn_Salir_2_Click()
    CZKEM1.Disconnect
    Unload Me
End Sub

Private Sub Cmb_Dipositivos_Click()
    If Cmb_Dipositivos.ListIndex > -1 Then
        'Llena informacion del dispositivo
        IP = ""
        Puerto = 0
        Maquina = 0
        CZKEM1.Disconnect
        Llena_Informacion_Dispositivo
        'If CZKEM1.Connect_Net(CStr(IP), Val(Puerto)) Then
        If CZKEM1.Connect_Net("172.16.1.25", 4370) Then
            Lbl_Estatus.Caption = "Conectado ..."
            Exit Sub
        Else
            Lbl_Estatus.Caption = "DesConectado ..."
            MsgBox "No se pudo establecer la conexion"
        End If
    End If
End Sub

Public Sub Inicializa()
    Call Conectar_Ayudante.Llena_Combo_Item("Equipo_ID, (CAST(No_Equipo as varchar) +' '+ Descripcion) as Equipo", "Cat_Equipos_Identificadores", Cmb_Dipositivos, 0, "No_Equipo")
    Grid_Registro_Checadas.Rows = 0
    Grid_Registro_Checadas.Cols = 6
    Consulta_Entradas_Registradas
End Sub

Private Sub CZKEM1_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
Dim Rs_Agrega_Detalles_Entradas_Equipos As rdoResultset     'Informacion del registro capturado
Dim Nombre_Empleado As String
Dim Empleado_ID_Registro As String
Dim Metodo_Verificacion As String
Dim E_S As String

On Error GoTo HANDLER

    If Grid_Registro_Checadas.Rows = 0 Then
        Grid_Registro_Checadas.AddItem "Usuario" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Tipo Acceso" & Chr(9) & "E/S"
    End If
    'Obtiene el nombre del usuario
    Mi_SQL = "SELECT CE.Empleado_ID, (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre "
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE"
    Mi_SQL = Mi_SQL & " WHERE No_Tarjeta = '" & Val(EnrollNumber) & "'"
    Nombre_Empleado = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Nombre")
    Empleado_ID_Registro = Conectar_Ayudante.Busca_Dato_BD(Mi_SQL, "Empleado_ID")
    'Verifica como acceso
    Select Case VerifyMethod
        Case 0: Metodo_Verificacion = "Manual"
        Case 1: Metodo_Verificacion = "Huella Digital"
    End Select
    'Verifica como E/S
    Select Case AttState
        Case 0: E_S = "E"
        Case 1: E_S = "S"
    End Select
    If Nombre_Empleado = "" Then Nombre_Empleado = "No registrado"
    Grid_Registro_Checadas.AddItem EnrollNumber & Chr(9) & Nombre_Empleado & Chr(9) & _
            Month & "/" & Day & "/" & Year & Chr(9) & Hour & ":" & Minute & ":" & Second & Chr(9) & Metodo_Verificacion & Chr(9) & E_S

    'Guarda el registro en la base de datos
    Conexion_Base.BeginTrans
        Set Rs_Agrega_Detalles_Entradas_Equipos = Conectar_Ayudante.Recordset_Agregar("Detalles_Entradas_Equipos")
            With Rs_Agrega_Detalles_Entradas_Equipos
                .AddNew
                    .rdoColumns("No_Operacion") = Conectar_Ayudante.Maximo_Catalogo("Detalles_Entradas_Equipos", "No_Operacion")
                    .rdoColumns("Empleado_ID") = Empleado_ID_Registro
                    .rdoColumns("No_Tarjeta") = Val(EnrollNumber)
                    .rdoColumns("No_Maquina") = Maquina
                    .rdoColumns("Fecha") = Format(Grid_Registro_Checadas.TextMatrix(Grid_Registro_Checadas.Rows - 1, 2), "MM/dd/yyyy")
                    .rdoColumns("Hora") = Format(Grid_Registro_Checadas.TextMatrix(Grid_Registro_Checadas.Rows - 1, 3), "HH:mm:ss")
                    .rdoColumns("E_S") = Trim(Grid_Registro_Checadas.TextMatrix(Grid_Registro_Checadas.Rows - 1, 5))
                    .rdoColumns("Metodo_Verificacion") = Trim(Grid_Registro_Checadas.TextMatrix(Grid_Registro_Checadas.Rows - 1, 4))
                    .rdoColumns("Usuario_Creo") = Nombre_Usuario
                    .rdoColumns("Fecha_Creo") = Now
                .Update
                .Close
            End With
        Set Rs_Agrega_Detalles_Entradas_Equipos = Nothing
    Conexion_Base.CommitTrans
    'Da formato al grid
    If Grid_Registro_Checadas.Rows > 1 Then Grid_Registro_Checadas.FixedRows = 1
    Grid_Registro_Checadas.ColWidth(0) = 500    'Usuario
    Grid_Registro_Checadas.ColWidth(1) = 3500   'Nombre Usuario
    Grid_Registro_Checadas.ColWidth(2) = 1200   'Fecha
    Grid_Registro_Checadas.ColWidth(3) = 800    'Hora
    Grid_Registro_Checadas.ColWidth(4) = 1200   'Metodo Verificacion
    Grid_Registro_Checadas.ColWidth(5) = 600    'E/S
Exit Sub
HANDLER:
Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Llena_Informacion_Dispositivo
    'DESCRIPCIÓN:           Agrega la informacion general del dispositivo en las cajas de texto
    'PARÁMETROS:
    'CREO:
    'FECHA_CREO:
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Llena_Informacion_Dispositivo()
Dim Rs_Consulta_Informacion_Dispositivo As rdoResultset     'Informacion dek dispositivo

On Error GoTo HANDLER
Me.MousePointer = 11
'Consulta la informacion para conectarse
Mi_SQL = "SELECT Direccion_IP, Puerto_IP, No_Equipo FROM Cat_Equipos_Identificadores"
Set Rs_Consulta_Informacion_Dispositivo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
With Rs_Consulta_Informacion_Dispositivo
    If Not .EOF Then
        IP = .rdoColumns("Direccion_IP")
        Puerto = .rdoColumns("Puerto_IP")
        Maquina = .rdoColumns("No_Equipo")
    Else
        MsgBox "No hay informacion para el dispositivo seleccionado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
End With
Set Rs_Consulta_Informacion_Dispositivo = Nothing

Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation + vbOKOnly, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CZKEM1.Disconnect
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Consulta_Entradas_Registradas
    'DESCRIPCIÓN:          Consulta las entradas ya registradas
    'PARÁMETROS:
    'CREO:                  Yañez Rodriguez Diego Neftali
    'FECHA_CREO:            11 Diciembre 2009
    'MODIFICO:
    'FECHA_MODIFICO
    'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub Consulta_Entradas_Registradas()
Dim Rs_Consulta_Detalles_Entradas_Equipos As rdoResultset       'Informacion de las checadas
'Mi_SQL = "SELECT  CE.Empleado_ID, (CE.Apellido_Paterno+' '+CE.Apellido_Materno+' '+CE.Nombre) as Nombre, "
'Mi_SQL = Mi_SQL & " DEE.*"
'Mi_SQL = Mi_SQL & " FROM Cat_Empleados CE, Detalles_Entradas_Equipos DEE"
'Mi_SQL = Mi_SQL & " WHERE Fecha = " & Par_Fecha & Format(Now, "MM/dd/yyyy") & Par_Fecha
'
'Set Rs_Consulta_Detalles_Entradas_Equipos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'With Rs_Consulta_Detalles_Entradas_Equipos
'    If Not .EOF Then
'        If Grid_Registro_Checadas.Rows = 0 Then
'            Grid_Registro_Checadas.AddItem "Usuario" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "Hora" & Chr(9) & "Tipo Acceso" & Chr(9) & "E/S"
'        End If
'        While Not .EOF
'            Grid_Registro_Checadas.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & _
'                    Format(.rdoColumns("Fecha"), "MM/dd/yyyy") & Chr(9) & Format(.rdoColumns("Hora"), "HH:mm:ss") & Chr(9) & _
'                    .rdoColumns("Metodo_Verificacion") & Chr(9) & .rdoColumns("E_S")
'            .MoveNext
'        Wend
'        .Close
'        'Da formato al grid
'        If Grid_Registro_Checadas.Rows > 1 Then Grid_Registro_Checadas.FixedRows = 1
'        Grid_Registro_Checadas.ColWidth(0) = 500    'Usuario
'        Grid_Registro_Checadas.ColWidth(1) = 3500   'Nombre Usuario
'        Grid_Registro_Checadas.ColWidth(2) = 1200   'Fecha
'        Grid_Registro_Checadas.ColWidth(3) = 800    'Hora
'        Grid_Registro_Checadas.ColWidth(4) = 1200   'Metodo Verificacion
'        Grid_Registro_Checadas.ColWidth(5) = 600    'E/S
'    End If
'End With
Set Rs_Consulta_Detalles_Entradas_Equipos = Nothing
End Sub
