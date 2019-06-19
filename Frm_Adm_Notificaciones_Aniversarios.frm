VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Adm_Notificaciones_Aniversarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificación de Vacaciones"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleMode       =   0  'User
   ScaleWidth      =   7560
   Begin VB.PictureBox Pic_Ope_Programacion_Cursos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.CommandButton Btn_Actualizar_Dias 
         Caption         =   "Actualizar Días"
         Height          =   555
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "A"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Ope_Programacion_Invitacion_Empleados 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleados"
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
         Height          =   4080
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Notificaciones_Aniversarios 
            Height          =   3360
            Left            =   75
            TabIndex        =   4
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5927
            _Version        =   393216
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "C"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Label Lbl_Notificaciones_Aniversarios 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOTIFICACIÓN DE ANIVERSARIOS"
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
         Left            =   540
         TabIndex        =   5
         Top             =   15
         Width           =   6345
      End
   End
End
Attribute VB_Name = "Frm_Adm_Notificaciones_Aniversarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicializa()

End Sub

Private Sub Btn_Actualizar_Dias_Click()
If Btn_Actualizar_Dias.Caption = "Actualizar Días" Then
Btn_Buscar.Caption = "Buscar"
Consultar_Vacaciones
Else
For I = 1 To Grid_Adm_Notificaciones_Aniversarios.Rows - 1
            If Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 4) <> "Empleado_Id" Then
                Call Modificar_Vacaciones_Empleados(Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 4), Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 2), Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 3))
            End If
        Next I
        
 Consultar_Vacaciones
End If
End Sub

Private Sub Btn_Buscar_Click()
If Btn_Buscar.Caption = "Buscar" Then
    Grid_Adm_Notificaciones_Aniversarios.Rows = 0
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros

'    Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Empleado_ID, No_Tarjeta, (ISNULL(Nombre, '')+ ' ' + ISNULL (Apellido_Paterno, '') +' '+ ISNULL(Apellido_Materno, '')) as Nombre, Fecha_Ingreso, ((YEAR(GETDATE()))-YEAR(Fecha_Ingreso)) as Anios  "
    Mi_SQL = Mi_SQL & ", Salario_Diario_Variable AS Vacaciones FROM Cat_Empleados Where (Day(Fecha_Ingreso) = " & Day(Now)
    Mi_SQL = Mi_SQL & " AND MONTH(Fecha_Ingreso) = " & Month(Now)
    Mi_SQL = Mi_SQL & " AND Estatus = 'A') "

'    MsgBox Mi_SQL
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Empleados
        If Not .EOF Then
        Btn_Buscar.Caption = "Enviar_Correos"
            If Grid_Adm_Notificaciones_Aniversarios.Rows <= 0 Then
            Grid_Adm_Notificaciones_Aniversarios.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Años" & Chr(9) & "Vacaciones" & Chr(9) & "Empleado_ID"
           End If
           While Not .EOF
                Grid_Adm_Notificaciones_Aniversarios.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Anios") & Chr(9) & .rdoColumns("Vacaciones") & Chr(9) & .rdoColumns("Empleado_ID")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Adm_Notificaciones_Aniversarios.FixedRows = 1
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(0) = 1000     'No_Tarjeta
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(1) = 4200   'Nombre
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(2) = 800   'Años
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(3) = 1000   'Vacaciones
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(4) = 0
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Empleados = Nothing
Else
Me.MousePointer = 11
Enviar_Correos
MsgBox ("Se enviaron los correos a los empleados con el dato")
Me.MousePointer = 0
End If
End Sub

'Private Sub Btn_Enviar_Correos_Click()
'Enviar_Correos
'
'End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Enviar_Correos
    'DESCRIPCIÓN:           Consulta los registos del grid para enviar el correo a cada empleado
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           06 Enero 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Enviar_Correos()
  Frm_Email.Obtener_Parametros_Correos
  
    For I = 1 To Grid_Adm_Notificaciones_Aniversarios.Rows - 1
        Dim Usuariooo As String
        Usuariooo = Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 4)
        Dim Dias As Integer
        Dias = Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 3)
        Dim Rs_Consulta_Ope_Email As rdoResultset       'Informacion de los registros
        'Consulta los datos generales del usuario
        Mi_SQL = "SELECT * "
        Mi_SQL = Mi_SQL & " FROM Cat_Empleados"
        Mi_SQL = Mi_SQL & " WHERE Empleado_ID = " & Usuariooo
        Set Rs_Consulta_Ope_Email = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Ope_Email
        If Not .EOF Then
            While Not .EOF
             If Not IsNull(.rdoColumns("Email")) Then
                Frm_Email.Correo = .rdoColumns("Email")
                End If
                Frm_Email.Asunto = "Notificación de aniversario"
                Frm_Email.Mensaje = "Nos complace informarle que el el día de hoy (" & Format(Day(Now), dd) & "-" & Format(Month(Now), mmmm) & "-" & Format(Year(Now), yyyy) & ") " & _
                "cumple " & Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 2) & " años laborando con nosotros " & vbLf & _
                "Le informamos también que podrá disponer de " & Val(Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 3)) & _
                " días de vacaciones, de los cuales " & Val(Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 3)) - Calcular_Días(Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 2), .rdoColumns("Tipo_Empleado")) & _
                " son acumulados anteriormente y " & Calcular_Días(Grid_Adm_Notificaciones_Aniversarios.TextMatrix(I, 2), .rdoColumns("Tipo_Empleado")) & " días consecuentes del aniversario de su fecha de ingreso"
                
                
               Dim enviar As Boolean
               If Not IsNull(.rdoColumns("Email")) And .rdoColumns("Email") <> "" Then
               enviar = Frm_Email.Mandar_Correo
               End If
                Unload Frm_Email
                .MoveNext
            Wend
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Ope_Email = Nothing

Next I

End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consultar_Vacaciones
    'DESCRIPCIÓN:           Consulta los empleados que cumplen vacaciones y no han sido actualizados en cuannto a días
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           03 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consultar_Vacaciones()
    Grid_Adm_Notificaciones_Aniversarios.Rows = 0
Dim Rs_Consulta_Cat_Empleados As rdoResultset       'Informacion de los registros

'    Grid_Ope_Programacion_Invitacion_Empleados.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Empleado_ID, No_Tarjeta, (ISNULL(Nombre, '')+ ' ' + ISNULL (Apellido_Paterno, '') +' '+ ISNULL(Apellido_Materno, '')) as Nombre, Fecha_Ingreso, (" & Year(Now) & "-YEAR(Fecha_Ingreso)) as Anios  "
    Mi_SQL = Mi_SQL & ", Salario_Diario_Variable AS Vacaciones FROM Cat_Empleados Where Day(Fecha_Ingreso) = Day('" & Format(Now, "mm/dd/yyyy") & "') "
    Mi_SQL = Mi_SQL & " AND MONTH(Fecha_Ingreso) = MONTH('" & Format(Now, "mm/dd/yyyy") & "') "
    Mi_SQL = Mi_SQL & "AND (Fecha_Vacaciones_Actualizacion IS NULL or ( "
    Mi_SQL = Mi_SQL & " Year (Fecha_Vacaciones_Actualizacion) <> " & Year(Now) & " "
    Mi_SQL = Mi_SQL & " AND MONTH(Fecha_Vacaciones_Actualizacion) <> " & Month(Now) & " "
    Mi_SQL = Mi_SQL & " AND DAY(Fecha_Vacaciones_Actualizacion) <> " & Day(Now) & "))"
   Mi_SQL = Mi_SQL & " AND Estatus= 'A'  "
'    MsgBox Mi_SQL
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Cat_Empleados
        If Not .EOF Then
            Btn_Actualizar_Dias.Caption = "Actualizar"
            If Grid_Adm_Notificaciones_Aniversarios.Rows <= 0 Then
            Grid_Adm_Notificaciones_Aniversarios.AddItem "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Años" & Chr(9) & "Vacaciones" & Chr(9) & "Empleado_ID"
           End If
           While Not .EOF
                Grid_Adm_Notificaciones_Aniversarios.AddItem .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Anios") & Chr(9) & .rdoColumns("Vacaciones") & Chr(9) & .rdoColumns("Empleado_ID")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Adm_Notificaciones_Aniversarios.FixedRows = 1
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(0) = 1000     'No_Tarjeta
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(1) = 4200   'Nombre
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(2) = 800   'Años
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(3) = 1000   'Vacaciones
            Grid_Adm_Notificaciones_Aniversarios.ColWidth(4) = 0
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Empleados = Nothing
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Vacaciones_Empleados
    'DESCRIPCIÓN:           Modifica el registro de los días de vacaciones
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    03 Marzo 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Vacaciones_Empleados(ByVal Empleado_ID As String, ByVal Años As String, ByVal Dias As String)
Dim Rs_Modificacion_Cat_Empleados As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
Dim Total_Dias As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Empleados"
    Mi_SQL = Mi_SQL & " WHERE Empleado_ID ='" & Trim(Empleado_ID) & "'"
    Set Rs_Modificacion_Cat_Empleados = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Empleados
        .Edit
'        Dim Días_Int As Integer
'        Dim Dias_Calculo As Integer
'        Dias_Calculo = Calcular_Días(Años)
'        Días_Int = Val(Dias)
'        Total_Dias = Días_Int + Dias_Calculo
    Dim Tipo_Empleado_Aux As String
        Tipo_Empleado_Aux = .rdoColumns("Tipo_Empleado")
        If Trim(UCase(.rdoColumns("Tipo_Empleado"))) <> "CONFIANZA" And Trim(UCase(.rdoColumns("Tipo_Empleado"))) <> "SINDICALIZADO" Then
            Tipo_Empleado_Aux = "otros"
        End If
            .rdoColumns("Salario_Diario_Variable") = Val(Dias) + Calcular_Días(Años, Trim(Tipo_Empleado_Aux))
            .rdoColumns("Fecha_Vacaciones_Actualizacion") = Now
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Empleados = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
'   MsgBox "La Institución ha sido modificada", vbInformation + vbOKOnly, Me.Caption
'   Consultar_Vacaciones
   Btn_Actualizar_Dias.Caption = "Actualizar Días"
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Vacaciones_Empleados
    'DESCRIPCIÓN:           Modifica el registro de los días de vacaciones
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    03 Marzo 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Function Calcular_Días(ByVal Años As String, ByVal Tipo_Empleado As String) As Integer

Dim Rs_Consulta_Dias_Vacaciones As rdoResultset       'Informacion de los registros
     Dim Diaasss As String
    Mi_SQL = "select * from Ope_Referencias_Reporte_Vacaciones "
    Set Rs_Consulta_Dias_Vacaciones = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Dias_Vacaciones
        While Not .EOF
        If Años > 10 And (Val(Años) Mod 5 = 0) Then
            If Trim(.rdoColumns("Tipo_Empleado")) = Trim(Tipo_Empleado) Then
'                Dim Diaasss As String
                Diaasss = .rdoColumns("Añosmas10")
                Calcular_Días = Diaasss
            End If
        Else
            If Trim(.rdoColumns("Tipo_Empleado")) = Trim(Tipo_Empleado) Then
               
                Diaasss = .rdoColumns("Año_" & Años)
                Calcular_Días = Diaasss
            End If
        End If
        .MoveNext
        Wend
            .Close
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Dias_Vacaciones = Nothing

'Dim Años_Actuales As Integer
'Años_Actuales = Val(Años)
'Calcular_Días = 0
'If Años > 10 And (Val(Años) Mod 5 = 0) Then
'    Calcular_Días = 2
'Else
'    If (Tipo_Empleado = "confianza") Then
'        Select Case Años_Actuales
'            Case 1, 2
'                Calcular_Días = 10
'            Case 3
'                Calcular_Días = 12
'            Case 4
'                Calcular_Días = 14
'            Case 5, 6, 7, 8, 9
'                Calcular_Días = 16
'            Case 10
'                Calcular_Días = 18
'        End Select
'    Else
'    Select Case Años_Actuales
'            Case 1, 2
'                Calcular_Días = 8
'            Case 3
'                Calcular_Días = 10
'            Case 4
'                Calcular_Días = 12
'            Case 5, 6, 7, 8, 9
'                Calcular_Días = 14
'            Case 10
'                Calcular_Días = 16
'        End Select
'    End If
'
'End If
End Function

Private Sub Btn_Salir_Click()
Unload Me
End Sub
