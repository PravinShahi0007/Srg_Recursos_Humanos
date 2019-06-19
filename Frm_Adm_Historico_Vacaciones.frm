VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Adm_Historico_Vacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico de Vacaciones"
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
   ScaleWidth      =   7590
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
      Begin VB.TextBox Txt_No_Tarjeta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Txt_Nombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox Txt_Empleado_ID 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Fra_Adm_Control_Calzado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid Grid_Adm_Historico_Vacaciones 
            Height          =   1920
            Left            =   75
            TabIndex        =   5
            Top             =   240
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   3387
            _Version        =   393216
            Rows            =   0
            Cols            =   100
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
         TabIndex        =   3
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "C"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Responsiva 
         Caption         =   "Imprimir Responsiva"
         Height          =   555
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "A"
         Top             =   5160
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker Dt_Adm_Control_Calzado_Fecha 
         Height          =   315
         Left            =   5400
         TabIndex        =   10
         Top             =   1440
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Format          =   110821377
         CurrentDate     =   42373
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha de entrega de calzado"
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "No. Tarjeta"
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Empleado ID"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Lbl_Hostorico_Vacaciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "HISTORICO VACACIONES"
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
         Left            =   1350
         TabIndex        =   6
         Top             =   15
         Width           =   4725
      End
   End
End
Attribute VB_Name = "Frm_Adm_Historico_Vacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Historico_Vacaciones
    'DESCRIPCIÓN:           Consulta los empleados y su histórico de vacaciones
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Historico_Vacaciones()
Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
    Grid_Adm_Historico_Vacaciones.Rows = 0
Dim Total_Años As Integer
Total_Años = Calculos_Años
Dim Columnas As String
    Dim Dia_Actual  As String
    Dia_Actual = Now
    'Consulta los datos generales del usuario
    Mi_SQL = "select Empleado_ID, No_Tarjeta, Cat_Empleados.Nombre+' '+ Apellido_Paterno+' '+ Apellido_Materno as Nomre_del_empleado, "
    Mi_SQL = Mi_SQL & " Cat_Puestos.Nombre as Puesto, Cat_Departamentos.Nombre as Departamento, '' as grupo, (Fecha_Ingreso+(365.5*2)) as Fecha_1,  Fecha_Ingreso, "
    Mi_SQL = Mi_SQL & " Cat_Empleados.Tipo_Empleado, DATEDIFF(DAY, Fecha_Ingreso, '" & Format(Dia_Actual, "mm/dd/yyyy") & "') as Dias_Trabajados "
    Mi_SQL = Mi_SQL & " From Cat_Empleados, Cat_Puestos, Cat_Departamentos "
    Mi_SQL = Mi_SQL & " Where Cat_Empleados.Puesto_ID = Cat_Puestos.Puesto_ID "
    Mi_SQL = Mi_SQL & " and Cat_Departamentos.Departamento_ID = Cat_Empleados.Departamento_ID "
'    Mi_SQL = Mi_SQL & "  and No_Tarjeta = 40 "
    Mi_SQL = Mi_SQL & " order BY No_Tarjeta"
    
    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Empleados
        If Not .EOF Then
            Columnas = "No Tarjeta" & Chr(9) & "Nombre" & Chr(9) & "Puesto" & Chr(9) & "Departamento" & Chr(9) & "Grupo" & Chr(9) & "Fecha" & Chr(9) & "Ingreso" & Chr(9) & "Tipo_Empleado" & Chr(9) & "Dias Trabajados" & Chr(9) & "Empleado_ID"
            If Total_Años > 0 Then
                For I = 1 To Total_Años
                Columnas = Columnas & Chr(9) & "Días año " & I & Chr(9) & " Restante"
                Next I
            End If
            Grid_Adm_Historico_Vacaciones.AddItem Columnas
            While Not .EOF
            Dim Numero_Días As Double
            Dim Dias_Restantes As Double
                Dim Cadena_Datos As String
                Cadena_Datos = .rdoColumns("No_Tarjeta") & Chr(9) & .rdoColumns("Nomre_del_empleado") & Chr(9) & .rdoColumns("Puesto") & Chr(9) & .rdoColumns("Departamento") & Chr(9) & .rdoColumns("Grupo") & Chr(9) & .rdoColumns("Fecha_1") & Chr(9) & .rdoColumns("Fecha_Ingreso") & Chr(9) & .rdoColumns("Tipo_Empleado") & Chr(9) & .rdoColumns("Dias_Trabajados") & Chr(9) & .rdoColumns("Empleado_ID")
                For I = 1 To Total_Años
                'Calcular_Numero_Dias
                Dim Ye As Integer
                Ye = Obtener_Valor_Ye(I)
                Dim Dias_Trabajados As Integer
'                Dias_Trabajados = Val(.rdoColumns("Dias_Trabajados"))
                Numero_Días = Calcular_Numero_Dias(Val(.rdoColumns("Dias_Trabajados")), Ye)
                'Calcular_Restantes
                If I = 1 Then
                Dias_Restantes = Val(.rdoColumns("Dias_Trabajados")) - 365
                Else
                Dias_Restantes = Dias_Restantes - 365
                End If
                
                Cadena_Datos = Cadena_Datos & Chr(9) & Numero_Días & Chr(9) & Dias_Restantes
                Next I
                Grid_Adm_Historico_Vacaciones.AddItem Cadena_Datos
                
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Adm_Historico_Vacaciones.FixedRows = 1
            Grid_Adm_Historico_Vacaciones.ColWidth(0) = 800     'No_Tarjeta
            Grid_Adm_Historico_Vacaciones.ColWidth(1) = 1500   'Nombre
            Grid_Adm_Historico_Vacaciones.ColWidth(2) = 800   'Puesto
            Grid_Adm_Historico_Vacaciones.ColWidth(3) = 800   'Departamento
            Grid_Adm_Historico_Vacaciones.ColWidth(4) = 500  'Grupo
            Grid_Adm_Historico_Vacaciones.ColWidth(5) = 800  'Fecha_1
            Grid_Adm_Historico_Vacaciones.ColWidth(6) = 800  'Fecha_Ingreso
            Grid_Adm_Historico_Vacaciones.ColWidth(7) = 800  'Tipo_Empleado
            Grid_Adm_Historico_Vacaciones.ColWidth(8) = 800  'Días_Trabajados
            Grid_Adm_Historico_Vacaciones.ColWidth(9) = 0  'Empleado_ID
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Adm_Empleados = Nothing
    

End Sub

Private Sub Btn_Buscar_Click()
Consulta_Historico_Vacaciones
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Calculos_Años
    'DESCRIPCIÓN:           Reimprime el grid con de acuerdo a las fechas
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Function Calculos_Años() As Integer
Dim Rs_Consulta_Adm_Empleados As rdoResultset       'Informacion de los registros
    Dim Dia_Actual  As String
    Dia_Actual = Now
    Dim Total_Años  As Integer
    Calculos_Años = 0
    'Consulta los datos generales del usuario
    Mi_SQL = "select TOP 1 Fecha_Ingreso, DATEDIFF (YEAR,  Fecha_Ingreso, '" & Format(Dia_Actual, "dd/mm/yyyy") & "') as Total_Años from Cat_Empleados"
    Mi_SQL = Mi_SQL & " ORDER BY Fecha_Ingreso ASC "
    Set Rs_Consulta_Adm_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Adm_Empleados
        If Not .EOF Then
            Calculos_Años = .rdoColumns("Total_Años")
        End If
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Obtener_Valor_Ye
    'DESCRIPCIÓN:           Obtiene el valor de la BD de acuerdo al año
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Function Obtener_Valor_Ye(ByVal año As Integer) As Integer
Dim Rs_Consulta_Referencia As rdoResultset       'Informacion de los registros
    'Consulta los datos generales del usuario
    Mi_SQL = "select Top 1 Referencia_Id, Año, Valor from Ope_Referencias_Reporte_Vacaciones"
    Mi_SQL = Mi_SQL & " where Año = " & año
    Set Rs_Consulta_Referencia = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    With Rs_Consulta_Referencia
        If Not .EOF Then
            Obtener_Valor_Ye = .rdoColumns("Valor")
        End If
    End With
End Function

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Calcular_Numero_Dias
    'DESCRIPCIÓN:           Obtiene el valor de la BD de acuerdo al año
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           04 Marzo 2016
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************

Private Function Calcular_Numero_Dias(ByVal Dias_Trabajados As Double, ByVal Ye As Integer) As Integer
If Dias_Trabajados - 365 > 0 Then
    Calcular_Numero_Dias = 10
Else
    If (Dias_Trabajados * ((1 * Ye) / 365)) < 0 Then
    Calcular_Numero_Dias = 0
    Else
    Calcular_Numero_Dias = Dias_Trabajados * ((1 * Ye) / 365)
    End If
End If
End Function
