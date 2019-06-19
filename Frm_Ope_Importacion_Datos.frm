VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Ope_Importacion_Datos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACTUALIZACION"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Cmd_Archivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_Actualizar_Productos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7005
      Left            =   -15
      ScaleHeight     =   7005
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   9525
      Begin VB.CommandButton Btn_Aplicar 
         Caption         =   "Aplicar"
         Height          =   500
         Left            =   150
         Picture         =   "Frm_Ope_Importacion_Datos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "C"
         Top             =   5715
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   500
         Left            =   7965
         Picture         =   "Frm_Ope_Importacion_Datos.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5715
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Fra_Registros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lista de Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         Left            =   105
         TabIndex        =   9
         Top             =   1455
         Width           =   9360
         Begin MSComctlLib.ProgressBar Pbar_Registros 
            Height          =   1155
            Left            =   915
            TabIndex        =   14
            Top             =   1350
            Visible         =   0   'False
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   2037
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Registros 
            Height          =   3795
            Left            =   75
            TabIndex        =   10
            Top             =   225
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6694
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Fra_Archivos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   105
         TabIndex        =   1
         Top             =   405
         Width           =   9360
         Begin VB.ComboBox Cmb_Tipo_Producto 
            Height          =   315
            ItemData        =   "Frm_Ope_Importacion_Datos.frx":0634
            Left            =   105
            List            =   "Frm_Ope_Importacion_Datos.frx":063E
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox Txt_Archivo 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   7065
         End
         Begin VB.TextBox Txt_Ruta 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   210
            Width           =   7065
         End
         Begin VB.CommandButton Btn_Explorar 
            Caption         =   "Explorar"
            Height          =   330
            Left            =   90
            TabIndex        =   3
            Top             =   585
            Width           =   1185
         End
         Begin VB.TextBox Txt_Tipo_Salida_Sucursal 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1110
            Visible         =   0   'False
            Width           =   7065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo"
            Height          =   195
            Left            =   1365
            TabIndex        =   8
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Lbl_Origen 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Origen"
            Height          =   195
            Left            =   1365
            TabIndex        =   7
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Lbl_Tipo_Salida_Sucursal 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1170
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Label Lbl_Actualizacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ACTUALIZACION DE EMPLEADOS"
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
         Left            =   2085
         TabIndex        =   13
         Top             =   30
         Width           =   6120
      End
   End
   Begin VB.PictureBox Pic_Actualizacion_Datos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6390
      Left            =   -15
      ScaleHeight     =   6390
      ScaleWidth      =   9705
      TabIndex        =   15
      Top             =   -15
      Visible         =   0   'False
      Width           =   9705
      Begin VB.CommandButton Btn_Costo_Promedio 
         Caption         =   "Costo Promedio"
         Height          =   465
         Left            =   4185
         TabIndex        =   23
         Top             =   5790
         Width           =   1185
      End
      Begin VB.CommandButton Btn_Capturar 
         Caption         =   "Importar"
         Height          =   465
         Left            =   120
         TabIndex        =   21
         Top             =   5790
         Width           =   1185
      End
      Begin VB.Frame Fra_Proveedores 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   75
         TabIndex        =   19
         Top             =   3075
         Width           =   9435
         Begin VB.Data Data_Proveedores 
            Caption         =   "Proveedores"
            Connect         =   "Excel 8.0;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   105
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   0  'Table
            RecordSource    =   "Proveedores$"
            Top             =   2220
            Width           =   9240
         End
         Begin MSDBGrid.DBGrid Dbg_Proveedores 
            Bindings        =   "Frm_Ope_Importacion_Datos.frx":064E
            Height          =   1980
            Left            =   120
            OleObjectBlob   =   "Frm_Ope_Importacion_Datos.frx":066D
            TabIndex        =   20
            Top             =   225
            Width           =   9195
         End
      End
      Begin VB.Frame Fra_Clientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         Left            =   75
         TabIndex        =   17
         Top             =   450
         Width           =   9435
         Begin VB.Data Data_Clientes 
            Caption         =   "Clientes"
            Connect         =   "Excel 8.0;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   90
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   0  'Table
            RecordSource    =   "Hoja1$"
            Top             =   2175
            Width           =   9255
         End
         Begin MSDBGrid.DBGrid Dbg_Clientes 
            Bindings        =   "Frm_Ope_Importacion_Datos.frx":1048
            Height          =   1950
            Left            =   105
            OleObjectBlob   =   "Frm_Ope_Importacion_Datos.frx":1064
            TabIndex        =   18
            Top             =   210
            Width           =   9210
         End
      End
      Begin VB.CommandButton Btn_Salir_Importacion 
         Caption         =   "Salir"
         Height          =   465
         Left            =   8250
         TabIndex        =   16
         Top             =   5790
         Width           =   1185
      End
      Begin VB.Label Lbl_Importacion_Datos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Importación de Datos"
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
         Left            =   2745
         TabIndex        =   22
         Top             =   15
         Width           =   3750
      End
   End
End
Attribute VB_Name = "Frm_Ope_Importacion_Datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'NOMBRE_FUNCION: Aplicar_Turnos
'DESCRIPCION: Actualiza los registros de empleados mediante el layout establecido
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 06-Febrero-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Aplicar_Turnos()
Dim Mi_SQL As String
Dim Rs_Empleados As rdoResultset
Dim Rs_Auxiliar As rdoResultset
Dim Departamento_ID As String
Dim Turno_ID As String
Dim Gap_ID As String
Dim Supervisor_ID As String
Dim Puesto_ID As String
Dim Seccion_ID As String

On Error GoTo HANDLER
    If Grid_Registros.Rows < 2 Then Exit Sub
    If MsgBox("¿Esta seguro de actualizar los turnos?", vbQuestion + vbYesNo) = vbYes Then
        Me.MousePointer = 11
        Pbar_Registros.Value = 0
        Pbar_Registros.Max = Grid_Registros.Rows - 1
        Pbar_Registros.Visible = True
        Conexion_Base.BeginTrans
        'Articulos
        For Fila = 1 To Grid_Registros.Rows - 1
            'Consulta si existe el producto
            Mi_SQL = "SELECT * FROM Cat_Empleados"
            Mi_SQL = Mi_SQL & " WHERE No_Tarjeta='" & Trim(Grid_Registros.TextMatrix(Fila, 0)) & "'"
            Set Rs_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Empleados.EOF Then
                'Busca el ID del departamento
                Mi_SQL = "SELECT * FROM Cat_Departamentos WHERE Clave='" & Trim(Grid_Registros.TextMatrix(Fila, 1)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Departamento_ID = Rs_Auxiliar.rdoColumns("Departamento_ID")
                Else
                    Departamento_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID del gap
                Mi_SQL = "SELECT * FROM Cat_Gaps WHERE Nombre='" & Trim(Grid_Registros.TextMatrix(Fila, 3)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Gap_ID = Rs_Auxiliar.rdoColumns("Gap_ID")
                Else
                    Gap_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID del supervisor
                Mi_SQL = "SELECT * FROM Cat_Empleados WHERE No_Tarjeta='" & Trim(Grid_Registros.TextMatrix(Fila, 4)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Supervisor_ID = Rs_Auxiliar.rdoColumns("Empleado_ID")
                Else
                    Supervisor_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID del puesto
                Mi_SQL = "SELECT * FROM Cat_Puestos WHERE Abreviatura='" & Trim(Grid_Registros.TextMatrix(Fila, 5)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Puesto_ID = Rs_Auxiliar.rdoColumns("Puesto_ID")
                Else
                    Puesto_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID de la seccion
                Mi_SQL = "SELECT * FROM Cat_Secciones WHERE Clave='" & Trim(Grid_Registros.TextMatrix(Fila, 6)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Seccion_ID = Rs_Auxiliar.rdoColumns("Clave")
                Else
                    Seccion_ID = ""
                End If
                Rs_Auxiliar.Close
                'Si hubo cambio de turno registra el movimiento en programacion
                If Trim(Rs_Empleados.rdoColumns("Turno_ID")) <> Format(Grid_Registros.TextMatrix(Fila, 2), "00000") Then
                    Mi_SQL = "INSERT INTO Adm_Cambios_Turnos (Empleado_ID,Turno_Anterior_ID,Turno_Nuevo_ID,Fecha_Cambio,Estatus,Usuario_Creo,Fecha_Creo)"
                    Mi_SQL = Mi_SQL & " VALUES('" & Rs_Empleados.rdoColumns("Empleado_ID") & "'"
                    Mi_SQL = Mi_SQL & " ,'" & Rs_Empleados.rdoColumns("Turno_ID") & "'"
                    Mi_SQL = Mi_SQL & " ,'" & Format(Grid_Registros.TextMatrix(Fila, 2), "00000") & "'"
                    Mi_SQL = Mi_SQL & " ,'" & Format(Now, "MM/dd/yyyy") & "'"
                    Mi_SQL = Mi_SQL & " ,'CAMBIADO'"
                    Mi_SQL = Mi_SQL & " ,'" & Nombre_Usuario & "'"
                    Mi_SQL = Mi_SQL & " ,GETDATE())"
                    Conexion_Base.Execute Mi_SQL
                End If
                'Actualiza el registro de empleado
                Mi_SQL = "UPDATE Cat_Empleados"
                Mi_SQL = Mi_SQL & " SET Turno_ID='" & Format(Grid_Registros.TextMatrix(Fila, 2), "00000") & "'"
                If Departamento_ID <> "" Then
                    Mi_SQL = Mi_SQL & " , Departamento_ID='" & Departamento_ID & "'"
                End If
                If Gap_ID <> "" Then
                    Mi_SQL = Mi_SQL & " , Gap_ID='" & Gap_ID & "'"
                End If
                If Supervisor_ID <> "" Then
                    Mi_SQL = Mi_SQL & " , Supervisor_ID='" & Supervisor_ID & "'"
                End If
                If Puesto_ID <> "" Then
                    Mi_SQL = Mi_SQL & " , Puesto_ID='" & Puesto_ID & "'"
                End If
                If Seccion_ID <> "" Then
                    Mi_SQL = Mi_SQL & " , Nomipaq_ID='" & Seccion_ID & "'"
                End If
                Mi_SQL = Mi_SQL & " WHERE Empleado_ID='" & Rs_Empleados.rdoColumns("Empleado_ID") & "'"
                Conexion_Base.Execute Mi_SQL
                Pbar_Registros.Value = Pbar_Registros.Value + 1
                Me.Refresh
            End If
            Rs_Empleados.Close
        Next Fila
        Conexion_Base.CommitTrans
        MsgBox "Importacion realizada con éxito", vbInformation
        Btn_Aplicar.Enabled = False
    End If
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Exit Sub
HANDLER:
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Aplicar_SAP_Altas
'DESCRIPCION: Da de alta los registros de empleados mediante el layout enviado por SAP
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 07-Febrero-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Aplicar_SAP_Altas()
Dim Mi_SQL As String
Dim Rs_Empleados As rdoResultset
Dim Rs_Alta_Empleados As rdoResultset
Dim Rs_Auxiliar As rdoResultset
Dim Seccion_ID As String
Dim Departamento_ID As String
Dim Puesto_ID As String

On Error GoTo HANDLER
    If Grid_Registros.Rows < 2 Then Exit Sub
    If MsgBox("¿Esta seguro de dar de alta los registros?", vbQuestion + vbYesNo) = vbYes Then
        Me.MousePointer = 11
        Pbar_Registros.Value = 0
        Pbar_Registros.Max = Grid_Registros.Rows - 1
        Pbar_Registros.Visible = True
        Conexion_Base.BeginTrans
        Set Rs_Alta_Empleados = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados")
        'Recorre el listado
        For Fila = 1 To Grid_Registros.Rows - 1
            'Valida no este dando de alta el mismo empleado
            Mi_SQL = "SELECT * FROM Cat_Empleados WHERE No_Tarjeta='" & Trim(Grid_Registros.TextMatrix(Fila, 0)) & "'"
            Set Rs_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Rs_Empleados.EOF Then
                'Busca el ID de la seccion
                Mi_SQL = "SELECT * FROM Cat_Secciones WHERE Clave='" & Trim(Grid_Registros.TextMatrix(Fila, 5)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Seccion_ID = Rs_Auxiliar.rdoColumns("Seccion_ID")
                Else
                    Seccion_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID del departamento
                Mi_SQL = "SELECT * FROM Cat_Departamentos WHERE Clave='" & Trim(Grid_Registros.TextMatrix(Fila, 6)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Departamento_ID = Rs_Auxiliar.rdoColumns("Departamento_ID")
                Else
                    Departamento_ID = ""
                End If
                Rs_Auxiliar.Close
                'Busca el ID del puesto
                Mi_SQL = "SELECT * FROM Cat_Puestos WHERE Abreviatura='" & Trim(Grid_Registros.TextMatrix(Fila, 7)) & "'"
                Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Auxiliar.EOF Then
                    Puesto_ID = Rs_Auxiliar.rdoColumns("Puesto_ID")
                Else
                    Puesto_ID = ""
                End If
                Rs_Auxiliar.Close
                'Da de alta el registro de empleados
                With Rs_Alta_Empleados
                    .AddNew
                        .rdoColumns("Empleado_ID") = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empleados", "Empleado_ID"), "00000")
                        .rdoColumns("Empresa_ID") = "00001"                 'SRG
                        If Departamento_ID <> "" Then
                            .rdoColumns("Departamento_ID") = Departamento_ID
                        End If
                        If Puesto_ID <> "" Then
                            .rdoColumns("Puesto_ID") = Puesto_ID
                        End If
                        .rdoColumns("Turno_ID") = Format(Trim(Grid_Registros.TextMatrix(Fila, 8)), "00000")
                        .rdoColumns("Motivo_Baja_ID") = "00000"
                        'Datos Personales
                        If Trim(Trim(Grid_Registros.TextMatrix(Fila, 1))) = "B" Then
                            .rdoColumns("Estatus") = "I"
                        Else
                            .rdoColumns("Estatus") = "A"
                        End If
'                        .rdoColumns("Estatus") = Trim(Grid_Registros.TextMatrix(Fila, 1))
'                        .rdoColumns("Tipo") = "E"                           'Empleado
                        .rdoColumns("Nombre") = Trim(Grid_Registros.TextMatrix(Fila, 2))
                        .rdoColumns("Apellido_Paterno") = Trim(Grid_Registros.TextMatrix(Fila, 3))
                        .rdoColumns("Apellido_Materno") = Trim(Grid_Registros.TextMatrix(Fila, 4))
                        .rdoColumns("Lugar_Nacimiento") = ""
                        .rdoColumns("Sexo") = "M"
                        .rdoColumns("Fecha_Nacimiento") = Format(Now, "MM/dd/yyyy")
                        .rdoColumns("Clave_Elector") = ""
                        .rdoColumns("Estado_Civil") = ""
                        .rdoColumns("RFC") = ""
                        .rdoColumns("Curp") = ""
                        .rdoColumns("Nss") = ""
                        .rdoColumns("Direccion") = ""
                        .rdoColumns("Colonia") = ""
                        .rdoColumns("Codigo_Postal") = ""
                        .rdoColumns("Ciudad") = ""
                        .rdoColumns("Estado") = ""
                        .rdoColumns("Imagen_Perfil") = Trim(Grid_Registros.TextMatrix(Fila, 0)) & ".JPG"
                        .rdoColumns("No_Tarjeta") = Trim(Grid_Registros.TextMatrix(Fila, 0))
                        .rdoColumns("Clave_SAP") = "SI" & Trim(Grid_Registros.TextMatrix(Fila, 0))
                        .rdoColumns("Fecha_Ingreso") = Format(Now, "MM/dd/yyyy")
                        .rdoColumns("Tipo_Empleado") = ""
                        .rdoColumns("Tipo_Contratacion") = "PLANTA"
                        .rdoColumns("Fecha_Termino_Contrato") = Format(Now, "MM/dd/yyyy")
                        .rdoColumns("Salario_Diario") = 0
                        .rdoColumns("Salario_Diario_Variable") = 0
                        .rdoColumns("Cedula_Identidad_Ciudadana") = ""
                        .rdoColumns("Trabaja_Domingos") = "N"
                        .rdoColumns("Infonavit") = "N"
                        .rdoColumns("Retardos") = 0
                        .rdoColumns("Fecha_Retardo") = "01/01/1960"
                        .rdoColumns("En_Caso_Emergencia") = ""
                        .rdoColumns("Telefono_Emergencia1") = ""
                        .rdoColumns("Telefono_Emergencia2") = ""
                        .rdoColumns("Alergia1") = ""
                        .rdoColumns("Alergia2") = ""
                        .rdoColumns("Alergia3") = ""
                        .rdoColumns("Fecha_Baja") = Format(Now, "MM/dd/yyyy")
                        .rdoColumns("Comentarios_Baja") = ""
                        .rdoColumns("Usuario_Creo") = Nombre_Usuario
                        .rdoColumns("Fecha_Creo") = Now
                    .Update
                End With
            End If
            Rs_Empleados.Close
            Pbar_Registros.Value = Pbar_Registros.Value + 1
            Me.Refresh
        Next Fila
        Rs_Alta_Empleados.Close
        Conexion_Base.CommitTrans
        MsgBox "Importacion realizada con éxito", vbInformation
        Btn_Aplicar.Enabled = False
    End If
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Exit Sub
HANDLER:
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Aplicar_SAP_Bajas
'DESCRIPCION: Da de baja de los registros de empleados mediante el layout enviado por SAP
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 11-Febrero-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Aplicar_SAP_Bajas()
Dim Mi_SQL As String
Dim Rs_Auxiliar As rdoResultset
Dim Baja_ID As String

On Error GoTo HANDLER
    If Grid_Registros.Rows < 2 Then Exit Sub
    If MsgBox("¿Esta seguro de actualizar los registros?", vbQuestion + vbYesNo) = vbYes Then
        Me.MousePointer = 11
        Pbar_Registros.Value = 0
        Pbar_Registros.Max = Grid_Registros.Rows - 1
        Pbar_Registros.Visible = True
        Conexion_Base.BeginTrans
        'Recorre el listado
        For Fila = 1 To Grid_Registros.Rows - 1
            'Busca el ID del motivo de la baja
            Mi_SQL = "SELECT * FROM Cat_Motivos_Baja WHERE Clave_SAP='" & Trim(Grid_Registros.TextMatrix(Fila, 5)) & "'"
            Set Rs_Auxiliar = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Auxiliar.EOF Then
                Baja_ID = Rs_Auxiliar.rdoColumns("Motivo_Baja_ID")
            Else
                Baja_ID = ""
            End If
            Rs_Auxiliar.Close
            'Actualiza el registro de empleado
            Mi_SQL = "UPDATE Cat_Empleados"
            Mi_SQL = Mi_SQL & " SET Estatus='I'"
            Mi_SQL = Mi_SQL & " ,Fecha_Baja='" & Format(Grid_Registros.TextMatrix(Fila, 6), "MM/dd/yyyy") & "'"
            If Baja_ID <> "" Then
                Mi_SQL = Mi_SQL & " ,Motivo_Baja_ID" = Baja_ID
            End If
            Mi_SQL = Mi_SQL & " ,Comentarios_Baja='BAJA DESDE ARCHIVO SAP'"
            Mi_SQL = Mi_SQL & " ,Usuario_Modifico='" & Nombre_Usuario & "'"
            Mi_SQL = Mi_SQL & " ,Fecha_Modifico=GETDATE()"
            Mi_SQL = Mi_SQL & " WHERE No_Tarjeta='" & Trim(Grid_Registros.TextMatrix(Fila, 0)) & "'"
            Conexion_Base.Execute Mi_SQL
            Pbar_Registros.Value = Pbar_Registros.Value + 1
            Me.Refresh
        Next Fila
        Conexion_Base.CommitTrans
        MsgBox "Importacion realizada con éxito", vbInformation
        Btn_Aplicar.Enabled = False
    End If
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Exit Sub
HANDLER:
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Explorar_SAP_Bajas()
On Error GoTo Fin
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    'Obtener la imagen de una ruta especifica de tipo xls
    Cmd_Archivo.Filter = "Actualizacion Texto (*.txt)|*.txt|Actualizacion en Excel (*.xls)|*.xls"
    Cmd_Archivo.FilterIndex = 1
    Cmd_Archivo.DialogTitle = "Seleccione el archivo"
    If Ruta_Archivo_Pedido <> "" Then
        Cmd_Archivo.InitDir = Ruta_Archivo_Pedido
        Cmd_Archivo.ShowSave
    Else
        Cmd_Archivo.ShowSave
    End If
    'Destino del archivo .txt
    Me.MousePointer = 11
    Txt_Ruta.Text = Cmd_Archivo.FileName
    Txt_Archivo.Text = Cmd_Archivo.FileTitle
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 7
    Grid_Registros.AddItem "No.Empleado" & Chr(9) & "Estatus" & Chr(9) & "Nombre" & Chr(9) & "ApellidoPaterno" & Chr(9) & "ApellidoMaterno" & Chr(9) & "ClaveBaja" & Chr(9) & "Fecha"
    'Abre el archivo de texto
    Open Cmd_Archivo.FileName For Input As #1
        If Not EOF(1) Then
            Do While Not EOF(1)
                'Captura el valor de linea
                Line Input #1, linea
                Linea_Datos = CStr(Trim(linea))
                If Linea_Datos <> "" Then
                    Grupos = Split(Linea_Datos, "|")
                    'Comienza a acumular los totales
                    Grid_Registros.AddItem Trim(Grupos(0)) & Chr(9) & Trim(Grupos(1)) & Chr(9) & Trim(Grupos(2)) _
                        & Chr(9) & Trim(Grupos(3)) & Chr(9) & Trim(Grupos(4)) & Chr(9) & Trim(Grupos(9)) _
                        & Chr(9) & Trim(Grupos(10))
                    Grid_Registros.FixedRows = 1
                End If
            Loop
        End If
    Close #1
    Grid_Registros.ColWidth(0) = 1200
    Grid_Registros.ColAlignment(0) = flexAlignCenterCenter
    Grid_Registros.ColWidth(1) = 600
    Grid_Registros.ColAlignment(1) = flexAlignCenterCenter
    Grid_Registros.ColWidth(2) = 1400
    Grid_Registros.ColAlignment(2) = flexAlignLeftCenter
    Grid_Registros.ColWidth(3) = 1400
    Grid_Registros.ColAlignment(3) = flexAlignLeftCenter
    Grid_Registros.ColWidth(4) = 1400
    Grid_Registros.ColAlignment(4) = flexAlignLeftCenter
    Grid_Registros.ColWidth(5) = 1400
    Grid_Registros.ColAlignment(5) = flexAlignLeftCenter
    Grid_Registros.ColWidth(6) = 1400
    Grid_Registros.ColAlignment(6) = flexAlignLeftCenter
    Btn_Aplicar.Enabled = True
    Me.MousePointer = 0
    Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Close #1
    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Aplicar_Empleados
'DESCRIPCION: Consulta la base de datos de los empleados y si no existe lo da de alta
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 19-Marzo-2014
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Aplicar_Empleados()
Dim Mi_SQL As String
Dim Rs_Empleados As rdoResultset
Dim Rs_Detalles As rdoResultset
Dim Rs_Alta_Cat_Empleados As rdoResultset
Dim Rs_Alta_Detalles As rdoResultset
Dim Fila As Integer
Dim Empleado_ID As String
Dim Departamento_ID As String
Dim Puesto_ID As String
Dim Turno_ID As String
Dim Nombre_Empleado() As String
Dim Supervisor_ID As String
Dim Gap_ID As String
Dim Contador_Nombres As Integer
Dim Mensaje As String

On Error GoTo HANDLER
    If Grid_Registros.Rows < 2 Then Exit Sub
    If MsgBox("¿Esta seguro de actualizar los empleados?", vbQuestion + vbYesNo) = vbYes Then
        Me.MousePointer = 11
        Pbar_Registros.Value = 0
        Pbar_Registros.Max = Grid_Registros.Rows - 1
        Pbar_Registros.Visible = True
        Conexion_Base.BeginTrans
        'Articulos
        For Fila = 1 To Grid_Registros.Rows - 1
            Mensaje = ""
            Mensaje = Trim(UCase(Grid_Registros.TextMatrix(Fila, 2)))
            If Trim(Grid_Registros.TextMatrix(Fila, 0)) <> "" Then
                'Busca ID del puesto
                Mi_SQL = "SELECT Puesto_ID FROM Cat_Puestos WHERE Nombre='" & Trim(Grid_Registros.TextMatrix(Fila, 15)) & "'"
                Set Rs_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Detalles.EOF Then
                    Puesto_ID = Rs_Detalles.rdoColumns("Puesto_ID")
                Else
                    'Crea el puesto si no existe
                    Set Rs_Alta_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Puestos")
                    Puesto_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Puestos", "Puesto_ID"), "00000")
                    Rs_Alta_Detalles.AddNew
                        Rs_Alta_Detalles.rdoColumns("Puesto_ID") = Puesto_ID
                        Rs_Alta_Detalles.rdoColumns("Nombre") = Trim(Grid_Registros.TextMatrix(Fila, 15))
                        Rs_Alta_Detalles.rdoColumns("Abreviatura") = ""
                        Rs_Alta_Detalles.rdoColumns("Descripcion") = ""
                        Rs_Alta_Detalles.rdoColumns("Usuario_Creo") = Nombre_Usuario
                        Rs_Alta_Detalles.rdoColumns("Fecha_Creo") = Now
                    Rs_Alta_Detalles.Update
                    Rs_Alta_Detalles.Close
                End If
                Rs_Detalles.Close
                'Busca ID del departamento
                Mi_SQL = "SELECT Departamento_ID FROM Cat_Departamentos WHERE Nombre='" & Trim(Grid_Registros.TextMatrix(Fila, 6)) & "'"
                Set Rs_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Detalles.EOF Then
                    Departamento_ID = Rs_Detalles.rdoColumns("Departamento_ID")
                Else
                    'Crea el departamento si no existe
                    Set Rs_Alta_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Departamentos")
                    Departamento_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Departamentos", "Departamento_ID"), "00000")
                    Rs_Alta_Detalles.AddNew
                        Rs_Alta_Detalles.rdoColumns("Departamento_ID") = Departamento_ID
                        Rs_Alta_Detalles.rdoColumns("Nombre") = Trim(Grid_Registros.TextMatrix(Fila, 6))
                        Rs_Alta_Detalles.rdoColumns("Clave") = Trim(Grid_Registros.TextMatrix(Fila, 5))
                        Rs_Alta_Detalles.rdoColumns("Comentarios") = ""
                        Rs_Alta_Detalles.rdoColumns("Usuario_Creo") = Nombre_Usuario
                        Rs_Alta_Detalles.rdoColumns("Fecha_Creo") = Now
                    Rs_Alta_Detalles.Update
                    Rs_Alta_Detalles.Close
                End If
                Rs_Detalles.Close
                'Busca ID del turno
                Mi_SQL = "SELECT Turno_ID FROM Cat_Turnos WHERE Turno_ID='" & Format(Grid_Registros.TextMatrix(Fila, 16), "00000") & "'"
                Set Rs_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                If Not Rs_Detalles.EOF Then
                    Turno_ID = Rs_Detalles.rdoColumns("Turno_ID")
                Else
                    Turno_ID = "00001"
                End If
                Rs_Detalles.Close
                'Busca ID del supervisor
'                Mi_SQL = "SELECT Empleado_ID FROM Cat_Empleados WHERE No_Tarjeta='" & Val(Grid_Registros.TextMatrix(Fila, 14)) & "'"
'                Set Rs_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'                If Not Rs_Detalles.EOF Then
'                    Supervisor_ID = Rs_Detalles.rdoColumns("Empleado_ID")
'                Else
'                    Supervisor_ID = ""
'                End If
'                Rs_Detalles.Close
                'Busca ID del gap
'                Mi_SQL = "SELECT Gap_ID FROM Cat_Gaps WHERE Nombre='" & Trim(Grid_Registros.TextMatrix(Fila, 43)) & "'"
'                Set Rs_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
'                If Not Rs_Detalles.EOF Then
'                    Gap_ID = Rs_Detalles.rdoColumns("Gap_ID")
'                Else
'                    Gap_ID = ""
'                End If
'                Rs_Detalles.Close
                'Consulta si existe el empleado
                Mi_SQL = "SELECT * FROM Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE No_Tarjeta='" & Trim(Grid_Registros.TextMatrix(Fila, 2)) & "'"
                Set Rs_Empleados = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
                If Rs_Empleados.EOF Then
            Mensaje = "Agregar:" + Trim(UCase(Grid_Registros.TextMatrix(Fila, 2)))
                    'Da de alta al empleado
                    Set Rs_Alta_Cat_Empleados = Conectar_Ayudante.Recordset_Agregar("Cat_Empleados")
                    Empleado_ID = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Empleados", "Empleado_ID"), "00000")
                    With Rs_Alta_Cat_Empleados
                        .AddNew
                            .rdoColumns("Empleado_ID") = Empleado_ID
                            If Trim(Grid_Registros.TextMatrix(Fila, 0)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 0)) <> "NULL" Then
                                .rdoColumns("Empresa_ID") = Format(Grid_Registros.TextMatrix(Fila, 0), "00000")
                            Else
                                .rdoColumns("Empresa_ID") = Null
                            End If
'                            If Supervisor_ID <> "" Then
'                                .rdoColumns("Supervisor_ID") = Supervisor_ID
'                            Else
'                                .rdoColumns("Supervisor_ID") = Null
'                            End If
                            If Departamento_ID <> "" Then
                                .rdoColumns("Departamento_ID") = Departamento_ID
                            Else
                                .rdoColumns("Departamento_ID") = Null
                            End If
                            If Puesto_ID <> "" Then
                                .rdoColumns("Puesto_ID") = Puesto_ID
                            Else
                                .rdoColumns("Puesto_ID") = Null
                            End If
                            If Turno_ID <> "" Then
                                .rdoColumns("Turno_ID") = Turno_ID
                            Else
                                .rdoColumns("Turno_ID") = Null
                            End If
'                            If Gap_ID <> "" Then
'                                .rdoColumns("Gap_ID") = Gap_ID
'                            Else
'                                .rdoColumns("Gap_ID") = Null
'                            End If
                            'If Trim(Grid_Registros.TextMatrix(Fila, 10)) = "S" Then
                            If Trim(Trim(Grid_Registros.TextMatrix(Fila, 1))) = "B" Then
                                .rdoColumns("Estatus") = "I"
                                If Trim(Grid_Registros.TextMatrix(Fila, 12)) <> "" Then
                                    .rdoColumns("Fecha_Baja") = Format(Trim(Grid_Registros.TextMatrix(Fila, 12)), "MM/dd/yyyy")
                                Else
                                    .rdoColumns("Fecha_Baja") = Format(Now, "MM/dd/yyyy")
                                End If
                                .rdoColumns("Motivo_Baja_ID") = "00001"
                            Else
                                .rdoColumns("Estatus") = "A"
                            End If
                            .rdoColumns("Comentarios_Baja") = ""
'                            If Trim(Grid_Registros.TextMatrix(Fila, 14)) = "0" Then
'                                .rdoColumns("Tipo") = "E"
'                            Else
'                                .rdoColumns("Tipo") = "S"
'                            End If
                            Nombre_Empleado = Split(Trim(Grid_Registros.TextMatrix(Fila, 3)))
                            Contador_Nombres = 0
                            For Contador_Nombres = 0 To UBound(Nombre_Empleado)
                                Select Case Contador_Nombres
                                    Case 0
                                        .rdoColumns("Apellido_Paterno") = Trim(Nombre_Empleado(Contador_Nombres))
                                    Case 1
                                        .rdoColumns("Apellido_Materno") = Trim(Nombre_Empleado(Contador_Nombres))
                                    Case Else
                                        If Contador_Nombres = 2 Then
                                            .rdoColumns("Nombre") = Trim(Nombre_Empleado(Contador_Nombres))
                                        Else
                                            .rdoColumns("Nombre") = Trim(.rdoColumns("Nombre") & " " & Trim(Nombre_Empleado(Contador_Nombres)))
                                        End If
                                End Select
                            Next
                            .rdoColumns("Lugar_Nacimiento") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 19)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 19)) <> "NULL" Then
                                .rdoColumns("Sexo") = Trim(Grid_Registros.TextMatrix(Fila, 19))
                            Else
                                .rdoColumns("Sexo") = "MASCULINO"
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 18)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 18)) <> "NULL" Then
                                .rdoColumns("Fecha_Nacimiento") = Format(Trim(Grid_Registros.TextMatrix(Fila, 18)), "MM/dd/yyyy")
                            Else
                             .rdoColumns("Fecha_Nacimiento") = "01/01/2000"
                            End If
                            .rdoColumns("Clave_Elector") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 21)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 21)) <> "NULL" Then
                                .rdoColumns("Estado_Civil") = Trim(Grid_Registros.TextMatrix(Fila, 21))
                            Else
                                .rdoColumns("Estado_Civil") = ""
                            End If
                            .rdoColumns("Cedula_Identidad_Ciudadana") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 20)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 20)) <> "NULL" Then
                                .rdoColumns("RFC") = Trim(Grid_Registros.TextMatrix(Fila, 20))
                            Else
                                .rdoColumns("RFC") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 31)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 31)) <> "NULL" Then
                                .rdoColumns("Curp") = Trim(Grid_Registros.TextMatrix(Fila, 31))
                            Else
                                .rdoColumns("Curp") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 22)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 22)) <> "NULL" Then
                                .rdoColumns("Nss") = Trim(Grid_Registros.TextMatrix(Fila, 22))
                            Else
                                .rdoColumns("Nss") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 23)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 23)) <> "NULL" Then
                                .rdoColumns("Direccion") = Trim(Grid_Registros.TextMatrix(Fila, 23))
                            Else
                                .rdoColumns("Direccion") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 30)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 30)) <> "NULL" Then
                                .rdoColumns("Direccion") = Trim(.rdoColumns("Direccion")) & " Núm. Ext. " & Trim(Grid_Registros.TextMatrix(Fila, 30))
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 24)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 24)) <> "NULL" Then
                                .rdoColumns("Colonia") = Trim(Grid_Registros.TextMatrix(Fila, 24))
                            Else
                                .rdoColumns("Colonia") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 26)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 26)) <> "NULL" Then
                                .rdoColumns("Codigo_Postal") = Trim(Grid_Registros.TextMatrix(Fila, 26))
                            Else
                                .rdoColumns("Codigo_Postal") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 27)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 27)) <> "NULL" Then
                                .rdoColumns("Ciudad") = Trim(Grid_Registros.TextMatrix(Fila, 27))
                            Else
                                .rdoColumns("Ciudad") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 25)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 25)) <> "NULL" Then
                                .rdoColumns("Estado") = Trim(Grid_Registros.TextMatrix(Fila, 25))
                            Else
                                .rdoColumns("Estado") = ""
                            End If
                            .rdoColumns("Imagen_Perfil") = Trim(Grid_Registros.TextMatrix(Fila, 2)) & ".JPG"
                            .rdoColumns("Nomipaq_ID") = ""
                            .rdoColumns("No_Tarjeta") = Trim(UCase(Grid_Registros.TextMatrix(Fila, 2)))
                            If Trim(Grid_Registros.TextMatrix(Fila, 4)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 4)) <> "NULL" Then
                                .rdoColumns("Fecha_Ingreso") = Format(Trim(Grid_Registros.TextMatrix(Fila, 4)), "MM/dd/yyyy")
                            Else
                                .rdoColumns("Fecha_Ingreso") = Format(Now, "MM/dd/yyyy")
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 7)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 7)) <> "NULL" Then
                                .rdoColumns("Tipo_Empleado") = Trim(Grid_Registros.TextMatrix(Fila, 7))
                            Else
                                .rdoColumns("Tipo_Empleado") = ""
                            End If
                            .rdoColumns("Tipo_Contratacion") = "Planta"
                            If Trim(Grid_Registros.TextMatrix(Fila, 32)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 32)) <> "NULL" Then
                                .rdoColumns("Salario_Diario") = Trim(Grid_Registros.TextMatrix(Fila, 32))
                            Else
                                .rdoColumns("Salario_Diario") = ""
                            End If
                            .rdoColumns("Salario_Diario_Variable") = 0
                            .rdoColumns("Trabaja_Domingos") = "N"
                            .rdoColumns("Infonavit") = "N"
                            .rdoColumns("Retardos") = 0
                            .rdoColumns("Fecha_Retardo") = "01/01/1960"
                            .rdoColumns("En_Caso_Emergencia") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 28)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 28)) <> "NULL" Then
                                .rdoColumns("Telefono_Emergencia1") = Trim(Grid_Registros.TextMatrix(Fila, 28))
                            Else
                                .rdoColumns("Telefono_Emergencia1") = ""
                            End If
                            .rdoColumns("Telefono_Emergencia2") = ""
                            .rdoColumns("Alergia1") = ""
                            .rdoColumns("Alergia2") = ""
                            .rdoColumns("Alergia3") = ""
                            .rdoColumns("Usuario_Creo") = Nombre_Usuario
                            .rdoColumns("Fecha_Creo") = Now
                        .Update
                    End With
                    Rs_Alta_Cat_Empleados.Close
                Else
            Mensaje = "Editar:" + Trim(UCase(Grid_Registros.TextMatrix(Fila, 2)))
                    With Rs_Empleados
                        .Edit
                            If Trim(Grid_Registros.TextMatrix(Fila, 0)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 0)) <> "NULL" Then
                                .rdoColumns("Empresa_ID") = Format(Grid_Registros.TextMatrix(Fila, 0), "00000")
                            Else
                                .rdoColumns("Empresa_ID") = Null
                            End If
'                            If Supervisor_ID <> "" Then
'                                .rdoColumns("Supervisor_ID") = Supervisor_ID
'                            Else
'                                .rdoColumns("Supervisor_ID") = Null
'                            End If
                            If Departamento_ID <> "" Then
                                .rdoColumns("Departamento_ID") = Departamento_ID
                            Else
                                .rdoColumns("Departamento_ID") = Null
                            End If
                            If Puesto_ID <> "" Then
                                .rdoColumns("Puesto_ID") = Puesto_ID
                            Else
                                .rdoColumns("Puesto_ID") = Null
                            End If
'                            If Turno_ID <> "" Then
'                                .rdoColumns("Turno_ID") = Turno_ID
'                            Else
'                                .rdoColumns("Turno_ID") = Null
'                            End If
'                            If Gap_ID <> "" Then
'                                .rdoColumns("Gap_ID") = Gap_ID
'                            Else
'                                .rdoColumns("Gap_ID") = Null
'                            End If
                            'If Trim(Grid_Registros.TextMatrix(Fila, 10)) = "S" Then
                            If Trim(Trim(Grid_Registros.TextMatrix(Fila, 1))) = "B" Then
                                .rdoColumns("Estatus") = "I"
                                If Trim(Grid_Registros.TextMatrix(Fila, 12)) <> "" Then
                                    .rdoColumns("Fecha_Baja") = Format(Trim(Grid_Registros.TextMatrix(Fila, 12)), "MM/dd/yyyy")
                                Else
                                    .rdoColumns("Fecha_Baja") = Format(Now, "MM/dd/yyyy")
                                End If
                                .rdoColumns("Motivo_Baja_ID") = "00001"
                            Else
                                .rdoColumns("Estatus") = "A"
                            End If
'                            .rdoColumns("Comentarios_Baja") = ""
'                            If Trim(Grid_Registros.TextMatrix(Fila, 14)) = "0" Then
'                                .rdoColumns("Tipo") = "E"
'                            Else
'                                .rdoColumns("Tipo") = "S"
'                            End If
                            Nombre_Empleado = Split(Trim(Grid_Registros.TextMatrix(Fila, 3)))
                            Contador_Nombres = 0
                            For Contador_Nombres = 0 To UBound(Nombre_Empleado)
                                Select Case Contador_Nombres
                                    Case 0
                                        .rdoColumns("Apellido_Paterno") = Trim(Nombre_Empleado(Contador_Nombres))
                                    Case 1
                                        .rdoColumns("Apellido_Materno") = Trim(Nombre_Empleado(Contador_Nombres))
                                    Case Else
                                        If Contador_Nombres = 2 Then
                                            .rdoColumns("Nombre") = Trim(Nombre_Empleado(Contador_Nombres))
                                        Else
                                            .rdoColumns("Nombre") = Trim(.rdoColumns("Nombre") & " " & Trim(Nombre_Empleado(Contador_Nombres)))
                                        End If
                                End Select
                            Next
'                            .rdoColumns("Lugar_Nacimiento") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 19)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 19)) <> "NULL" Then
                                .rdoColumns("Sexo") = Trim(Grid_Registros.TextMatrix(Fila, 19))
                            Else
                                .rdoColumns("Sexo") = "MASCULINO"
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 18)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 18)) <> "NULL" Then
                                .rdoColumns("Fecha_Nacimiento") = Format(Trim(Grid_Registros.TextMatrix(Fila, 18)), "MM/dd/yyyy")
                            Else
                             .rdoColumns("Fecha_Nacimiento") = "01/01/2000"
                            End If
'                            .rdoColumns("Clave_Elector") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 21)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 21)) <> "NULL" Then
                                .rdoColumns("Estado_Civil") = Trim(Grid_Registros.TextMatrix(Fila, 21))
                            Else
                                .rdoColumns("Estado_Civil") = ""
                            End If
'                            .rdoColumns("Cedula_Identidad_Ciudadana") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 20)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 20)) <> "NULL" Then
                                .rdoColumns("RFC") = Trim(Grid_Registros.TextMatrix(Fila, 20))
                            Else
                                .rdoColumns("RFC") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 31)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 31)) <> "NULL" Then
                                .rdoColumns("Curp") = Trim(Grid_Registros.TextMatrix(Fila, 31))
                            Else
                                .rdoColumns("Curp") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 22)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 22)) <> "NULL" Then
                                .rdoColumns("Nss") = Trim(Grid_Registros.TextMatrix(Fila, 22))
                            Else
                                .rdoColumns("Nss") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 23)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 23)) <> "NULL" Then
                                .rdoColumns("Direccion") = Trim(Grid_Registros.TextMatrix(Fila, 23))
                            Else
                                .rdoColumns("Direccion") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 30)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 30)) <> "NULL" Then
                                .rdoColumns("Direccion") = Trim(.rdoColumns("Direccion")) & " Núm. Ext. " & Trim(Grid_Registros.TextMatrix(Fila, 30))
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 24)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 24)) <> "NULL" Then
                                .rdoColumns("Colonia") = Trim(Grid_Registros.TextMatrix(Fila, 24))
                            Else
                                .rdoColumns("Colonia") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 26)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 26)) <> "NULL" Then
                                .rdoColumns("Codigo_Postal") = Trim(Grid_Registros.TextMatrix(Fila, 26))
                            Else
                                .rdoColumns("Codigo_Postal") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 27)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 27)) <> "NULL" Then
                                .rdoColumns("Ciudad") = Trim(Grid_Registros.TextMatrix(Fila, 27))
                            Else
                                .rdoColumns("Ciudad") = ""
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 25)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 25)) <> "NULL" Then
                                .rdoColumns("Estado") = Trim(Grid_Registros.TextMatrix(Fila, 25))
                            Else
                                .rdoColumns("Estado") = ""
                            End If
'                            .rdoColumns("Imagen_Perfil") = Trim(Grid_Registros.TextMatrix(Fila, 2)) & ".JPG"
'                            .rdoColumns("Nomipaq_ID") = ""
                            .rdoColumns("No_Tarjeta") = Trim(UCase(Grid_Registros.TextMatrix(Fila, 2)))
                            If Trim(Grid_Registros.TextMatrix(Fila, 4)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 4)) <> "NULL" Then
                                .rdoColumns("Fecha_Ingreso") = Format(Trim(Grid_Registros.TextMatrix(Fila, 4)), "MM/dd/yyyy")
                            Else
                                .rdoColumns("Fecha_Ingreso") = Format(Now, "MM/dd/yyyy")
                            End If
                            If Trim(Grid_Registros.TextMatrix(Fila, 7)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 7)) <> "NULL" Then
                                .rdoColumns("Tipo_Empleado") = Trim(Grid_Registros.TextMatrix(Fila, 7))
                            Else
                                .rdoColumns("Tipo_Empleado") = ""
                            End If
'                            .rdoColumns("Tipo_Contratacion") = "Planta"
'                            .rdoColumns("Fecha_Termino_Contrato") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 32)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 32)) <> "NULL" Then
                                .rdoColumns("Salario_Diario") = Trim(Grid_Registros.TextMatrix(Fila, 32))
                            Else
                                .rdoColumns("Salario_Diario") = ""
                            End If
'                            .rdoColumns("Salario_Diario_Variable") = 0
'                            .rdoColumns("Trabaja_Domingos") = "N"
'                            .rdoColumns("Infonavit") = "N"
'                            .rdoColumns("Retardos") = 0
'                            .rdoColumns("Fecha_Retardo") = "01/01/1960"
'                            .rdoColumns("En_Caso_Emergencia") = ""
                            If Trim(Grid_Registros.TextMatrix(Fila, 28)) <> "" And Trim(Grid_Registros.TextMatrix(Fila, 28)) <> "NULL" Then
                                .rdoColumns("Telefono_Emergencia1") = Trim(Grid_Registros.TextMatrix(Fila, 28))
                            Else
                                .rdoColumns("Telefono_Emergencia1") = ""
                            End If
'                            .rdoColumns("Telefono_Emergencia2") = ""
'                            .rdoColumns("Alergia1") = ""
'                            .rdoColumns("Alergia2") = ""
'                            .rdoColumns("Alergia3") = ""
                            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
                            .rdoColumns("Fecha_Modifico") = Now
                        .Update
                    End With
                End If
                Rs_Empleados.Close
            End If
            Pbar_Registros.Value = Pbar_Registros.Value + 1
            Frm_Ope_Importacion_Datos.Refresh
            Frm_Ope_Importacion_Datos.Pbar_Registros.Refresh
            DoEvents
        Next Fila
            Mensaje = "Comit"
        Conexion_Base.CommitTrans
        MsgBox "Importacion realizada con éxito", vbInformation
        Btn_Aplicar.Enabled = False
    End If
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
Exit Sub
HANDLER:
    Me.MousePointer = 0
    Pbar_Registros.Visible = False
    Conexion_Base.RollbackTrans
    MsgBox Mensaje
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Aplicar_Click()
    Select Case Cmb_Tipo_Producto.Text
        Case "SQL"
            Aplicar_Empleados
        Case "EXCEL"
            Aplicar_Empleados
'        Case "Turnos"
'            Aplicar_Turnos
'        Case "SAP Altas"
'            Aplicar_SAP_Altas
'        Case "SAP Bajas"
'            Aplicar_SAP_Bajas
    End Select
End Sub


Private Sub Btn_Capturar_Click()
    MDIFrm_Apl_Principal.MousePointer = 11
    If MsgBox("¿Esta seguro de cargar el catálogo?", vbQuestion + vbYesNo) = vbYes Then
        MsgBox "Importación realizada con éxito", vbInformation
        Btn_Capturar.Enabled = False
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Costo_Promedio
'DESCRIPCION: Recalcula el costo promedio de las entradas
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 29-Septiembre-2011
'MODIFICO      :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************'
Private Sub Btn_Costo_Promedio_Click()
Dim Mi_SQL As String
Dim Rs_Consulta_Entrada_Detalles As rdoResultset
Dim Rs_Actualiza_Entrada_Detalles As rdoResultset
Dim Costo_Promedio As Double
Dim Cantidad As Long
Dim Producto_ID As String

On Error GoTo HANDLER
    If Nombre_Usuario_ID <> "00001" Then Exit Sub
    If MsgBox("¿Está seguro de recalcular el costo promedio de las entradas?", vbQuestion + vbYesNo, "Recalculo de Entradas") = vbNo Then Exit Sub
    Me.MousePointer = 11
    'Conexion_Base.BeginTrans
    'Consulta los detalles de las entradas que tengan costo promedio 0 para calcular su costo
    Mi_SQL = "SELECT Cat_Productos.Producto_ID,Cat_Productos.Costo AS Costo_Producto,Cat_Productos.Costo_Promedio,Cat_Productos.Existencia,Ope_Entradas_Detalles.No_Entrada,Ope_Entradas_Detalles.Cantidad,ISNULL(Ope_Entradas_Detalles.Costo_Compra,0) AS Costo_Entrada"
    Mi_SQL = Mi_SQL & " FROM Cat_Productos,Ope_Entradas_Detalles"
    Mi_SQL = Mi_SQL & " WHERE Cat_Productos.Producto_ID=Ope_Entradas_Detalles.Producto_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Entradas_Detalles.Tipo_Producto='LINEA'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Productos.Producto_ID,Ope_Entradas_Detalles.No_Entrada"
    Set Rs_Consulta_Entrada_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Entrada_Detalles.EOF
        With Rs_Consulta_Entrada_Detalles
            'Resetea la cantidad
            If Producto_ID <> .rdoColumns("Producto_ID") Then
                Cantidad = 0
                Debug.Print .rdoColumns("Producto_ID")
            End If
            Producto_ID = .rdoColumns("Producto_ID")
            'Calcula el costo promedio
            Costo_Promedio = (Cantidad * Costo_Promedio) + (.rdoColumns("Cantidad") * .rdoColumns("Costo_Entrada"))
            If Val(Cantidad + .rdoColumns("Cantidad")) > 0 Then  'Obtiene el costo promedio
                If Costo_Promedio > 0 Then
                    Costo_Promedio = Costo_Promedio / (Cantidad + .rdoColumns("Cantidad"))
                Else
                    'If .rdoColumns("Costo_Promedio") > 0 Then
                    '    Costo_Promedio = Val(.rdoColumns("Costo_Promedio"))
                    'Else
                        If .rdoColumns("Costo_Entrada") Then
                            Costo_Promedio = Val(.rdoColumns("Costo_Entrada"))
                        Else
                            Costo_Promedio = Val(.rdoColumns("Costo_Producto"))
                        End If
                    'End If
                End If
            Else
                'If .rdoColumns("Costo_Promedio") > 0 Then
                '    Costo_Promedio = Val(.rdoColumns("Costo_Promedio"))
                'Else
                    If .rdoColumns("Costo_Entrada") Then
                        Costo_Promedio = Val(.rdoColumns("Costo_Entrada"))
                    Else
                        Costo_Promedio = Val(.rdoColumns("Costo_Producto"))
                    End If
                'End If
            End If
            'Actualiza el costo promedio de la entrada
            Mi_SQL = "UPDATE Ope_Entradas_Detalles"
            Mi_SQL = Mi_SQL & " SET Costo_Promedio=" & Costo_Promedio
            Mi_SQL = Mi_SQL & " WHERE No_Entrada='" & .rdoColumns("No_Entrada") & "'"
            Mi_SQL = Mi_SQL & " AND Producto_ID='" & .rdoColumns("Producto_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Tipo_Producto='LINEA'"
            Conexion_Base.Execute Mi_SQL
            'Actualiza el costo promedio en la tabla de productos
            Mi_SQL = "UPDATE Cat_Productos"
            Mi_SQL = Mi_SQL & " SET Costo_Promedio=" & Costo_Promedio
            Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & .rdoColumns("Producto_ID") & "'"
            Conexion_Base.Execute Mi_SQL
            'Actualiza las salidas de la entrada
            Mi_SQL = "UPDATE Ope_Salidas_Detalles"
            Mi_SQL = Mi_SQL & " SET Costo_Promedio=" & Costo_Promedio
            Mi_SQL = Mi_SQL & " WHERE Producto_ID='" & .rdoColumns("Producto_ID") & "'"
            Mi_SQL = Mi_SQL & " AND Tipo_Producto='LINEA'"
            Conexion_Base.Execute Mi_SQL
            Cantidad = Cantidad + .rdoColumns("Cantidad")
            .MoveNext
        End With
    Wend
    Rs_Consulta_Entrada_Detalles.Close
    'Conexion_Base.CommitTrans
    MsgBox "El costo promedio de las entradas ha sido calculado", vbInformation
    Me.MousePointer = 0
    Exit Sub
HANDLER:
    Me.MousePointer = 0
    'Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Explorar_Turnos()

On Error GoTo Fin
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    'Obtener la imagen de una ruta especifica de tipo xls
    Cmd_Archivo.Filter = "Actualizacion Texto (*.txt)|*.txt|Actualizacion en Excel (*.xls)|*.xls"
    Cmd_Archivo.FilterIndex = 1
    Cmd_Archivo.DialogTitle = "Seleccione el archivo"
    If Ruta_Archivo_Pedido <> "" Then
        Cmd_Archivo.InitDir = Ruta_Archivo_Pedido
        Cmd_Archivo.ShowSave
    Else
        Cmd_Archivo.ShowSave
    End If
    'Destino del archivo .txt
    Me.MousePointer = 11
    Txt_Ruta.Text = Cmd_Archivo.FileName
    Txt_Archivo.Text = Cmd_Archivo.FileTitle
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 7
    Grid_Registros.AddItem "No.Empleado" & Chr(9) & "Departamento" & Chr(9) & "Turno" & Chr(9) & "GAP" & Chr(9) & "Supervisor" & Chr(9) & "Puesto" & Chr(9) & "Seccion"
    'Abre el archivo de texto
    Open Cmd_Archivo.FileName For Input As #1
        If Not EOF(1) Then
            Do While Not EOF(1)
                'Captura el valor de linea
                Line Input #1, linea
                Linea_Datos = CStr(Trim(linea))
                If Linea_Datos <> "" Then
                    Grupos = Split(Linea_Datos, Chr(9))
                    If Grupos(0) <> "No.Empleado" Then
                        'Comienza a acumular los totales
                        Grid_Registros.AddItem Trim(Grupos(0)) & Chr(9) & Trim(Grupos(1)) & Chr(9) & Trim(Grupos(2)) & Chr(9) & Trim(Grupos(3)) & Chr(9) & Trim(Grupos(4)) & Chr(9) & Trim(Grupos(5)) & Chr(9) & Trim(Grupos(6))
                        Grid_Registros.FixedRows = 1
                    End If
                End If
            Loop
        End If
    Close #1
    Grid_Registros.ColWidth(0) = 1100
    Grid_Registros.ColAlignment(0) = flexAlignCenterCenter
    Grid_Registros.ColWidth(1) = 1100
    Grid_Registros.ColAlignment(1) = flexAlignLeftCenter
    Grid_Registros.ColWidth(2) = 1100
    Grid_Registros.ColAlignment(2) = flexAlignLeftCenter
    Grid_Registros.ColWidth(3) = 1100
    Grid_Registros.ColAlignment(3) = flexAlignLeftCenter
    Grid_Registros.ColWidth(4) = 1100
    Grid_Registros.ColAlignment(4) = flexAlignLeftCenter
    Grid_Registros.ColWidth(5) = 1100
    Grid_Registros.ColAlignment(5) = flexAlignLeftCenter
    Grid_Registros.ColWidth(6) = 1100
    Grid_Registros.ColAlignment(6) = flexAlignLeftCenter
    Btn_Aplicar.Enabled = True
    Me.MousePointer = 0
    Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Close #1
    End If
End Sub

Private Sub Explorar_SAP_Altas()
On Error GoTo Fin
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    'Obtener la imagen de una ruta especifica de tipo xls
    Cmd_Archivo.Filter = "Actualizacion Texto (*.txt)|*.txt|Actualizacion en Excel (*.xls)|*.xls"
    Cmd_Archivo.FilterIndex = 1
    Cmd_Archivo.DialogTitle = "Seleccione el archivo"
    If Ruta_Archivo_Pedido <> "" Then
        Cmd_Archivo.InitDir = Ruta_Archivo_Pedido
        Cmd_Archivo.ShowSave
    Else
        Cmd_Archivo.ShowSave
    End If
    'Destino del archivo .txt
    Me.MousePointer = 11
    Txt_Ruta.Text = Cmd_Archivo.FileName
    Txt_Archivo.Text = Cmd_Archivo.FileTitle
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 9
    Grid_Registros.AddItem "No.Empleado" & Chr(9) & "Estatus" & Chr(9) & "Nombre" & Chr(9) & "ApellidoPaterno" & Chr(9) & "ApellidoMaterno" & Chr(9) & "Seccion" & Chr(9) & "Departamento" & Chr(9) & "Puesto" & Chr(9) & "Turno"
    'Abre el archivo de texto
    Open Cmd_Archivo.FileName For Input As #1
        If Not EOF(1) Then
            Do While Not EOF(1)
                'Captura el valor de linea
                Line Input #1, linea
                Linea_Datos = CStr(Trim(linea))
                If Linea_Datos <> "" Then
                    Grupos = Split(Linea_Datos, "|")
                    'Comienza a acumular los totales
                    Grid_Registros.AddItem Trim(Grupos(0)) & Chr(9) & Trim(Grupos(1)) & Chr(9) & Trim(Grupos(2)) _
                        & Chr(9) & Trim(Grupos(3)) & Chr(9) & Trim(Grupos(4)) & Chr(9) & Trim(Grupos(5)) _
                        & Chr(9) & Trim(Grupos(6)) & Chr(9) & Trim(Grupos(7)) & Chr(9) & Trim(Grupos(8))
                    Grid_Registros.FixedRows = 1
                End If
            Loop
        End If
    Close #1
    Grid_Registros.ColWidth(0) = 1200
    Grid_Registros.ColAlignment(0) = flexAlignCenterCenter
    Grid_Registros.ColWidth(1) = 600
    Grid_Registros.ColAlignment(1) = flexAlignCenterCenter
    Grid_Registros.ColWidth(2) = 1400
    Grid_Registros.ColAlignment(2) = flexAlignLeftCenter
    Grid_Registros.ColWidth(3) = 1400
    Grid_Registros.ColAlignment(3) = flexAlignLeftCenter
    Grid_Registros.ColWidth(4) = 1400
    Grid_Registros.ColAlignment(4) = flexAlignLeftCenter
    Grid_Registros.ColWidth(5) = 700
    Grid_Registros.ColAlignment(5) = flexAlignLeftCenter
    Grid_Registros.ColWidth(6) = 700
    Grid_Registros.ColAlignment(6) = flexAlignLeftCenter
    Grid_Registros.ColWidth(7) = 700
    Grid_Registros.ColAlignment(7) = flexAlignLeftCenter
    Grid_Registros.ColWidth(8) = 700
    Grid_Registros.ColAlignment(8) = flexAlignLeftCenter
    Btn_Aplicar.Enabled = True
    Me.MousePointer = 0
    Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Close #1
    End If
End Sub

Private Sub Explorar_Empleados_Vista()
Dim iLinea As Integer
Dim Rs_Vista_Empleados As rdoResultset

On Error GoTo Fin
    iLinea = 1
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    iLinea = 2
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 28
    iLinea = 3
    Grid_Registros.AddItem "empresa" _
        & Chr(9) & "estatus" _
        & Chr(9) & "codigo" _
        & Chr(9) & "nombre" _
        & Chr(9) & "fchalta" _
        & Chr(9) & "centro" _
        & Chr(9) & "nomdepto" _
        & Chr(9) & "sindicato" _
        & Chr(9) & "fch_ini_contrato" _
        & Chr(9) & "fchterm" _
        & Chr(9) & "activo" _
        & Chr(9) & "puesto" _
        & Chr(9) & "fchbaja" _
        & Chr(9) & "supervisornombre" _
        & Chr(9) & "supervisor" _
        & Chr(9) & "actividad" _
        & Chr(9) & "turno" _
        & Chr(9) & "tipocontrato" _
        & Chr(9) & "fecha_nac" _
        & Chr(9) & "sexo" _
        & Chr(9) & "rfc" _
        & Chr(9) & "estado_civil" _
        & Chr(9) & "nss" & Chr(9) & "domicilio" & Chr(9) & "colonia" & Chr(9) & "poblacion" & Chr(9) & "codpostal" & Chr(9) & "sueldo"
    iLinea = 4
    'Establece la conexión a la vista
    Conectar_Ayudante.Conexion_Servidor_Vista
    iLinea = 5
    'Realiza la consulta de la vista
    Mi_SQL = "SELECT * FROM Sistema_Recursos_Humanos"
    Mi_SQL = Mi_SQL & " WHERE (fchbaja='19000101'"    'La fecha baja es parámetro para saber si siguen activos
    Mi_SQL = Mi_SQL & " OR fchbaja>'" & Format(DateAdd("M", -1, Now), "yyyyMMdd") & "')"
    Mi_SQL = Mi_SQL & " AND codigo<='10000'"            'Empleados mayores a 10,000 no considerar
    Mi_SQL = Mi_SQL & " ORDER BY codigo"
    iLinea = 6
'    Set Rs_Vista_Empleados = Conexion_Servidor.OpenResultset(Mi_SQL, 2)
    
    Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Servidor
        .SQL = Mi_SQL
        .LockType = rdConcurReadOnly
        .CursorType = rdUseOdbc
        .RowsetSize = 20
        Set Rs_Vista_Empleados = .OpenResultset(rdOpenStatic)
    End With
    
    
    iLinea = 7
    While Not Rs_Vista_Empleados.EOF
        iLinea = 8
        'Llena el grid
        Grid_Registros.AddItem Rs_Vista_Empleados.rdoColumns("empresa") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("estatus") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("codigo") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("nombre") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fchalta") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("centro") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("nomdepto") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("sindicato") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fch_ini_contrato") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fchterm") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("activo") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("puesto") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fchbaja") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("supervisornombre") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("supervisor") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("actividad") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("turno") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("tipocontrato") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fchnac") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("sexo") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("rfc") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("edocivil") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("afiliacion") & Chr(9) & Rs_Vista_Empleados.rdoColumns("domicilio") & Chr(9) & Rs_Vista_Empleados.rdoColumns("colonia") & Chr(9) & Rs_Vista_Empleados.rdoColumns("poblacion") & Chr(9) & Rs_Vista_Empleados.rdoColumns("codpostal") & Chr(9) & Rs_Vista_Empleados.rdoColumns("sueldo")
        iLinea = 9
        Rs_Vista_Empleados.MoveNext
    Wend
    iLinea = 10
    Rs_Vista_Empleados.Close
    iLinea = 11
    Conexion_Servidor.Close
    iLinea = 12
    If Grid_Registros.Rows > 1 Then
        Grid_Registros.FixedCols = 3
        Grid_Registros.FixedRows = 1
        Grid_Registros.ColWidth(0) = 700        'empresa
        Grid_Registros.ColWidth(1) = 0          'estatus
        Grid_Registros.ColWidth(2) = 800        'codigo
        Grid_Registros.ColWidth(3) = 2000       'nombre
        Grid_Registros.ColWidth(4) = 1000       'fchalta
        Grid_Registros.ColWidth(5) = 0          'centro
        Grid_Registros.ColWidth(6) = 1500       'nomdepto
        Grid_Registros.ColWidth(7) = 0          'sindicato
        Grid_Registros.ColWidth(8) = 0          'fch_ini_contrato
        Grid_Registros.ColWidth(9) = 0          'fchterm
        Grid_Registros.ColWidth(10) = 550       'activo
        Grid_Registros.ColWidth(11) = 0         'puesto
        Grid_Registros.ColWidth(12) = 1000      'fchbaja
        Grid_Registros.ColWidth(13) = 1500      'supervisornombre
        Grid_Registros.ColWidth(14) = 500       'supervisor
        Grid_Registros.ColWidth(15) = 1500      'actividad
        Grid_Registros.ColWidth(16) = 500       'turno
        Grid_Registros.ColWidth(17) = 0         'tipocontrato
        Grid_Registros.ColWidth(18) = 500       'fecha_nacimiento
        Grid_Registros.ColWidth(19) = 200       'sexo
        Grid_Registros.ColWidth(20) = 800       'rfc
        Grid_Registros.ColWidth(21) = 500       'estado civil
        Grid_Registros.ColWidth(22) = 800       'nss
        Grid_Registros.ColWidth(23) = 0         'domicilio
        Grid_Registros.ColWidth(24) = 0         'colonia
        Grid_Registros.ColWidth(25) = 0         'poblacion
        Grid_Registros.ColWidth(26) = 0         'codpostal
        Grid_Registros.ColWidth(27) = 0         'sueldo
    End If
    iLinea = 13
    Btn_Aplicar.Enabled = True
    iLinea = 14
    Me.MousePointer = 0
Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "[Frm_Ope_Importacion_Datos|Explorar_Empleados_Vista|Linea:" & iLinea & "]" & Err.Description, vbExclamation
        Close #1
    End If
End Sub

Private Sub Explorar_Empleados_Nueva_Vista()
Dim iLinea As Integer
Dim Rs_Vista_Empleados As rdoResultset

On Error GoTo Fin
    iLinea = 1
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    iLinea = 2
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 33
    iLinea = 3
    Grid_Registros.AddItem "empresa" _
        & Chr(9) & "estatus" _
        & Chr(9) & "codigo" _
        & Chr(9) & "nombre" _
        & Chr(9) & "fchalta" _
        & Chr(9) & "centro" _
        & Chr(9) & "nomdepto" _
        & Chr(9) & "sindicato" _
        & Chr(9) & "fch_ini_contrato" _
        & Chr(9) & "fchterm" _
        & Chr(9) & "activo" _
        & Chr(9) & "puesto" _
        & Chr(9) & "fchbaja" _
        & Chr(9) & "supervisornombre" _
        & Chr(9) & "supervisor" _
        & Chr(9) & "actividad" _
        & Chr(9) & "turno" _
        & Chr(9) & "tipocontrato" _
        & Chr(9) & "fecha_nac" _
        & Chr(9) & "sexo" _
        & Chr(9) & "rfc" _
        & Chr(9) & "estado_civil" _
        & Chr(9) & "nss" & Chr(9) & "domicilio" & Chr(9) & "colonia" & Chr(9) & "poblacion" & Chr(9) & "codpostal" & Chr(9) & "municipio" & Chr(9) & "telefono" & Chr(9) & "nacional" & Chr(9) & "numeroExterior" & Chr(9) & "curp" & Chr(9) & "sueldo"
    iLinea = 4
    'Establece la conexión a la vista
    Conectar_Ayudante.Conexion_Servidor_Vista
    iLinea = 5
    'Realiza la consulta de la vista
    Mi_SQL = "SELECT * FROM Sistema_Recursos_Humanos1"
    Mi_SQL = Mi_SQL & " WHERE"
'    Mi_SQL = Mi_SQL & " WHERE (fch_baja='19000101'"    'La fecha baja es parámetro para saber si siguen activos
'    Mi_SQL = Mi_SQL & " OR fch_baja>'" & Format(DateAdd("M", -1, Now), "yyyyMMdd") & "')"
    Mi_SQL = Mi_SQL & " No_Nomina <= '10000'"            'Empleados mayores a 10,000 no considerar
    Mi_SQL = Mi_SQL & " ORDER BY No_Nomina"
    iLinea = 6
'    Set Rs_Vista_Empleados = Conexion_Servidor.OpenResultset(Mi_SQL, 2)
    
    Dim Mi_Query As New rdoQuery 'Obtiene los valores de la consulta
    With Mi_Query
        Set .ActiveConnection = Conexion_Servidor
        .SQL = Mi_SQL
        .LockType = rdConcurReadOnly
        .CursorType = rdUseOdbc
        .RowsetSize = 20
        Set Rs_Vista_Empleados = .OpenResultset(rdOpenStatic)
    End With
    
    Dim Sueldo As Double
    
    iLinea = 7
    While Not Rs_Vista_Empleados.EOF
        iLinea = 8
        Sueldo = 0
        If Not IsNull(Rs_Vista_Empleados.rdoColumns("sueldo")) Then
            Sueldo = Rs_Vista_Empleados.rdoColumns("sueldo") * 0
        End If
        'Llena el grid
        Grid_Registros.AddItem Rs_Vista_Empleados.rdoColumns("empresa") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("ultmov") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("No_Nomina") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("Nombre") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fchalta") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("centro") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("nomdepto") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("Sindicato") _
            & Chr(9) & "" _
            & Chr(9) & "" _
            & Chr(9) & "" _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("actividad") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("fch_baja") _
            & Chr(9) & "" _
            & Chr(9) & "" _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("actividad") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("turno") _
            & Chr(9) & "" _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("Fecha_nacimiento") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("Sexo") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("rfc") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("edocivil") _
            & Chr(9) & Rs_Vista_Empleados.rdoColumns("afiliacion") & Chr(9) & Rs_Vista_Empleados.rdoColumns("domicilio") & Chr(9) & Rs_Vista_Empleados.rdoColumns("colonia") & Chr(9) & Rs_Vista_Empleados.rdoColumns("poblacion") & Chr(9) & Rs_Vista_Empleados.rdoColumns("codpostal") & Chr(9) & Rs_Vista_Empleados.rdoColumns("municipio") & Chr(9) & Rs_Vista_Empleados.rdoColumns("telefono") & Chr(9) & Rs_Vista_Empleados.rdoColumns("nacional") & Chr(9) & Rs_Vista_Empleados.rdoColumns("numeroExterior") & Chr(9) & Rs_Vista_Empleados.rdoColumns("curp") & Chr(9) & Sueldo
        iLinea = 9
        Rs_Vista_Empleados.MoveNext
    Wend
    iLinea = 10
    Rs_Vista_Empleados.Close
    iLinea = 11
    Conexion_Servidor.Close
    iLinea = 12
    If Grid_Registros.Rows > 1 Then
        Grid_Registros.FixedCols = 3
        Grid_Registros.FixedRows = 1
        Grid_Registros.ColWidth(0) = 700        'empresa
        Grid_Registros.ColWidth(1) = 0          'ultmov
        Grid_Registros.ColWidth(2) = 800        'codigo
        Grid_Registros.ColWidth(3) = 2000       'nombre
        Grid_Registros.ColWidth(4) = 1000       'fchalta
        Grid_Registros.ColWidth(5) = 0          'centro
        Grid_Registros.ColWidth(6) = 1500       'nomdepto
        Grid_Registros.ColWidth(7) = 0          'sindicato
        Grid_Registros.ColWidth(8) = 0          'fch_ini_contrato
        Grid_Registros.ColWidth(9) = 0          'fchterm
        Grid_Registros.ColWidth(10) = 550       'activo
        Grid_Registros.ColWidth(11) = 0         'puesto
        Grid_Registros.ColWidth(12) = 1000      'fchbaja
        Grid_Registros.ColWidth(13) = 1500      'supervisornombre
        Grid_Registros.ColWidth(14) = 500       'supervisor
        Grid_Registros.ColWidth(15) = 1500      'actividad
        Grid_Registros.ColWidth(16) = 500       'turno
        Grid_Registros.ColWidth(17) = 0         'tipocontrato
        Grid_Registros.ColWidth(18) = 500       'fecha_nacimiento
        Grid_Registros.ColWidth(19) = 200       'sexo
        Grid_Registros.ColWidth(20) = 800       'rfc
        Grid_Registros.ColWidth(21) = 500       'estado civil
        Grid_Registros.ColWidth(22) = 800       'nss
        Grid_Registros.ColWidth(23) = 0         'domicilio
        Grid_Registros.ColWidth(24) = 0         'colonia
        Grid_Registros.ColWidth(25) = 0         'poblacion
        Grid_Registros.ColWidth(26) = 0         'codpostal
        Grid_Registros.ColWidth(27) = 0         'municipio
        Grid_Registros.ColWidth(28) = 0         'telefono
        Grid_Registros.ColWidth(29) = 0         'nacional
        Grid_Registros.ColWidth(30) = 0         'numeroExterior
        Grid_Registros.ColWidth(31) = 0         'curp
        Grid_Registros.ColWidth(32) = 0         'sueldo
    End If
    iLinea = 13
    Btn_Aplicar.Enabled = True
    iLinea = 14
    Me.MousePointer = 0
Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox "[Frm_Ope_Importacion_Datos|Explorar_Empleados_Nueva_Vista|Linea:" & iLinea & "]" & Err.Description, vbExclamation
        Close #1
    End If
End Sub

Private Sub Explorar_Empleados()

On Error GoTo Fin
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    'Obtener la imagen de una ruta especifica de tipo xls
    Cmd_Archivo.Filter = "Actualizacion Texto (*.txt)|*.txt|Actualizacion en Excel (*.xls)|*.xls"
    Cmd_Archivo.FilterIndex = 1
    Cmd_Archivo.DialogTitle = "Seleccione el archivo"
    If Ruta_Archivo_Pedido <> "" Then
        Cmd_Archivo.InitDir = Ruta_Archivo_Pedido
        Cmd_Archivo.ShowSave
    Else
        Cmd_Archivo.ShowSave
    End If
    'Destino del archivo .txt
    Me.MousePointer = 11
    Txt_Ruta.Text = Cmd_Archivo.FileName
    Txt_Archivo.Text = Cmd_Archivo.FileTitle
    Grid_Registros.Rows = 0
    Grid_Registros.Cols = 18
    'Abre el archivo de texto
    Open Cmd_Archivo.FileName For Input As #1
        If Not EOF(1) Then
            Do While Not EOF(1)
                'Captura el valor de linea
                Line Input #1, linea
                Linea_Datos = CStr(Trim(linea))
                If Linea_Datos <> "" Then
                    Grupos = Split(Linea_Datos, Chr(9))
                    'Llena el grid
                    Grid_Registros.AddItem Trim(Grupos(0)) & Chr(9) & Trim(Grupos(1)) _
                        & Chr(9) & Trim(Grupos(2)) & Chr(9) & Trim(Grupos(3)) _
                        & Chr(9) & Trim(Grupos(4)) & Chr(9) & Trim(Grupos(5)) _
                        & Chr(9) & Trim(Grupos(6)) & Chr(9) & Trim(Grupos(7)) _
                        & Chr(9) & Trim(Grupos(8)) & Chr(9) & Trim(Grupos(9)) _
                        & Chr(9) & Trim(Grupos(10)) & Chr(9) & Trim(Grupos(11)) _
                        & Chr(9) & Trim(Grupos(12)) & Chr(9) & Trim(Grupos(13)) _
                        & Chr(9) & Trim(Grupos(14)) & Chr(9) & Trim(Grupos(15)) _
                        & Chr(9) & Trim(Grupos(16)) & Chr(9) & Trim(Grupos(17))
                End If
            Loop
        End If
    Close #1
    If Grid_Registros.Rows > 1 Then
        Grid_Registros.FixedCols = 1
        Grid_Registros.FixedRows = 1
    End If
    Btn_Aplicar.Enabled = True
    Me.MousePointer = 0
Exit Sub
Fin:
    Me.MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Close #1
    End If
End Sub

Private Sub Btn_Explorar_Click()
    Cmb_Tipo_Producto.Enabled = False
    Select Case Cmb_Tipo_Producto.Text
        Case "SQL"
            Explorar_Empleados_Nueva_Vista
            Cmb_Tipo_Producto.ListIndex = 0
        Case "EXCEL"
            Explorar_Empleados
            Cmb_Tipo_Producto.ListIndex = 1
'        Case "Turnos"
'            Explorar_Turnos
'            Cmb_Tipo_Producto.ListIndex = 1
'        Case "SAP Altas"
'            Explorar_SAP_Altas
'            Cmb_Tipo_Producto.ListIndex = 2
'        Case "SAP Bajas"
'            Explorar_SAP_Bajas
'            Cmb_Tipo_Producto.ListIndex = 3
    End Select
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Btn_Salir_Importacion_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Cmb_Tipo_Producto.ListIndex = 0
End Sub

