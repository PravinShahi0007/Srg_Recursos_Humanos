VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_Cat_Areas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CATALOGO DE AREAS"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   8025
   Begin VB.CommandButton Btn_Modificar 
      Caption         =   "Modificar"
      Height          =   555
      Left            =   1770
      Picture         =   "Frm_Cat_Departamentos_Supervisor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "M"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   555
      Left            =   3370
      Picture         =   "Frm_Cat_Departamentos_Supervisor.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "B"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   6530
      Picture         =   "Frm_Cat_Departamentos_Supervisor.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Nuevo 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   120
      Picture         =   "Frm_Cat_Departamentos_Supervisor.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "A"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton Btn_Buscar 
      Caption         =   "Buscar"
      Height          =   555
      Left            =   4920
      Picture         =   "Frm_Cat_Departamentos_Supervisor.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "C"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.PictureBox Pic_Tipos_Notas_Credito 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   0
      ScaleHeight     =   6285
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      Begin VB.Frame Fra_Areas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Areas"
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
         Left            =   75
         TabIndex        =   1
         Top             =   3525
         Width           =   7860
         Begin MSFlexGridLib.MSFlexGrid Grid_Areas 
            Height          =   2400
            Left            =   75
            TabIndex        =   2
            Top             =   240
            Width           =   7725
            _ExtentX        =   13626
            _ExtentY        =   4233
            _Version        =   393216
            Rows            =   0
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin TabDlg.SSTab Tab_Cat_Areas 
         Height          =   3000
         Left            =   75
         TabIndex        =   9
         Top             =   480
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   5292
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Generales"
         TabPicture(0)   =   "Frm_Cat_Departamentos_Supervisor.frx":050A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra_Generales_areas"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Empleados Areas"
         TabPicture(1)   =   "Frm_Cat_Departamentos_Supervisor.frx":0526
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fra_Cat_Empleados"
         Tab(1).Control(1)=   "Cmb_Cat_Empleados"
         Tab(1).ControlCount=   2
         Begin VB.Frame Fra_Generales_areas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Generales"
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
            Height          =   2570
            Left            =   60
            TabIndex        =   16
            Top             =   360
            Width           =   7740
            Begin VB.TextBox Txt_Nombre 
               Height          =   315
               Left            =   1335
               MaxLength       =   50
               TabIndex        =   20
               Top             =   660
               Width           =   6255
            End
            Begin VB.TextBox Txt_Area_ID 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   270
               Width           =   2370
            End
            Begin VB.ComboBox Cmb_Cat_Empleados_Departamento 
               Height          =   315
               ItemData        =   "Frm_Cat_Departamentos_Supervisor.frx":0542
               Left            =   1335
               List            =   "Frm_Cat_Departamentos_Supervisor.frx":0544
               TabIndex        =   18
               Top             =   1010
               Width           =   6255
            End
            Begin VB.TextBox Txt_Observaciones 
               Height          =   675
               Left            =   1320
               MaxLength       =   1000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   1335
               Width           =   6255
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Comentarios"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   75
               TabIndex        =   24
               Top             =   1470
               Width           =   750
            End
            Begin VB.Label Lbl_Descripcion_Tiempos 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   23
               Top             =   720
               Width           =   660
            End
            Begin VB.Label Lbl_Area_ID 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Area ID"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   75
               TabIndex        =   22
               Top             =   330
               Width           =   540
            End
            Begin VB.Label Lbl_Departamentos 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Departamento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   75
               TabIndex        =   21
               Top             =   1050
               Width           =   1200
            End
         End
         Begin VB.ComboBox Cmb_Cat_Empleados 
            Height          =   315
            Left            =   -73920
            TabIndex        =   14
            Top             =   550
            Width           =   4980
         End
         Begin VB.Frame Fra_Cat_Empleados 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2570
            Left            =   -74940
            TabIndex        =   10
            Top             =   360
            Width           =   7740
            Begin VB.CommandButton Btn_Eliminar_Epleados 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   6840
               TabIndex        =   12
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton Btn_Agregar_Empleados 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   5940
               TabIndex        =   11
               Top             =   180
               Width           =   855
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Empleados_Areas 
               Height          =   1920
               Left            =   45
               TabIndex        =   13
               Top             =   600
               Width           =   7635
               _ExtentX        =   13467
               _ExtentY        =   3387
               _Version        =   393216
               Rows            =   0
               Cols            =   3
               FixedRows       =   0
               BackColorBkg    =   16777215
               Appearance      =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Empleado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   840
            End
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "AREAS"
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
         Left            =   3420
         TabIndex        =   8
         Top             =   0
         Width           =   1305
      End
   End
End
Attribute VB_Name = "Frm_Cat_Areas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Agregar_Empleados_Click()
Dim Cont_Fila As Integer
    'Valida los datos del dependiente
    If Cmb_Cat_Empleados.ListIndex > -1 Then
        'Agrega el dependiente a la lista
        'Busca si el equipo ya ha sido agregado
        For Cont_Fila = 1 To Grid_Empleados_Areas.Rows - 1
            If Format(Cmb_Cat_Empleados.ItemData(Cmb_Cat_Empleados.ListIndex), "00000") = _
              Trim(Grid_Empleados_Areas.TextMatrix(Cont_Fila, 0)) Then
                MsgBox "El empleado ya se agregó", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
        Next
        Grid_Empleados_Areas.Cols = 2
        If Grid_Empleados_Areas.Rows = 0 Then
            Grid_Empleados_Areas.AddItem "Empleado_ID" & Chr(9) & "Empleado"
            Grid_Empleados_Areas.ColWidth(0) = 0    'Equipo_ID
            Grid_Empleados_Areas.ColWidth(1) = 3500 'Equipo
            Grid_Empleados_Areas.ColAlignment(1) = flexAlignLeftCenter
        End If
        Grid_Empleados_Areas.AddItem Format(Cmb_Cat_Empleados.ItemData(Cmb_Cat_Empleados.ListIndex), "00000") & Chr(9) & _
            Trim(Cmb_Cat_Empleados.Text)
        Cmb_Cat_Empleados.ListIndex = -1
        Grid_Empleados_Areas.FixedRows = 1
    End If

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
    
    Nombre = InputBox("Proporcione el No. area, Nombre para buscar la AREA")
    Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
    
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Consulta_Cat_Areas(Nombre)
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Areas
'DESCRIPCIÓN:           Consulta las areas y los muestra en el grid
'PARÁMETROS :           Nombre: Indica el nombre de la area
'CREO       :           Yazmin Flores Ramirez
'FECHA_CREO :           14-Noviembre-2014
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Areas(Nombre As String)
Dim Rs_Consulta_Cat_Areas As rdoResultset       'Informacion de los registros
    
    Grid_Areas.Rows = 0
    
    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT Area_ID, Cat_Areas.Nombre, Cat_Departamentos.Nombre as Departameto"
    Mi_SQL = Mi_SQL & " FROM Cat_Areas"
    Mi_SQL = Mi_SQL & " inner join  Cat_Departamentos on Cat_Departamentos.Departamento_ID=Cat_Areas.Departamento_ID"
    Mi_SQL = Mi_SQL & " WHERE Cat_Areas.Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Areas.Nombre"
    Set Rs_Consulta_Cat_Areas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    
    With Rs_Consulta_Cat_Areas
        If Not .EOF Then
            
            Grid_Areas.AddItem "Area ID" & Chr(9) & "Nombre" & Chr(9) & "Departamento"
            While Not .EOF
                Grid_Areas.AddItem .rdoColumns("Area_ID") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Departameto")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del grid_usuarios
            Grid_Areas.FixedRows = 1
            Grid_Areas.ColWidth(0) = 0      'Area_ID
            Grid_Areas.ColWidth(1) = 6000   'Nombre
            Grid_Areas.ColWidth(2) = 1800   'Departamento
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Areas = Nothing
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo HANDLER
    If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
        Conexion_Base.BeginTrans
        If Trim(Txt_Area_ID.Text) <> "" Then
            If Conectar_Ayudante.Elimina_Catalogo("Cat_Areas_Detalles", "Area_ID", Trim(Txt_Area_ID.Text)) = True Then
               If Conectar_Ayudante.Elimina_Catalogo("Cat_Areas", "Area_ID", Trim(Txt_Area_ID.Text)) = True Then
                   If Grid_Areas.Rows = 2 Then
                       Grid_Areas.Rows = 0
                   Else
                       Grid_Areas.RemoveItem Grid_Areas.RowSel
                   End If
                   Call Conectar_Ayudante.Limpiar_Textos(Me)
                   MsgBox "Area eliminada", vbInformation + vbOKOnly, Me.Caption
               Else
                   MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
               End If
            Else
                   MsgBox "No se pudo eliminar el registro", vbExclamation + vbOKOnly, Me.Caption
            End If
        Else
            MsgBox "Seleccione una area para poder eliminar", vbInformation + vbOKOnly, Me.Caption
        End If
    Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia los textos de la forma
    Conexion_Base.CommitTrans
    End If
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Eliminar_Epleados_Click()
    If Grid_Empleados_Areas.Rows > 0 Then
        If Grid_Empleados_Areas.Rows = 2 Then
            Grid_Empleados_Areas.Rows = 0
        Else
            Grid_Empleados_Areas.RemoveItem Grid_Empleados_Areas.RowSel
        End If
    End If
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN:  Modifica_Cat_Areas
'DESCRIPCIÓN:           Modifica el registro de la area
'PARÁMETROS :
'CREO       :           Yazmin Flores Ramirez
'FECHA_CREO        :    14-Noviembre-2014
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modifica_Cat_Areas()
Dim Rs_Modificacion_Cat_Areas As rdoResultset 'Informacion del registro
Dim Rs_Alta_Cat_Areas_Detalles As rdoResultset
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Areas"
    Mi_SQL = Mi_SQL & " WHERE Area_ID ='" & Trim(Txt_Area_ID.Text) & "'"
    Set Rs_Modificacion_Cat_Areas = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    With Rs_Modificacion_Cat_Areas
        .Edit
            .rdoColumns("Departamento_ID") = Format(Cmb_Cat_Empleados_Departamento.ItemData(Cmb_Cat_Empleados_Departamento.ListIndex), "00000")
            .rdoColumns("Nombre") = Trim(Txt_Nombre.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Observaciones.Text))
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
    Set Rs_Modificacion_Cat_Areas = Nothing
    'Agrega los checadores
    Mi_SQL = "DELETE Cat_Areas_Detalles WHERE Area_ID = '" & Trim(Txt_Area_ID.Text) & "'"
    Conexion_Base.Execute Mi_SQL
    Set Rs_Alta_Cat_Areas_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Areas_Detalles")
    With Rs_Alta_Cat_Areas_Detalles
        For Cont_Fila = 1 To Grid_Empleados_Areas.Rows - 1
            .AddNew
                .rdoColumns("Area_ID") = Trim(Txt_Area_ID.Text)
                .rdoColumns("Empleado_ID") = Trim(Grid_Empleados_Areas.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Areas_Detalles = Nothing
    
    With Grid_Areas
        .TextMatrix(.RowSel, 1) = Trim(Txt_Nombre.Text)
        .TextMatrix(.RowSel, 2) = Trim(Cmb_Cat_Empleados_Departamento.Text)
    End With
    'Deshabilita y habilita los controles de la forma para no dejar introducir nuevos valores
    Fra_Generales_areas.Enabled = False
    Fra_Areas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Modificar.Caption = "Modificar"
    Btn_Buscar.Enabled = True
    Btn_Nuevo.Enabled = True
    Btn_Eliminar.Enabled = True
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Areas", Me)
    MsgBox "La Area ha sido modificada", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN:  Alta_Cat_Areas
'DESCRIPCIÓN:           Da de alta un registro en Cat_Areas
'PARÁMETROS :
'CREO       :           Yazmin Flores Ramirez
'FECHA_CREO :           14 Noviembre 2014
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Cat_Areas()
Dim Rs_Alta_Cat_Areas As rdoResultset            'Informacion del registro
Dim Rs_Alta_Cat_Areas_Detalles As rdoResultset 'Informacion de los checadores
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    Set Rs_Alta_Cat_Areas = Conectar_Ayudante.Recordset_Agregar("Cat_Areas")
    'Agrega el reigstro del Empresa
    With Rs_Alta_Cat_Areas
        .AddNew
            Txt_Area_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Areas", "Area_ID"), "00000")
            .rdoColumns("Area_ID") = Trim(Txt_Area_ID.Text)
            .rdoColumns("Departamento_ID") = Format(Cmb_Cat_Empleados_Departamento.ItemData(Cmb_Cat_Empleados_Departamento.ListIndex), "00000")
            .rdoColumns("Nombre") = Trim(Txt_Nombre.Text)
            .rdoColumns("Comentarios") = Trim(UCase(Txt_Observaciones.Text))
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
        .Close
    End With
    Set Rs_Alta_Cat_Areas = Nothing
    'Agrega los checadores
    Set Rs_Alta_Cat_Areas_Detalles = Conectar_Ayudante.Recordset_Agregar("Cat_Areas_Detalles")
    With Rs_Alta_Cat_Areas_Detalles
        For Cont_Fila = 1 To Grid_Empleados_Areas.Rows - 1
            .AddNew
                .rdoColumns("Area_ID") = Trim(Txt_Area_ID.Text)
                .rdoColumns("Empleado_ID") = Trim(Grid_Empleados_Areas.TextMatrix(Cont_Fila, 0))
                .rdoColumns("Usuario_Creo") = Nombre_Usuario
                .rdoColumns("Fecha_Creo") = Now
            .Update
        Next
    End With
    Set Rs_Alta_Cat_Areas_Detalles = Nothing
    'Habilita y deshabilita los controles de la forma para que el usuario no pueda introducir
    'o modificar los valoes
    
    Fra_Generales_areas.Enabled = False
    Fra_Areas.Enabled = True
    Btn_Salir.Caption = "Salir"
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Buscar.Enabled = True
    Btn_Modificar.Enabled = True
    Btn_Eliminar.Enabled = True
    'Pone un encabezado en el grid
    With Grid_Areas
        If .Rows = 0 Then
            .AddItem "Area ID" & Chr(9) & "Nombre" & Chr(9) & "Departamento"
        End If
        'Llena el grid con los datos del nuevo usuario
        .AddItem Trim(Txt_Area_ID.Text) & Chr(9) & Trim(Txt_Nombre.Text) & Chr(9) & Trim(Cmb_Cat_Empleados_Departamento.Text)
        
        'Configura el tamaño de las columnas del grid_usuarios
        .FixedRows = 1
        .ColWidth(0) = 0      'Area_ID
        .ColWidth(1) = 6000   'Nombre
        .ColWidth(2) = 1800   'Departamento

    End With
    Conexion_Base.CommitTrans
    Call Conectar_Ayudante.Limpiar_Textos(Me)
    Call Conectar_Ayudante.Validacion_Accesos_Sistema("SubMenu_Cat_Areas", Me)
    MsgBox "Empresa dada de alta", vbOKOnly + vbInformation, Me.Caption
    Exit Sub

'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Btn_Modificar_Click()
    If Btn_Modificar.Caption = "Modificar" Then
    'Revisa que exista un registro a modificar y prepara la interfaz
    If Trim(Txt_Area_ID.Text) <> "" Then
        Fra_Generales_areas.Enabled = True
        Fra_Areas.Enabled = False
        Txt_Nombre.SetFocus
        On Error Resume Next
        SendKeys "{Home}+{End}"
    Else
        MsgBox "Seleccione una area para poder modificar", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    Btn_Modificar.Caption = "Actualizar"
    Btn_Eliminar.Enabled = False
    Btn_Nuevo.Enabled = False
    Btn_Buscar.Enabled = False
    Btn_Salir.Caption = "Regresar"
    Else
        If Trim(Txt_Nombre.Text) <> "" Then
            If Grid_Empleados_Areas.Rows > 0 Then
                Modifica_Cat_Areas
                Cmb_Cat_Empleados_Departamento.ListIndex = -1
                Tab_Cat_Areas.Tab = 0
            Else
                MsgBox "Ingrese por lo menos un empleado", vbOKOnly + vbInformation, Me.Caption
                Cmb_Cat_Empleados.SetFocus
                Tab_Cat_Areas.Tab = 1
            End If
        Else
            MsgBox "Ingrese el Nombre de la area", vbOKOnly + vbInformation, Me.Caption
            Txt_Nombre.SetFocus
        End If
    End If
End Sub

Private Sub Btn_Nuevo_Click()
Dim Catacter As String 'Indica el caractere que se desea comparar
    
    If Btn_Nuevo.Caption = "Nuevo" Then
        Btn_Nuevo.Caption = "Dar de Alta"
        Btn_Modificar.Enabled = False
        Btn_Eliminar.Enabled = False
        Btn_Buscar.Enabled = False
        Btn_Salir.Caption = "Regresar"
        Call Conectar_Ayudante.Limpiar_Textos(Me) 'Limpia las cajas de texto
        
        Txt_Area_ID.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Areas", "Area_ID"), "00000")
        Fra_Generales_areas.Enabled = True
        Fra_Areas.Enabled = False
        'Llena los checadores
        Cmb_Cat_Empleados_Departamento.Text = ""
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Cat_Empleados_Departamento, 1, "Nombre")
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados ", Cmb_Cat_Empleados, 1, "Apellido_Paterno", "AND Tipo='S' AND Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%' OR " & "Apellido_Paterno LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%')", False, "")
            
        Cmb_Cat_Empleados.ListIndex = -1
        Cmb_Cat_Empleados_Departamento.ListIndex = -1
        Grid_Empleados_Areas.Rows = 0
        Tab_Cat_Areas.Tab = 0
        Txt_Nombre.SetFocus
    Else
        'Valida la informacion obligatoria
        If Trim(Txt_Nombre.Text) <> "" Then
            If Grid_Empleados_Areas.Rows > 0 Then
                Alta_Cat_Areas
                Cmb_Cat_Empleados_Departamento.ListIndex = -1
                Tab_Cat_Areas.Tab = 0
            Else
                MsgBox "Ingrese por lo menos un empleado", vbOKOnly + vbInformation, Me.Caption
                Cmb_Cat_Empleados.SetFocus
                Tab_Cat_Areas.Tab = 1
            End If
        Else
            MsgBox "Ingrese el Nombre de la area", vbOKOnly + vbInformation, Me.Caption
            Txt_Nombre.SetFocus
        End If
    End If
End Sub

Private Sub Btn_Salir_Click()
    If Btn_Salir.Caption = "Salir" Then
        Unload Me
    Else
        Call Conectar_Ayudante.Limpiar_Textos(Me)
        Btn_Nuevo.Enabled = True
        Btn_Modificar.Enabled = True
        Btn_Eliminar.Enabled = True
        Btn_Buscar.Enabled = True
        Btn_Modificar.Caption = "Modificar"
        Btn_Nuevo.Caption = "Nuevo"
        Btn_Salir.Caption = "Salir"
        
        Fra_Generales_areas.Enabled = False
        Fra_Areas.Enabled = True
        Call Conectar_Ayudante.Validacion_Accesos_Sistema("Submenu_Cat_Areas", Me)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Departamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Cat_Empleados_Departamento, 1, "Nombre")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Cat_Empleados_Departamento_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Cat_Empleados_Departamento, KeyCode)
End Sub

Private Sub Cmb_Cat_Empleados_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Empleado_ID, (Apellido_Paterno+' '+Apellido_Materno+' '+Nombre) AS Nombre", "Cat_Empleados WHERE Estatus='A' AND (Nombre LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%' OR " & "Apellido_Paterno LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%' OR Apellido_Materno LIKE '%" & Trim(Cmb_Cat_Empleados.Text) & "%' " & IIf(IsNumeric(Cmb_Cat_Empleados.Text), " OR No_Tarjeta = " & Trim(Cmb_Cat_Empleados.Text), "") & ")", Cmb_Cat_Empleados, 0, "Apellido_Paterno", "", False, "")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7590
    Me.Width = 8265
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Grid_Areas_Click()
Dim Rs_Consulta_Cat_Areas As rdoResultset    'Informacion de la empresa
Dim Rs_Consulta_Cat_Areas_Detalles As rdoResultset    'Informacion de la empresa
With Grid_Areas
    If .Rows > 1 Then
        Mi_SQL = "SELECT * FROM Cat_Areas"
        Mi_SQL = Mi_SQL & " WHERE Area_ID ='" & Trim(.TextMatrix(.RowSel, 0)) & "'"
        Set Rs_Consulta_Cat_Areas = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Cat_Areas
            If Not .EOF Then
            
                Txt_Area_ID.Text = .rdoColumns("Area_ID")
                Txt_Nombre.Text = .rdoColumns("Nombre")
                If Not IsNull(.rdoColumns("Comentarios")) Then Txt_Observaciones.Text = .rdoColumns("Comentarios")
                Cmb_Cat_Empleados_Departamento.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Departamento_ID")), Cmb_Cat_Empleados_Departamento)
                If Not IsNull(.rdoColumns("Departamento_ID")) Then
                    Call Conectar_Ayudante.Llena_Combo_Item("Departamento_ID, Nombre", "Cat_Departamentos", Cmb_Cat_Empleados_Departamento, 1, "Nombre")
                    Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Cat_Empleados_Departamento)
                Else
                    Cmb_Cat_Empleados_Departamento.ListIndex = -1
                End If
                
                Call Conectar_Ayudante.Asigna_Item_Combo(.rdoColumns("Departamento_ID"), Cmb_Cat_Empleados_Departamento)
                Grid_Empleados_Areas.Rows = 0
                'Llena los checadores de la empresa
                Mi_SQL = "SELECT Cat_Areas_Detalles.Area_ID, Cat_Empleados.Empleado_ID, (Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre) as Empleado"
                Mi_SQL = Mi_SQL & " FROM Cat_Areas_Detalles, Cat_Empleados"
                Mi_SQL = Mi_SQL & " WHERE Cat_Areas_Detalles.Empleado_ID = Cat_Empleados.Empleado_ID"
                Mi_SQL = Mi_SQL & " AND Cat_Areas_Detalles.Area_ID = '" & .rdoColumns("Area_ID") & "'"
                Set Rs_Consulta_Cat_Areas_Detalles = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
                    If Not Rs_Consulta_Cat_Areas_Detalles.EOF Then
                        Grid_Empleados_Areas.Cols = 2
                        If Grid_Empleados_Areas.Rows = 0 Then
                            Grid_Empleados_Areas.AddItem "Empleado_ID" & Chr(9) & "Empleado"
                            Grid_Empleados_Areas.ColWidth(0) = 0    'Empleado_ID
                            Grid_Empleados_Areas.ColWidth(1) = 3500 'Empleado
                            Grid_Empleados_Areas.ColAlignment(1) = flexAlignLeftCenter
                        End If
                        While Not Rs_Consulta_Cat_Areas_Detalles.EOF
                                Grid_Empleados_Areas.AddItem Rs_Consulta_Cat_Areas_Detalles.rdoColumns("Empleado_ID") & Chr(9) & _
                                    Rs_Consulta_Cat_Areas_Detalles.rdoColumns("Empleado")
                            Rs_Consulta_Cat_Areas_Detalles.MoveNext
                        Wend
                        Grid_Empleados_Areas.FixedRows = 1
                    End If
                Set Rs_Consulta_Cat_Areas_Detalles = Nothing
                .Close
            End If
            Set Rs_Consulta_Cat_Areas = Nothing
        End With
    End If
End With
End Sub

Private Sub Grid_Areas_EnterCell()
    Grid_Areas_Click
End Sub


Private Sub Txt_Nombre_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub

Private Sub Txt_Observaciones_KeyPress(KeyAscii As Integer)
    Call Conectar_Ayudante.Quitar_Caracteres_Raros(KeyAscii, True)
End Sub
