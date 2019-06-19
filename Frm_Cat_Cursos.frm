VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frm_Cat_Cursos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATALOGOS"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5891.195
   ScaleMode       =   0  'User
   ScaleWidth      =   4277.287
   Begin VB.PictureBox Pic_Cat_Empleados 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   0
      Picture         =   "Frm_Cat_Cursos.frx":0000
      ScaleHeight     =   9375
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.Frame Fra_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cursos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2880
         Left            =   120
         TabIndex        =   26
         Top             =   5520
         Width           =   7305
         Begin MSFlexGridLib.MSFlexGrid Grid_Cat_Cursos 
            Height          =   2520
            Left            =   60
            TabIndex        =   13
            Top             =   240
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   4445
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
      Begin VB.Frame Fra_Generales_Cat_Cursos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generales Cursos"
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
         Height          =   5055
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   7290
         Begin VB.TextBox Txt_Cat_Cursos_Clave 
            Height          =   315
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1320
            Width           =   2370
         End
         Begin VB.ComboBox Cmb_Cat_Cursos_Tipo_Curso 
            Height          =   315
            ItemData        =   "Frm_Cat_Cursos.frx":0C42
            Left            =   1125
            List            =   "Frm_Cat_Cursos.frx":0C4C
            TabIndex        =   6
            Top             =   1320
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Cursos_Temario 
            Height          =   555
            Left            =   1125
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   4440
            Width           =   5925
         End
         Begin VB.TextBox Txt_Cat_Cursos_Lista_Material 
            Height          =   675
            Left            =   1125
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   3720
            Width           =   5925
         End
         Begin VB.TextBox Txt_Cat_Cursos_Req_Minimos 
            Height          =   675
            Left            =   1125
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   3000
            Width           =   5925
         End
         Begin VB.TextBox Txt_Cat_Cursos_Objetivo 
            Height          =   675
            Left            =   1125
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   2280
            Width           =   5925
         End
         Begin VB.CheckBox Chk_Cat_Cursos_Auditable 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Auditable"
            Height          =   315
            Left            =   3600
            TabIndex        =   5
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox Txt_Cat_Cursos_Nombre 
            Height          =   315
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   5925
         End
         Begin VB.TextBox Txt_Cat_Cursos_Curso_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Cursos_Total_Horas 
            Height          =   315
            Left            =   1125
            MaxLength       =   5
            TabIndex        =   4
            Top             =   960
            Width           =   2370
         End
         Begin VB.TextBox Txt_Cat_Cursos_Descriopcion 
            Height          =   555
            Left            =   1125
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1680
            Width           =   5925
         End
         Begin VB.ComboBox Cmb_Cat_Cursos_Estatus 
            Height          =   315
            ItemData        =   "Frm_Cat_Cursos.frx":0C62
            Left            =   4680
            List            =   "Frm_Cat_Cursos.frx":0C6C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2370
         End
         Begin VB.Label Lbl_Clave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Clave"
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
            Index           =   2
            Left            =   3600
            TabIndex        =   32
            Top             =   1410
            Width           =   570
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Temario"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   31
            Top             =   4530
            Width           =   570
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Lista Material"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   30
            Top             =   3810
            Width           =   930
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Req. Minimos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   29
            Top             =   3090
            Width           =   960
         End
         Begin VB.Label Lbl_Objetivo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Objetivo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   2370
            Width           =   930
         End
         Begin VB.Label Lbl_Nombre 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Nombre"
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
            Left            =   75
            TabIndex        =   25
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Lbl_Tipo_Nota_Credito_ID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Curso ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   24
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Lbl_Tipo_Curso_id 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Tipo Curso"
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
            Left            =   75
            TabIndex        =   23
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label Lbl_Total_Horas 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Total Hrs"
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
            Left            =   75
            TabIndex        =   22
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Lbl_Descripcion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   1770
            Width           =   840
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "*Estatus"
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
            Left            =   3600
            TabIndex        =   20
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   4440
         Picture         =   "Frm_Cat_Cursos.frx":0C82
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "C"
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Nuevo 
         Caption         =   "Nuevo"
         Height          =   555
         Left            =   120
         Picture         =   "Frm_Cat_Cursos.frx":0D84
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "A"
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Salir 
         Caption         =   "Salir"
         Height          =   555
         Left            =   5880
         Picture         =   "Frm_Cat_Cursos.frx":0E86
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Eliminar 
         Caption         =   "Eliminar"
         Height          =   555
         Left            =   3000
         Picture         =   "Frm_Cat_Cursos.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "B"
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.CommandButton Btn_Modificar 
         Caption         =   "Modificar"
         Height          =   555
         Left            =   1560
         Picture         =   "Frm_Cat_Cursos.frx":108A
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "M"
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1350
      End
      Begin VB.Label Lbl_Cursos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CURSOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   7185
      End
   End
End
Attribute VB_Name = "Frm_Cat_Cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicializa()
Consulta_Cat_Cursos ""
 Call Conectar_Ayudante.Llena_Combo_Item("Tipo_Curso_Id, Nombre", "Cat_Tipos_Cursos", Cmb_Cat_Cursos_Tipo_Curso, 0, "Tipo_Curso_Id", "", False, "")
' Call Conectar_Ayudante.Llena_Combo_Item("Instructor_Id, Nombre", "Cat_Instructores", Cmb_Cat_Cursos_Instructor, 0, "Instructor_Id", "", False, "")
End Sub

'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Consulta_Cat_Cursos
    'DESCRIPCIÓN:           Consulta los Cursos y los muestra en el grid
    'PARÁMETROS :           Nombre: Indica el nombre de la Institución
    'CREO       :           Ana Laura Huichapa Ramírez
    'FECHA_CREO :           23 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Consulta_Cat_Cursos(Nombre As String)
Dim Rs_Consulta_Cat_Cursos As rdoResultset       'Informacion de los registros

    Grid_Cat_Cursos.Rows = 0

    'Consulta los datos generales del usuario
    Mi_SQL = "SELECT *"
    Mi_SQL = Mi_SQL & " FROM Cat_Cursos_Capacitaciones"
    Mi_SQL = Mi_SQL & " WHERE Nombre LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " OR Clave LIKE '%" & Nombre & "%'"
    Mi_SQL = Mi_SQL & " ORDER BY Nombre"
'    MsgBox Mi_SQL

    Set Rs_Consulta_Cat_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)

    With Rs_Consulta_Cat_Cursos
        If Not .EOF Then

            Grid_Cat_Cursos.AddItem "Curso ID" & Chr(9) & "Clave" & Chr(9) & "Nombre" & Chr(9) & "Estatus"
            While Not .EOF
                Grid_Cat_Cursos.AddItem .rdoColumns("Curso_ID") & Chr(9) & .rdoColumns("Clave") & Chr(9) & .rdoColumns("Nombre") & Chr(9) & .rdoColumns("Estatus")
                .MoveNext
            Wend
            'Configura el tamaño de las columnas del Grid_Cat_Instituciones
            Grid_Cat_Cursos.FixedRows = 1
            Grid_Cat_Cursos.ColWidth(0) = 800     'Intitución_ID
            Grid_Cat_Cursos.ColWidth(1) = 1000   'clave
            Grid_Cat_Cursos.ColWidth(2) = 4000   'nombre
            Grid_Cat_Cursos.ColWidth(3) = 1000   'estatus
            .Close
        End If
    End With
    'Cierra el manejador del registro
    Set Rs_Consulta_Cat_Cursos = Nothing

End Sub

Private Sub Btn_Buscar_Click()
Dim Nombre As String 'Obtiene el nombre a consultar
        Nombre = InputBox("Proporcione el Nombre o Clave para buscar los cursos")
        Nombre = Conectar_Ayudante.Quitar_Caracter(Nombre, "'")
        Consulta_Cat_Cursos Nombre
       
End Sub

Private Sub Btn_Eliminar_Click()
On Error GoTo Fin
    If Txt_Cat_Cursos_Curso_Id.Text <> "" Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            Mi_SQL = "DELETE FROM Cat_Cursos_Capacitaciones WHERE Curso_ID='" & Trim(Txt_Cat_Cursos_Curso_Id.Text) & "'"
            Conexion_Base.Execute Mi_SQL
            'Quita los datos del usuario contenidos en el Grid
            If Grid_Cat_Cursos.Rows = 2 Then
                Grid_Cat_Cursos.Rows = 0
            Else
                Grid_Cat_Cursos.RemoveItem Grid_Cat_Cursos.RowSel
            End If 'Grid_productos
            MsgBox "Curso Eliminado", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        MsgBox ("Es necesario seleccionar un registro para eliminar")
    End If
Exit Sub
Fin:
    If Err.Number <> 0 Then
        Conexion_Base.RollbackTrans
        If Err.Number = 40002 Then
            MsgBox "NO SE ELIMINO EL REGISTRO. El registro tiene dependencia en otro catalogo, elimine la relacion y vuelva a intentarlo.", vbExclamation, "Mensaje"
        Else
            MsgBox Err.Description, vbExclamation
        End If
    End If
End Sub

Private Sub Btn_Modificar_Click()
If Btn_Modificar.Caption = "Modificar" Then
    If Txt_Cat_Cursos_Curso_Id.Text <> "" Then
        Call Configurar_Formulario(True)
        Btn_Modificar.Enabled = True
        Btn_Modificar.Caption = "Guardar"
        Cmb_Cat_Cursos_Estatus.SetFocus
        Fra_Cursos.Enabled = False
    Else
        MsgBox ("Es necesario seleccionar un registro para modificar")
    End If
Else
    Modificar_Cat_Cursos
    Limpiar_Formulario
    Btn_Modificar.Caption = "Modificar"
    Configurar_Formulario (False)
    Btn_Salir.Caption = "Salir"
    Fra_Cursos.Enabled = True
End If
End Sub

Private Sub Btn_Nuevo_Click()
If Btn_Nuevo.Caption = "Nuevo" Then
    Call Configurar_Formulario(True)
    Limpiar_Formulario
    Btn_Nuevo.Enabled = True
    Btn_Nuevo.Caption = "Guardar"
    Cmb_Cat_Cursos_Estatus.SetFocus
    Fra_Cursos.Enabled = False
Else
    If Validar_Componentes Then

        Call Alta_Curso
        Limpiar_Formulario
        Btn_Nuevo.Caption = "Nuevo"
        Configurar_Formulario (False)
        Btn_Salir.Caption = "Salir"
        Fra_Cursos.Enabled = True
    Else
        MsgBox ("Todos los campos marcados con * son necesarios")
    End If
End If



End Sub

Private Sub Configurar_Formulario(ByVal Habilitar As Boolean)
Fra_Generales_Cat_Cursos.Enabled = Habilitar
Btn_Nuevo.Enabled = Not Habilitar
Btn_Modificar.Enabled = Not Habilitar
Btn_Eliminar.Enabled = Not Habilitar
Btn_Buscar.Enabled = Not Habilitar
Btn_Salir.Caption = "Cancelar"

End Sub
Function Validar_Componentes() As Boolean
Validar_Componentes = True
If Txt_Cat_Cursos_Clave.Text = "" Then
Validar_Componentes = False
End If
If Txt_Cat_Cursos_Nombre.Text = "" Then
Validar_Componentes = False
End If
If Txt_Cat_Cursos_Total_Horas.Text = "" Then
Validar_Componentes = False
End If
If Cmb_Cat_Cursos_Tipo_Curso.ListIndex = -1 Then
Validar_Componentes = False
End If
If Cmb_Cat_Cursos_Estatus.ListIndex = -1 Then
Validar_Componentes = False
End If
'If Cmb_Cat_Cursos_Instructor.ListIndex = -1 Then
'Validar_Componentes = False
'End If
'If Txt_Cat_Cursos_Tipo.Text = "" Then
'Validar_Componentes = False
'End If
'If Txt_Cat_Cursos_Clave_SAP.Text = "" Then
'Validar_Componentes = False
'End If

End Function


'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN: Alta_Curso
    'DESCRIPCIÓN: Da de alta un nuevo registro con los datos del curso que el usuario
    '             asigno, así como da de alta los menu y submenus a los cuales
    '             puede accesar el usuario
    'PARÁMETROS :
    'CREO       : Ana Laua Huichapa Ramírez
    'FECHA_CREO : 28-Diciembre-2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Alta_Curso()
'Dim Menus As Integer                                'Contador que sirve para ver en que posición me encuentro en el grid
Dim Rs_Alta_Cat_Cursos As rdoResultset            'Manejo del registro de Cat_Instituciones, da de alta la institución
'Dim Rs_Seguridad_Seguridad_Sistema As rdoResultset  'Manejo de registro de Seguridad_Sistema, guarda a que menus son lo que va a tener acceso el usuario
Dim Ctl As Control
Set Conectar_Ayudante = New Ayudante

On Error GoTo HANDLER
    Conexion_Base.BeginTrans
'    Conexion_Servidor.BeginTrans
    
    'Alta de Institución
    Set Rs_Alta_Cat_Cursos = Conectar_Ayudante.Recordset_Agregar("Cat_Cursos_Capacitaciones")
    'Llena la tabla de Cat_Instituciones con los datos contenidos en las cajas de textos
    With Rs_Alta_Cat_Cursos
        .AddNew
            Txt_Cat_Cursos_Curso_Id.Text = Format(Conectar_Ayudante.Maximo_Catalogo("Cat_Cursos_Capacitaciones", "Curso_Id"), "00000")
            .rdoColumns("Curso_ID") = Txt_Cat_Cursos_Curso_Id.Text
            .rdoColumns("Nombre") = UCase(Txt_Cat_Cursos_Nombre.Text)
            .rdoColumns("Horas") = Val(Txt_Cat_Cursos_Total_Horas.Text)
'            .rdoColumns("Tipo") = UCase(Txt_Cat_Cursos_Tipo.Text)
'            .rdoColumns("Instructor") = Format(Cmb_Cat_Cursos_Instructor.ItemData(Cmb_Cat_Cursos_Instructor.ListIndex), "00000")
            .rdoColumns("Descripcion") = UCase(Txt_Cat_Cursos_Descriopcion.Text)
'            .rdoColumns("Clave_SAP") = Txt_Cat_Cursos_Clave_SAP.Text
            .rdoColumns("Clave") = Trim(Txt_Cat_Cursos_Clave.Text)
            .rdoColumns("Objetivo") = UCase(Txt_Cat_Cursos_Objetivo.Text)
            .rdoColumns("Requisitos_Minimos") = UCase(Txt_Cat_Cursos_Req_Minimos.Text)
            .rdoColumns("Lista_Material") = UCase(Txt_Cat_Cursos_Lista_Material.Text)
            .rdoColumns("Temario") = UCase(Txt_Cat_Cursos_Temario.Text)
            Dim Auditable As String
            If Chk_Cat_Cursos_Auditable.Value = True Or Chk_Cat_Cursos_Auditable = 1 Then
            Auditable = "SI"
            Else
            Auditable = "NO"
            End If
            .rdoColumns("Auditable") = Auditable
            .rdoColumns("Estatus") = Cmb_Cat_Cursos_Estatus.Text
            .rdoColumns("Tipo_Curso_Id") = Format(Cmb_Cat_Cursos_Tipo_Curso.ItemData(Cmb_Cat_Cursos_Tipo_Curso.ListIndex), "00000")
            .rdoColumns("Usuario_Creo") = Nombre_Usuario
            .rdoColumns("Fecha_Creo") = Now
        .Update
    End With
    Rs_Alta_Cat_Cursos.Close
    Conexion_Base.CommitTrans
    MsgBox "Curso agregado", vbInformation
    Consulta_Cat_Cursos ""
    Exit Sub
'Si existe un error se hace rollback de la transacción y no se hacen cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub

Private Sub Limpiar_Formulario()
Txt_Cat_Cursos_Curso_Id.Text = ""
Txt_Cat_Cursos_Nombre.Text = ""
Txt_Cat_Cursos_Total_Horas.Text = ""
Txt_Cat_Cursos_Descriopcion.Text = ""
Txt_Cat_Cursos_Clave.Text = ""
Txt_Cat_Cursos_Objetivo.Text = ""
Txt_Cat_Cursos_Req_Minimos.Text = ""
Txt_Cat_Cursos_Lista_Material.Text = ""
Txt_Cat_Cursos_Temario.Text = ""
Chk_Cat_Cursos_Auditable.Value = 0
Chk_Cat_Cursos_Auditable.Value = False
Cmb_Cat_Cursos_Estatus.ListIndex = -1
Cmb_Cat_Cursos_Tipo_Curso.ListIndex = -1
End Sub

Private Sub Btn_Salir_Click()
If Btn_Salir.Caption = "Salir" Then
        Unload Me
Else
    Limpiar_Formulario
    Configurar_Formulario (False)
    Btn_Nuevo.Caption = "Nuevo"
    Btn_Modificar.Caption = "Modificar"
    Btn_Salir.Caption = "Salir"
    Fra_Cursos.Enabled = True
End If
    
End Sub

Private Sub Form_Load()
    Me.Height = 9700
    Me.Width = 7600
    Me.Top = 100
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Grid_Cat_Cursos_Click()
Dim Rs_Consulta_Cat_Cursos As rdoResultset
    If Grid_Cat_Cursos.Rows > 1 Then
        Txt_Cat_Cursos_Curso_Id.Text = Grid_Cat_Cursos.TextMatrix(Grid_Cat_Cursos.RowSel, 0)
        Mi_SQL = "SELECT * FROM Cat_Cursos_Capacitaciones"
        Mi_SQL = Mi_SQL & "  WHERE Curso_ID='" & Txt_Cat_Cursos_Curso_Id.Text & "'"
        Set Rs_Consulta_Cat_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        'Llena las cajas de texto
        If Not Rs_Consulta_Cat_Cursos.EOF Then
            With Rs_Consulta_Cat_Cursos
                Txt_Cat_Cursos_Curso_Id.Text = .rdoColumns("Curso_ID")
                Txt_Cat_Cursos_Nombre.Text = .rdoColumns("Nombre")
                
                If IsNull(.rdoColumns("Horas")) Then
                Txt_Cat_Cursos_Total_Horas.Text = ""
                Else
                Txt_Cat_Cursos_Total_Horas.Text = .rdoColumns("Horas")
               End If
                 If Not IsNull(.rdoColumns("Tipo_Curso_Id")) Then
'
                    Cmb_Cat_Cursos_Tipo_Curso.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Tipo_Curso_Id"), "Cat_Tipos_Cursos", "Nombre", "Tipo_Curso_Id"), Cmb_Cat_Cursos_Tipo_Curso)
'                    Cmb_Cat_Cursos_Tipo_Curso.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Conectar_Ayudante.Buscar_Nombre(.rdoColumns("Tipo_Curso_Id"), "Cat_Tipos_Cursos", "Nombre", "Tipo_Curso_Id"), Cmb_Cat_Cursos_Tipo_Curso)
                Else
                    Cmb_Cat_Cursos_Tipo_Curso.ListIndex = -1
                End If
                If IsNull(.rdoColumns("Descripcion")) Then
                Txt_Cat_Cursos_Descriopcion.Text = ""
                Else
                Txt_Cat_Cursos_Descriopcion.Text = .rdoColumns("Descripcion")
                End If
                If IsNull(.rdoColumns("Clave")) Then
                Txt_Cat_Cursos_Clave.Text = ""
                Else
                Txt_Cat_Cursos_Clave.Text = .rdoColumns("Clave")
                End If
                If IsNull(.rdoColumns("Objetivo")) Then
                Txt_Cat_Cursos_Objetivo.Text = ""
                Else
                Txt_Cat_Cursos_Objetivo.Text = .rdoColumns("Objetivo")
                End If
                If IsNull(.rdoColumns("Requisitos_Minimos")) Then
                Txt_Cat_Cursos_Req_Minimos.Text = ""
                Else
                Txt_Cat_Cursos_Req_Minimos.Text = .rdoColumns("Requisitos_Minimos")
                End If
                If IsNull(.rdoColumns("Lista_Material")) Then
                Txt_Cat_Cursos_Lista_Material.Text = ""
                Else
                Txt_Cat_Cursos_Lista_Material.Text = .rdoColumns("Lista_Material")
                End If
                If IsNull(.rdoColumns("Temario")) Then
                Txt_Cat_Cursos_Temario.Text = ""
                Else
                Txt_Cat_Cursos_Temario.Text = .rdoColumns("Temario")
                End If
                If IsNull(.rdoColumns("Auditable")) Then
                
                Chk_Cat_Cursos_Auditable.Value = 0
                Chk_Cat_Cursos_Auditable.Value = False
                Else
                
                Dim Auditable As String
                Auditable = .rdoColumns("Auditable")
                If Auditable = "SI" Then
                Chk_Cat_Cursos_Auditable.Value = 1
'                Chk_Cat_Cursos_Auditable.Value = True
                Else
                Chk_Cat_Cursos_Auditable.Value = 0
                Chk_Cat_Cursos_Auditable.Value = False
                End If
                End If
                If Not IsNull(.rdoColumns("Estatus")) Then
                    Cmb_Cat_Cursos_Estatus.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Estatus")), Cmb_Cat_Cursos_Estatus)
                Else
                    Cmb_Cat_Cursos_Estatus.ListIndex = -1
                End If
''            If Not IsNull(.rdoColumns("Tipo_Curso_Id")) Then
''                    Cmb_Cat_Cursos_Tipo_Curso.ListIndex = Conectar_Ayudante.Buscar_Cadena_Combo(Trim(.rdoColumns("Tipo_Curso_Id")), Cmb_Cat_Cursos_Tipo_Curso)
''                Else
''                    Cmb_Cat_Cursos_Tipo_Curso.ListIndex = -1
''                End If
            End With
        End If
        Rs_Consulta_Cat_Cursos.Close
    End If
End Sub
'*******************************************************************************
    'NOMBRE DE LA FUNCIÓN:  Modificar_Cat_Cursos
    'DESCRIPCIÓN:           Modifica el registro de la Institución
    'PARÁMETROS :
    'CREO       :           Ana Laura Huichapa Ramirez
    'FECHA_CREO        :    28 Diciembre 2015
    'MODIFICO          :
    'FECHA_MODIFICO    :
    'CAUSA_MODIFICACIÓN:
'*******************************************************************************
Private Sub Modificar_Cat_Cursos()
Dim Rs_Modificacion_Cat_Cursos As rdoResultset 'Informacion del registro
Dim Cont_Fila As Integer
On Error GoTo HANDLER
    Conexion_Base.BeginTrans
    'Consulta el Usuario actual seleccionado
    Mi_SQL = "SELECT * FROM Cat_Cursos_Capacitaciones"
    Mi_SQL = Mi_SQL & " WHERE Curso_ID ='" & Trim(Txt_Cat_Cursos_Curso_Id.Text) & "'"
    Set Rs_Modificacion_Cat_Cursos = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
    'Modifica los datos de la tabla Cat_Usuarios
    If (Validar_Componentes) Then
    With Rs_Modificacion_Cat_Cursos
        .Edit
            .rdoColumns("Nombre") = Trim(Txt_Cat_Cursos_Nombre.Text)
            .rdoColumns("Horas") = Trim(Txt_Cat_Cursos_Total_Horas.Text)
'            .rdoColumns("Tipo") = Trim(Txt_Cat_Cursos_Tipo.Text)
'            .rdoColumns("Instructor") = Format(Cmb_Cat_Cursos_Instructor.ItemData(Cmb_Cat_Cursos_Instructor.ListIndex), "00000")
            .rdoColumns("Descripcion") = Trim(Txt_Cat_Cursos_Descriopcion.Text)
'            .rdoColumns("Clave_SAP") = Trim(Txt_Cat_Cursos_Clave_SAP.Text)
            .rdoColumns("Clave") = Trim(Txt_Cat_Cursos_Clave.Text)
            .rdoColumns("Objetivo") = UCase(Txt_Cat_Cursos_Objetivo.Text)
            .rdoColumns("Requisitos_Minimos") = UCase(Txt_Cat_Cursos_Req_Minimos.Text)
            .rdoColumns("Lista_Material") = UCase(Txt_Cat_Cursos_Lista_Material.Text)
            .rdoColumns("Temario") = UCase(Txt_Cat_Cursos_Temario.Text)
            Dim Auditable As String
            If Chk_Cat_Cursos_Auditable.Value = True Or Chk_Cat_Cursos_Auditable = 1 Then
            Auditable = "SI"
            Else
            Auditable = "NO"
            End If
            .rdoColumns("Auditable") = Auditable
            .rdoColumns("Tipo_Curso_Id") = Format(Cmb_Cat_Cursos_Tipo_Curso.ItemData(Cmb_Cat_Cursos_Tipo_Curso.ListIndex), "00000")
            .rdoColumns("Estatus") = Trim(Cmb_Cat_Cursos_Estatus.Text)
            .rdoColumns("Usuario_Modifico") = Nombre_Usuario
            .rdoColumns("Fecha_Modifico") = Now
        .Update
        .Close
    End With
   End If
    Set Rs_Modificacion_Cat_Cursos = Nothing
    'Agrega los checadores
   
    Conexion_Base.CommitTrans
   MsgBox "El curso ha sido modificado", vbInformation + vbOKOnly, Me.Caption
   Consulta_Cat_Cursos ""
    
    Exit Sub
'Ante error hace el rollback en la transacción para no guardar los cambios en la base de datos
HANDLER:
    Conexion_Base.RollbackTrans
    For Each Er In rdoErrors
        MsgBox Er.Description
    Next Er
End Sub


