VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Ope_Registro_Calificaciones 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE CALIFICACIONES"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8145
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_Cursos 
      BackColor       =   &H8000000E&
      Caption         =   "Busqueda de Cursos"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton Btn_Buscar 
         Caption         =   "Buscar"
         Height          =   435
         Left            =   6600
         Picture         =   "Frm_Ope_Registro_Calificaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "C"
         Top             =   1560
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Cursos 
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker Dtp_Fecha_Fin 
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122093569
         CurrentDate     =   42374
      End
      Begin MSComCtl2.DTPicker Dtp_Fecha_Inicio 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122093569
         CurrentDate     =   42374
      End
      Begin VB.ComboBox Cmb_Cursos 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   6500
      End
      Begin VB.ComboBox Cmb_Instructor 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   6500
      End
      Begin VB.ComboBox Cmb_Institucion 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   6500
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Curso"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Instructor"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Institucion"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Fra_Registrar_Calificaciones 
      BackColor       =   &H8000000E&
      Caption         =   "Registrar Calificaciones"
      Height          =   6015
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton Btn_Guardar 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   6960
         TabIndex        =   31
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Btn_Regresar 
         Caption         =   "Regresar"
         Height          =   495
         Left            =   6960
         TabIndex        =   30
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton Btn_Subir_Archivo 
         Caption         =   "Subir Archivo"
         Height          =   495
         Left            =   6960
         TabIndex        =   29
         Top             =   3240
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Calificaciones 
         Height          =   3615
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Appearance      =   0
      End
      Begin VB.TextBox Txt_Fecha_Fin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Txt_Fecha_Inicio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Txt_Curso 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox Txt_Instructor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1080
         Width           =   6135
      End
      Begin VB.TextBox Txt_Institucion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   10
         Top             =   720
         Width           =   6135
      End
      Begin VB.TextBox Txt_estatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Txt_No_Curso 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         Caption         =   "No Curso"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Curso"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Estatus"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Instructor"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Institucion"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Frm_Ope_Registro_Calificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicializa()
    Dtp_Fecha_Inicio.Value = DateAdd("d", -1, Now)
    Dtp_Fecha_Fin.Value = DateAdd("d", 0, Now)
    'Carga las empresas localidades
    Call Conectar_Ayudante.Llena_Combo_Item("Institucion_ID, Nombre", "Cat_Instituciones WHERE Estatus='ACTIVO'", Cmb_Institucion, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Instructor_ID, Nombre", "Cat_Instructores WHERE Estatus='ACTIVO'", Cmb_Instructor, 0, "Nombre")
    Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID, Nombre", "Cat_Cursos_Capacitaciones WHERE Estatus='ACTIVO'", Cmb_Cursos, 0, "Nombre")
    Grid_Cursos.Cols = 0
    Grid_Cursos.Rows = 0
End Sub

Private Sub Btn_Buscar_Click()
Dim Rs_Consulta_Cursos As rdoResultset

    Grid_Cursos.Cols = 7
    Grid_Cursos.Rows = 0
    Mi_SQL = "SELECT Ope_Programacion_Cursos.No_Programa_Curso,Cat_Cursos_Capacitaciones.Nombre as Nombre_Curso, Cat_Instituciones.Clave as Institucion,"
    Mi_SQL = Mi_SQL & " Cat_Instructores.Nombre + ' ' + Cat_Instructores.Apellido_Paterno + ' ' + Cat_Instructores.Apellido_Materno as Instructor,"
    Mi_SQL = Mi_SQL & " Fecha_Inicio,Hora_inicio,Fecha_Fin,Hora_Fin,"
    Mi_SQL = Mi_SQL & " Ope_Programacion_Cursos.Estatus"
    Mi_SQL = Mi_SQL & " FROM Ope_Programacion_Cursos, Cat_Cursos_Capacitaciones, Cat_Instituciones, Cat_Instructores"
    Mi_SQL = Mi_SQL & " WHERE Ope_Programacion_Cursos.Curso_ID = Cat_Cursos_Capacitaciones.Curso_ID"
    Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Instructor_Id=Cat_Instructores.Instructor_Id"
    Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Institucion_Id=Cat_Instituciones.Institucion_Id"
    Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Estatus='ACTIVO'"
    If Cmb_Institucion.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Institucion_Id= '" & Format(Cmb_Institucion.ItemData(Cmb_Institucion.ListIndex), "00000") & "'"
    End If
    If Cmb_Instructor.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Instructor_Id = '" & Format(Cmb_Instructor.ItemData(Cmb_Instructor.ListIndex), "00000") & "'"
    End If
    If Cmb_Cursos.ListIndex > 0 Then
        Mi_SQL = Mi_SQL & " AND Ope_Programacion_Cursos.Curso_ID= '" & Format(Cmb_Cursos.ItemData(Cmb_Cursos.ListIndex), "00000") & "'"
    End If
    If (Dtp_Fecha_Inicio.Value <> Dtp_Fecha_Fin.Value) Then
        Mi_SQL = Mi_SQL & " AND Fecha_Inicio>='" & Format(Dtp_Fecha_Inicio.Value, "yyyy/MM/dd") & "' AND Fecha_Fin<='" & Format(Dtp_Fecha_Fin.Value, "yyyy/MM/dd") & "'"
    End If
    Set Rs_Consulta_Cursos = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Cursos.EOF Then
        Grid_Cursos.AddItem "no." & Chr(9) & "Curso" _
                    & Chr(9) & "Institucion" & Chr(9) & "Instructor" _
                    & Chr(9) & "Fecha_Inicio" & Chr(9) & "Fecha_Fin" _
                    & Chr(9) & "Estatus"
        Rs_Consulta_Cursos.MoveFirst
        While Not Rs_Consulta_Cursos.EOF
            Grid_Cursos.AddItem Rs_Consulta_Cursos.rdoColumns("No_Programa_Curso") & Chr(9) & Rs_Consulta_Cursos.rdoColumns("nombre_curso") _
                        & Chr(9) & Rs_Consulta_Cursos.rdoColumns("institucion") & Chr(9) & Rs_Consulta_Cursos.rdoColumns("instructor") _
                        & Chr(9) & Rs_Consulta_Cursos.rdoColumns("fecha_inicio") & " " & Format(Rs_Consulta_Cursos.rdoColumns("hora_inicio"), "hh:mm:ss") _
                        & Chr(9) & Rs_Consulta_Cursos.rdoColumns("fecha_fin") & " " & Format(Rs_Consulta_Cursos.rdoColumns("hora_fin"), "hh:mm:ss") _
                        & Chr(9) & Rs_Consulta_Cursos.rdoColumns("estatus")
            Rs_Consulta_Cursos.MoveNext
        Wend
        
        With Grid_Cursos
            If .Rows > 1 Then .FixedRows = 1
                .FixedCols = 0
                .ColWidth(0) = 0         'no
                .ColWidth(1) = 2000         'Curso
                .ColWidth(2) = 2000         'Institucion
                .ColWidth(3) = 2000         'Instructor
                .ColWidth(4) = 1700         'Fecha_Inicio
                .ColWidth(5) = 1700         'fecha_Fin
                .ColWidth(6) = 1500         'Estatus
        End With
    Else
        MsgBox "No se encontraron resultados", vbInformation, "Mensaje"
    End If
End Sub

Private Sub Btn_Guardar_Click()
Dim Cont As Integer
Dim Bandera As Boolean
Bandera = False

On Error GoTo Fin
    For Cont = 0 To Grid_Calificaciones.Rows - 1
        If (Grid_Calificaciones.TextMatrix(Cont, 2) = "") Then
            Bandera = True
        End If
    Next Cont
    
    If Bandera = True Then ' Si faltan empleados por evaluar PREGUNTA
        If MsgBox("Aun existen empleados sin cargar evaluacion. ¿Desea Continuar?", vbYesNo, "Mensaje") = vbYes Then
            'si desea guardar aun con los faltantes
            Conexion_Base.BeginTrans
            For Cont = 1 To Grid_Calificaciones.Rows - 1
                Mi_SQL = "INSERT INTO Ope_Evaluaciones_Empleados (No_Programa_Curso,Empleado_Id,Archivo,Fecha,Usuario_Creo,Fecha_Creo) "
                Mi_SQL = Mi_SQL & "VALUES('" & Txt_No_Curso.Text & "','" & Grid_Calificaciones.TextMatrix(Cont, 0)
                Mi_SQL = Mi_SQL & "','" & Grid_Calificaciones.TextMatrix(Cont, 2) & "','" & Format(Now, "MM/dd/yyyy")
                Mi_SQL = Mi_SQL & "','" & Usuario & "',GETDATE())"
                Conexion_Base.Execute Mi_SQL
            Next Cont
            Mi_SQL = ""
            Mi_SQL = "UPDATE Ope_Programacion_Cursos "
            Mi_SQL = Mi_SQL & "SET Estatus='CERRADO' "
            Mi_SQL = Mi_SQL & "WHERE No_Programa_Curso = '" & Txt_No_Curso.Text & "'"
            Conexion_Base.Execute Mi_SQL
            
            Conexion_Base.CommitTrans
            MsgBox "La operacion se realizo satisfactoriamente.", vbInformation, "Mensaje"
            Btn_Regresar_Click
            Grid_Cursos.Rows = 0
            Grid_Cursos.Cols = 0
        End If
    Else 'si la evaluacion esta completa guarda
        Conexion_Base.BeginTrans
        For Cont = 1 To Grid_Calificaciones.Rows - 1
            Mi_SQL = "INSERT INTO Ope_Evaluaciones_Empleados (No_Programa_Curso,Empleado_Id,Archivo,Fecha,Usuario_Creo,Fecha_Creo) "
            Mi_SQL = Mi_SQL & "VALUES('" & Txt_No_Curso.Text & "','" & Grid_Calificaciones.TextMatrix(Cont, 0)
            Mi_SQL = Mi_SQL & "','" & Grid_Calificaciones.TextMatrix(Cont, 2) & "','" & Format(Now, "MM/dd/yyyy")
            Mi_SQL = Mi_SQL & "','" & Usuario & "',GETDATE())"
            Conexion_Base.Execute Mi_SQL
        Next Cont
        Mi_SQL = ""
        Mi_SQL = "UPDATE Ope_Programacion_Cursos "
        Mi_SQL = Mi_SQL & "SET Estatus='CERRADO' "
        Mi_SQL = Mi_SQL & "WHERE No_Programa_Curso = '" & Txt_No_Curso.Text & "'"
        Conexion_Base.Execute Mi_SQL
        
        Conexion_Base.CommitTrans
        MsgBox "La operacion se realizo satisfactoriamente.", vbInformation, "Mensaje"
        Btn_Regresar_Click
        Grid_Cursos.Rows = 0
        Grid_Cursos.Cols = 0
    End If
Exit Sub
Fin:
    If Err.Number <> 0 Then
        Conexion_Base.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If

End Sub

Private Sub Btn_Regresar_Click()
    Grid_Calificaciones.Rows = 0
    Grid_Cursos.Rows = 0
    Txt_No_Curso.Text = ""
    Txt_estatus.Text = ""
    Txt_Fecha_Fin.Text = ""
    Txt_Fecha_Inicio.Text = ""
    Txt_Institucion.Text = ""
    Txt_Instructor.Text = ""
    Txt_No_Curso.Text = ""
    Fra_Registrar_Calificaciones.Visible = False
    Fra_Cursos.Visible = True
End Sub

Private Sub Btn_Subir_Archivo_Click()
On Error GoTo Fin
    'Set CancelError is True
    CommonDialog1.CancelError = True
    'Titulo de la ventana
    CommonDialog1.DialogTitle = "Seleccione el Archivo de Evaluación"
    'Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    'Set filters
    CommonDialog1.Filter = "Archivo PDF (*.pdf)|*.pdf"
    'Specify default filter
    CommonDialog1.FilterIndex = 1
    'Display the Open dialog box
    CommonDialog1.ShowOpen
    'Display name of selected file
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(CommonDialog1.FileName, "ARCHIVO") = True Then
        Grid_Calificaciones.TextMatrix(Grid_Calificaciones.RowSel, 2) = Txt_No_Curso.Text & "_" & Grid_Calificaciones.TextMatrix(Grid_Calificaciones.RowSel, 0) & ".pdf"
        
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Evaluaciones_Empleados", "CARPETA") = True Then
            Call FileCopy(CommonDialog1.FileName, App.Path & "\Evaluaciones_Empleados\" & Txt_No_Curso.Text & "_" & Grid_Calificaciones.TextMatrix(Grid_Calificaciones.RowSel, 0) & ".pdf")
        Else
            Call MkDir(App.Path & "\Evaluaciones_Empleados")
            Call FileCopy(CommonDialog1.FileName, App.Path & "\Evaluaciones_Empleados\" & Txt_No_Curso.Text & "_" & Grid_Calificaciones.TextMatrix(Grid_Calificaciones.RowSel, 0) & ".pdf")
        End If
        MsgBox "El archivo se cargo satisfactoriamente.", vbInformation, "Mensaje"
    End If
Exit Sub
Fin:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If

End Sub

Private Sub Cmb_Cursos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Curso_ID, Nombre", "Cat_Cursos_Capacitaciones", Cmb_Cursos, 1, "Nombre", " AND Estatus='ACTIVO'")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Institucion_Click()
Dim Rs_Cat_Instructor As rdoResultset     'Informcion de los empleados
    If Cmb_Institucion.ListIndex > -1 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Instructor_ID, Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno", "Cat_Instructores WHERE Institucion_ID='" & Format(Cmb_Institucion.ItemData(Cmb_Institucion.ListIndex), "00000") & "'", Cmb_Instructor, 0, "Nombre", , False, "TODAS")
    End If
End Sub

Private Sub Cmb_Institucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Institucion_ID, Nombre", "Cat_Instituciones", Cmb_Institucion, 1, "Nombre", " AND Estatus='ACTIVO'")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Institucion_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Institucion, KeyCode)
End Sub


Private Sub Cmb_Instructor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Conectar_Ayudante.Llena_Combo_Item("Instructor_ID, Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno", "Cat_Instructores", Cmb_Instructor, 1, "Nombre", " AND Estatus='ACTIVO'")
    Else
        Call Conectar_Ayudante.Quitar_Caracter_Raro(KeyAscii)
    End If
End Sub

Private Sub Cmb_Instructor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Conectar_Ayudante.Buscar_List_Combo(Cmb_Instructor, KeyCode)
End Sub



Private Sub Grid_Cursos_DblClick()
Dim Rs_Empleados As rdoResultset

    Fra_Cursos.Visible = False
    Fra_Registrar_Calificaciones.Visible = True
    Txt_No_Curso.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 0)
    Txt_Curso.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 1)
    Txt_Institucion.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 2)
    Txt_Instructor.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 3)
    Txt_Fecha_Inicio.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 4)
    Txt_Fecha_Fin.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 5)
    Txt_estatus.Text = Grid_Cursos.TextMatrix(Grid_Cursos.RowSel, 6)
    
    Mi_SQL = "SELECT Empleado_ID,(Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno) as Nombre"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados WHERE Empleado_ID IN"
    Mi_SQL = Mi_SQL & " (SELECT Empleado_ID FROM Ope_Lista_Asistencia "
    Mi_SQL = Mi_SQL & " WHERE No_Programa_Curso='" & Txt_No_Curso.Text & "')"
    Set Rs_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Empleados.EOF Then
        Grid_Calificaciones.Rows = 0
        Grid_Calificaciones.Cols = 3
        
        Grid_Calificaciones.AddItem "Empleado_ID" & Chr(9) & "Empleado" & Chr(9) & "Archivo"
        Rs_Empleados.MoveFirst
        While Not Rs_Empleados.EOF
            Grid_Calificaciones.AddItem Rs_Empleados.rdoColumns("Empleado_ID") & Chr(9) & Rs_Empleados.rdoColumns("Nombre") & Chr(9) & ""
            Rs_Empleados.MoveNext
        Wend
        
        With Grid_Calificaciones
            If .Rows > 1 Then .FixedRows = 1
                .FixedCols = 0
                .ColWidth(0) = 0         'Empleado_ID
                .ColWidth(1) = 3000         'Nombre
                .ColWidth(2) = 5000         'Archivo
                .ColAlignment(2) = 0
        End With
    End If
End Sub

