VERSION 5.00
Begin VB.Form Frm_Pro_Descarga_Informacion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DESCARGA DE INFORMACION"
   ClientHeight    =   2985
   ClientLeft      =   6045
   ClientTop       =   4200
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   10485
   Begin VB.CommandButton Btn_Muestreo 
      Caption         =   "MUESTREO CALIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   7290
      Picture         =   "Frm_Pro_Descarga_Informacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   735
      Width           =   3030
   End
   Begin VB.Timer Tmr_Frecuencia_Captura 
      Interval        =   1000
      Left            =   3390
      Top             =   2400
   End
   Begin VB.ComboBox Cmb_Turno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "Frm_Pro_Descarga_Informacion.frx":030A
      Left            =   315
      List            =   "Frm_Pro_Descarga_Informacion.frx":030C
      TabIndex        =   0
      Top             =   90
      Width           =   10035
   End
   Begin VB.CommandButton Btn_Captura_de_produccion 
      Caption         =   "CAPTURA DE PRODUCCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   315
      Picture         =   "Frm_Pro_Descarga_Informacion.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   735
      Width           =   3030
   End
   Begin VB.CommandButton Btn_Tiempo_Muerto 
      Caption         =   "TIEMPO MUERTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   3885
      Picture         =   "Frm_Pro_Descarga_Informacion.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   735
      Width           =   3030
   End
End
Attribute VB_Name = "Frm_Pro_Descarga_Informacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Frecuencia_Captura As Double
Dim Cont_Timer As Double

Private Sub Btn_Captura_de_produccion_Click()
    If Cmb_Turno.ListIndex > -1 Then
        Catalogo = "PRODUCCION"
        Unload Frm_Pro_Captura_Produccion
        Load Frm_Pro_Captura_Produccion
        Frm_Pro_Captura_Produccion.Caption = "CAPTURA DE PRODUCCION"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Pro_Captura_Produccion.Pic_Captura_Produccion, Frm_Pro_Captura_Produccion)
        Frm_Pro_Captura_Produccion.Busca_Turno Format(Cmb_Turno.ItemData(Cmb_Turno.ListIndex), "0000000000")
        Frm_Pro_Captura_Produccion.Txt_Total.SetFocus
        'Frm_Pro_Captura_Produccion.Txt_Codigo_Barras.SetFocus
        If Encontro_Turno = False Then
            Unload Frm_Pro_Captura_Produccion
            Encontro_Turno = False
        Else
            Unload Me
        End If
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Busca_Turno
'DESCRIPCIÓN: Busca los turnos del usuario registrado
'CREO: Sergio Ulises Durán Hernández
'FECHA_CREO: 17-Febrero-2009
'MODIFICO:
'FECHA_MODIFICO:
'CAUSA_MODIFICACIÓN:
'PARÁMETROS:
'*******************************************************************************
Public Sub Busca_Turno()
Dim Rs_Pro_Captura_Turno As rdoResultset
    
    Mi_SQL = "SELECT Captura_Turno_ID,Orden_Produccion,Tipo_Orden,Frecuencia_Captura"
    Mi_SQL = Mi_SQL & " FROM Pro_Captura_Turno"
    Mi_SQL = Mi_SQL & " WHERE Pro_Captura_Turno.Estatus='PENDIENTE'"
    If Rol = "PRODUCCION OPERACION" Then
        Mi_SQL = Mi_SQL & " AND Supervisor_Turno_ID='" & Usuario_ID & "'"
    Else
        If Rol <> "ADMINISTRADOR" And Rol <> "PRODUCCION" Then   'Impide a los demás usuarios ver los turnos
            Mi_SQL = Mi_SQL & " AND Pro_Captura_Turno.Captura_Turno_ID IS NULL"
        End If
    End If
    Mi_SQL = Mi_SQL & " AND (Fecha='" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " OR Fecha_Reprogramacion='" & Format(Now, "MM/dd/yyyy") & "'"
    Mi_SQL = Mi_SQL & " AND Pro_Captura_Turno.Estatus='REPROGRAMADO')"
    Set Rs_Pro_Captura_Turno = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    Cmb_Turno.Text = ""
    Cmb_Turno.Clear
    If Not Rs_Pro_Captura_Turno.EOF Then
        While Not Rs_Pro_Captura_Turno.EOF
            If Rs_Pro_Captura_Turno.rdoColumns("Tipo_Orden") = "PRODUCCION" Then
                Cmb_Turno.AddItem "TURNO " & Rs_Pro_Captura_Turno.rdoColumns("Captura_Turno_ID") & " - O.P. " & Rs_Pro_Captura_Turno.rdoColumns("Orden_Produccion")
            Else
                Cmb_Turno.AddItem "TURNO " & Rs_Pro_Captura_Turno.rdoColumns("Captura_Turno_ID") & " - O.A. " & Rs_Pro_Captura_Turno.rdoColumns("Orden_Produccion")
            End If
            Cmb_Turno.ItemData(Cmb_Turno.NewIndex) = Rs_Pro_Captura_Turno.rdoColumns("Captura_Turno_ID")
            Rs_Pro_Captura_Turno.MoveNext
        Wend
        Cmb_Turno.ListIndex = 0
    End If
    Rs_Pro_Captura_Turno.Close
End Sub

Private Sub Btn_Muestreo_Click()
    If Cmb_Turno.ListIndex > -1 Then
        Catalogo = "CALIDAD"
        Unload Frm_Pro_Captura_Produccion
        Load Frm_Pro_Captura_Produccion
        Frm_Pro_Captura_Produccion.Caption = "CAPTURA DE TOMA DE MUESTRAS"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Pro_Captura_Produccion.Pic_Captura_Produccion, Frm_Pro_Captura_Produccion)
        Frm_Pro_Captura_Produccion.Busca_Turno Format(Cmb_Turno.ItemData(Cmb_Turno.ListIndex), "0000000000")
        Frm_Pro_Captura_Produccion.Txt_Total.SetFocus
        'Frm_Pro_Captura_Produccion.Txt_Codigo_Barras.SetFocus
        If Encontro_Turno = False Then
            Unload Frm_Pro_Captura_Produccion
            Encontro_Turno = False
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub Btn_Tiempo_Muerto_Click()
    If Cmb_Turno.ListIndex > -1 Then
        Catalogo = "TIEMPO"
        Unload Frm_Pro_Captura_Produccion
        Load Frm_Pro_Captura_Produccion
        Frm_Pro_Captura_Produccion.Caption = "TIEMPO MUERTO"
        Call Conectar_Ayudante.Cargar_Picture(Frm_Pro_Captura_Produccion.Pic_Tiempo_Muerto, Frm_Pro_Captura_Produccion)
        Frm_Pro_Captura_Produccion.Busca_Turno Format(Cmb_Turno.ItemData(Cmb_Turno.ListIndex), "0000000000")
        Frm_Pro_Captura_Produccion.Cmb_Tipo_Pago.SetFocus
        If Encontro_Turno = False Then
            Unload Frm_Pro_Captura_Produccion
            Encontro_Turno = False
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub Cmb_Turno_Click()
    Cont_Timer = 0
End Sub


Private Sub Form_Load()
    'Medidas de la Forma
    Me.Width = 10700
    Me.Height = 3500
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Busca_Turno
    Btn_Captura_de_produccion.Enabled = True
    Btn_Tiempo_Muerto.Enabled = True
    Btn_Muestreo.Enabled = True
    If Rol <> "ADMINISTRADOR" Then
        If Mid(Rol, 1, 10) <> "CALIDAD" Then
            Btn_Muestreo.Enabled = False
        Else
            Btn_Captura_de_produccion.Enabled = False
            Btn_Tiempo_Muerto.Enabled = False
        End If
    End If
End Sub

Private Sub Tmr_Frecuencia_Captura_Timer()
Dim Rs_Frecuencia_Captura As rdoResultset
Dim Mi_SQL As String
    If Cmb_Turno.ListIndex > -1 Then
        If Cont_Timer = 0 Then
            Mi_SQL = " SELECT Frecuencia_Captura FROM Pro_Captura_Turno"
            Mi_SQL = Mi_SQL & " WHERE Captura_Turno_ID='" & Format(Cmb_Turno.ItemData(Cmb_Turno.ListIndex), "0000000000") & "'"
            Set Rs_Frecuencia_Captura = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            If Not Rs_Frecuencia_Captura.EOF Then
               If Not IsNull(Rs_Frecuencia_Captura.rdoColumns("Frecuencia_Captura")) Then Frecuencia_Captura = Val(Rs_Frecuencia_Captura.rdoColumns("Frecuencia_Captura"))
                Frecuencia_Captura = (Frecuencia_Captura * 60) * 1000
            End If
        End If
        If Cont_Timer = Frecuencia_Captura Then
            Cont_Timer = 0
            Btn_Captura_de_produccion_Click
        End If
        Cont_Timer = Cont_Timer + 1000
    End If
End Sub

