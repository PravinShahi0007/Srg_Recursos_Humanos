VERSION 5.00
Begin VB.Form Frm_Cat_Empleados_Credencial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   3345
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "Salir"
      Height          =   660
      Left            =   2010
      Picture         =   "Frm_Cat_Empleados_Credencial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4470
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   660
      Left            =   135
      Picture         =   "Frm_Cat_Empleados_Credencial.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "A"
      Top             =   4470
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.Label Lbl_Nombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre del Empleado"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   3360
      Width           =   3165
   End
   Begin VB.Label Lbl_Apellido 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apellido del Empleado"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   3750
      Width           =   3210
   End
   Begin VB.Image Img_Logo_Empresa 
      Height          =   1080
      Left            =   255
      Picture         =   "Frm_Cat_Empleados_Credencial.frx":0B14
      Stretch         =   -1  'True
      Top             =   465
      Width           =   2550
   End
   Begin VB.Image Img_Cat_Empleados_Foto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1440
      Left            =   915
      Picture         =   "Frm_Cat_Empleados_Credencial.frx":A9D31
      Stretch         =   -1  'True
      Top             =   1755
      Width           =   1305
   End
End
Attribute VB_Name = "Frm_Cat_Empleados_Credencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ruta As String

Private Sub Btn_Imprimir_Click()
Dim Mi_Impresora As Printer

On Error GoTo HANDLER
    'Imprime la forma
    MDIFrm_Apl_Principal.CommonDialog1.ShowPrinter
    Btn_Imprimir.Visible = False
    Btn_Salir.Visible = False
    PrintForm
'    'Comienza la impresion del encabezado
'    Printer.ScaleMode = vbCentimeters
'    'Printer.Orientation = vbHorizontal
'    Printer.FontSize = 10
'    Printer.Font = "Arial"
'    Printer.FontBold = True
'    Call Printer.PaintPicture(Img_Logo_Empresa.picture, 1, 0.5, 6, 2)
'    Printer.FontSize = 16
'    Printer.CurrentX = 8
'    Printer.CurrentY = 0.5
'    Printer.Print "GAPS DE EMPLEADOS"
'    Printer.CurrentX = 8
'    Printer.CurrentY = 1.2
'    Printer.Print Lbl_Nombre_Gap.Caption
'    Printer.CurrentX = 8
'    Printer.CurrentY = 1.9
'    Printer.Print Lbl_Comentarios_Gap.Caption
'    Printer.Line (0.5, 0.25)-(27.5, 21.25), , B
'    'Imagenes de empleados
'    Printer.FontSize = 10
'    Printer.FontBold = False
'    If Lbl_Empleado_1.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_1.picture, 1, 3, 4, 4)
'        Printer.CurrentX = 1
'        Printer.CurrentY = 7.3
'        Printer.Print Lbl_Empleado_1.Caption
'        Printer.CurrentX = 1
'        Printer.CurrentY = 7.6
'        Printer.Print Lbl_Puesto_1.Caption
'    End If
'    If Lbl_Empleado_2.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_2.picture, 10, 3, 4, 4)
'        Printer.CurrentX = 10
'        Printer.CurrentY = 7.3
'        Printer.Print Lbl_Empleado_2.Caption
'        Printer.CurrentX = 10
'        Printer.CurrentY = 7.6
'        Printer.Print Lbl_Puesto_2.Caption
'    End If
'    If Lbl_Empleado_3.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_3.picture, 19, 3, 4, 4)
'        Printer.CurrentX = 19
'        Printer.CurrentY = 7.3
'        Printer.Print Lbl_Empleado_3.Caption
'        Printer.CurrentX = 19
'        Printer.CurrentY = 7.6
'        Printer.Print Lbl_Puesto_3.Caption
'    End If
'    If Lbl_Empleado_4.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_4.picture, 1, 9, 4, 4)
'        Printer.CurrentX = 1
'        Printer.CurrentY = 13.3
'        Printer.Print Lbl_Empleado_4.Caption
'        Printer.CurrentX = 1
'        Printer.CurrentY = 13.6
'        Printer.Print Lbl_Puesto_4.Caption
'    End If
'    If Lbl_Empleado_5.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_5.picture, 10, 9, 4, 4)
'        Printer.CurrentX = 10
'        Printer.CurrentY = 13.3
'        Printer.Print Lbl_Empleado_5.Caption
'        Printer.CurrentX = 10
'        Printer.CurrentY = 13.6
'        Printer.Print Lbl_Puesto_5.Caption
'    End If
'    If Lbl_Empleado_6.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_6.picture, 19, 9, 4, 4)
'        Printer.CurrentX = 19
'        Printer.CurrentY = 13.3
'        Printer.Print Lbl_Empleado_6.Caption
'        Printer.CurrentX = 19
'        Printer.CurrentY = 13.6
'        Printer.Print Lbl_Puesto_6.Caption
'    End If
'    If Lbl_Empleado_7.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_7.picture, 1, 15, 4, 4)
'        Printer.CurrentX = 1
'        Printer.CurrentY = 19.3
'        Printer.Print Lbl_Empleado_7.Caption
'        Printer.CurrentX = 1
'        Printer.CurrentY = 19.6
'        Printer.Print Lbl_Puesto_7.Caption
'    End If
'    If Lbl_Empleado_8.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_8.picture, 10, 15, 4, 4)
'        Printer.CurrentX = 10
'        Printer.CurrentY = 19.3
'        Printer.Print Lbl_Empleado_8.Caption
'        Printer.CurrentX = 10
'        Printer.CurrentY = 19.6
'        Printer.Print Lbl_Puesto_8.Caption
'    End If
'    If Lbl_Empleado_9.Caption <> "" Then
'        Call Printer.PaintPicture(Img_Foto_9.picture, 19, 15, 4, 4)
'        Printer.CurrentX = 19
'        Printer.CurrentY = 19.3
'        Printer.Print Lbl_Empleado_9.Caption
'        Printer.CurrentX = 19
'        Printer.CurrentY = 19.6
'        Printer.Print Lbl_Puesto_9.Caption
'    End If
    Printer.EndDoc
    Btn_Imprimir.Visible = True
    Btn_Salir.Visible = True
    MsgBox "Credencia enviada a impresion", vbInformation, "Impresion de Gaps"
Exit Sub
HANDLER:
    MsgBox "Ha ocurrido un problema al enviar a impresión, verifique su impresora esté encendida y funcionando correctamte", vbExclamation
    Btn_Imprimir.Visible = True
    Btn_Salir.Visible = True
End Sub

Private Sub Btn_Salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 500
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
Public Sub Cargar_Credencial()
'    Ruta = "C:\Users\ahuichapa\Desktop\Proyectos Contel\Sistema RH\Proyecto\Logos_Empresas\images.JPG"
    Img_Logo_Empresa.picture = LoadPicture(Ruta)
End Sub

