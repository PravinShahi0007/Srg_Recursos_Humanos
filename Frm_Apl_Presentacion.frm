VERSION 5.00
Begin VB.Form Frm_Apl_Presentacion 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4350
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Fra_General 
      BackColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      Begin VB.Image Img_Logo_Empresa 
         Height          =   1935
         Left            =   0
         Picture         =   "Frm_Apl_Presentacion.frx":0000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Image Img_Logo_Contel 
         Height          =   885
         Left            =   4695
         Picture         =   "Frm_Apl_Presentacion.frx":A921D
         Top             =   1575
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SRG Global M�xico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   105
         TabIndex        =   13
         Top             =   300
         Width           =   3360
      End
      Begin VB.Label Lbl_Warning 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_Apl_Presentacion.frx":AA739
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   6960
      End
      Begin VB.Label Lbl_Nombre_Proyecto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE DEL PROYECTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3915
      End
      Begin VB.Label Lbl_Plataforma 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Visual Basic 6.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label Lbl_Version 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Versi�n "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   9
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Lbl_Compa�ia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONECTIVIDAD Y TELECOMUNICACI�N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   8
         Top             =   2625
         Width           =   3765
      End
      Begin VB.Label Lbl_Copyright 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   7
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Lbl_Product_Licenced 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This product is licenced to:"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2025
      End
      Begin VB.Label Lbl_Plataform 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Platform:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3150
         TabIndex        =   5
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Lbl_A�o 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2014"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   4
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label Lbl_C 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "�"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3840
         TabIndex        =   3
         Top             =   1875
         Width           =   225
      End
      Begin VB.Label Lbl_Departamento_Desarrollo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento de Desarrollo de Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   2
         Top             =   2925
         Width           =   3375
      End
      Begin VB.Label Lbl_Designed_By 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Designed By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3120
         TabIndex        =   1
         Top             =   2400
         Width           =   915
      End
   End
End
Attribute VB_Name = "Frm_Apl_Presentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    lbl_Version.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lbl_Nombre_Proyecto.Caption = App.Title
    Me.Width = 7392
    Me.Height = 4428
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    If Ciclos = 6 Then
        Unload Frm_Apl_Presentacion
        Load MDIFrm_Apl_Principal
        Screen.MousePointer = 0
    Else
        Ciclos = Ciclos + 1
    End If
End Sub

