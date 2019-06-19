VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Previo_Pdf 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Explorador de archivos"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Previo_Pdf.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1845
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Previo_Pdf.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Previo_Pdf.frx":0E54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Mandar a la Impresora el Documento"
            Object.Tag             =   "I"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar la Vista Previa"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Frm_Previo_Pdf.frx":13EE
      Begin VB.TextBox Txt_Titulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "VISTA PREVIA DE LA IMPRESION"
         Top             =   150
         Width           =   9615
      End
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   8295
      Left            =   135
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   10335
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.OLE OLE1 
      Height          =   8505
      Left            =   120
      OleObjectBlob   =   "Frm_Previo_Pdf.frx":1708
      SizeMode        =   1  'Stretch
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   10515
   End
   Begin VB.Label lblproyecto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1260
      TabIndex        =   0
      Top             =   630
      Width           =   435
   End
End
Attribute VB_Name = "Frm_Previo_Pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngFormWidth As Long  'Ancho de la formap
Private lngFormHeight As Long 'Largo de la forma

Private Sub Form_Load()
Dim Ctl As Control

    Me.Top = 0
    Me.Left = 0
    AcroPDF1.Height = Me.ScaleY(28, vbCentimeters, vbTwips)
    AcroPDF1.Width = Me.ScaleX(21.5, vbCentimeters, vbTwips)
    Me.Width = AcroPDF1.Width + 200
    Me.Height = Screen.Height - 2200
    AcroPDF1.Height = Me.Height - 1200
    AcroPDF1.setShowToolbar False
    'dimensiones en variables
    'lngFormWidth = 10695
    'lngFormHeight = 8265
    '
    'On Error Resume Next
    '
    'For Each Ctl In Me
    '   Ctl.Tag = Ctl.Left & " " & Ctl.Top & " " & Ctl.Width & " " & Ctl.Height & " "
    '   Ctl.Tag = Ctl.Tag & Ctl.Font.Size & " "
    'Next Ctl
    Toolbar1.Style = tbrFlat
End Sub

Private Sub Form_Resize()
'Dim D(4), ScaleX, ScaleY As Double
'Dim i, TempPoz, StartPoz As Long
'Dim Ctl As Control
'Dim TempVisible As Boolean
'
''Calcula la escala
'ScaleX = ScaleWidth / lngFormWidth
'ScaleY = ScaleHeight / lngFormHeight
'
'On Error Resume Next
'For Each Ctl In Me
'    StartPoz = 1
'    For i = 0 To 4
'        TempPoz = InStr(StartPoz, Ctl.Tag, " ", vbTextCompare)
'        If TempPoz > 0 Then
'           D(i) = Mid(Ctl.Tag, StartPoz, TempPoz - StartPoz)
'           StartPoz = TempPoz + 1
'        Else
'          D(i) = 0
'        End If
'        Ctl.Width = D(2) * ScaleX
'        Ctl.Height = D(3) * ScaleY
'        Ctl.Left = D(0) * ScaleX
'        Ctl.Top = D(1) * ScaleY
'        If ScaleX < ScaleY Then
'            Ctl.Font.Size = D(4) * ScaleX
'        Else
'            Ctl.Font.Size = D(4) * ScaleY
'        End If
'    Next i
'Next Ctl
'
'On Error GoTo 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        If OLE1.Visible = True Then
            OLE1.DoVerb (-1)
        End If
        AcroPDF1.printWithDialog
    Case 3
        Unload Me
    End Select
End Sub
